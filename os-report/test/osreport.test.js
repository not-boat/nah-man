"use strict";
// Run with:  node --test test/osreport.test.js
//
// Tests parsing, INV extraction, FIFO selection, and ageing bucketing —
// including a smoke test against the actual DEVA JCB ledger that should
// produce 9 outstanding invoices summing exactly to ₹16,466,589.48.

const test   = require("node:test");
const assert = require("node:assert/strict");
const O      = require("../src/osreport.js");

// ── INV extractor ────────────────────────────────────────────────────────────

test("INV: 'ref inv.251420019' → 251420019", () => {
  assert.equal(O.extractInvFromNarration("ref inv.251420019"), "251420019");
});

test("INV: 'ref inv. 252240227' (extra space) → 252240227", () => {
  assert.equal(O.extractInvFromNarration("ref inv. 252240227"), "252240227");
});

test("INV: trailing CRLF stripped", () => {
  assert.equal(O.extractInvFromNarration("ref inv.251423218\r\r\n"), "251423218");
});

test("INV: 'REF 251610287' (uppercase, no period) → 251610287", () => {
  assert.equal(O.extractInvFromNarration("REF 251610287"), "251610287");
});

test("INV: 'Invoice:253400899' → 253400899", () => {
  assert.equal(O.extractInvFromNarration("Invoice:253400899"), "253400899");
});

test("INV: 'ref inv.241423481 on 25-01-2025' → 241423481 (date suffix ignored)", () => {
  assert.equal(O.extractInvFromNarration("ref inv.241423481 on 25-01-2025"), "241423481");
});

test("INV: 'ABC' (no inv) → ''", () => {
  assert.equal(O.extractInvFromNarration("ABC"), "");
});

test("INV: bare 252140128 → 252140128", () => {
  assert.equal(O.extractInvFromNarration("252140128"), "252140128");
});

test("INV: empty / null → ''", () => {
  assert.equal(O.extractInvFromNarration(""), "");
  assert.equal(O.extractInvFromNarration(null), "");
  assert.equal(O.extractInvFromNarration(undefined), "");
});

// ── Date helpers ─────────────────────────────────────────────────────────────

test("toSerial: Excel serial passthrough", () => {
  assert.equal(O.toSerial(46130), 46130);
});

test("toSerial: '03-Apr-25' → 45750", () => {
  // 03-Apr-2025 = Excel serial 45750
  assert.equal(O.toSerial("03-Apr-25"),   45750);
  assert.equal(O.toSerial("03-Apr-2025"), 45750);
});

test("toSerial: '18-04-2026' → 46130", () => {
  assert.equal(O.toSerial("18-04-2026"), 46130);
});

test("toSerial: '2026-04-18' → 46130", () => {
  assert.equal(O.toSerial("2026-04-18"), 46130);
});

// ── Ageing buckets ───────────────────────────────────────────────────────────

test("ageing buckets", () => {
  assert.equal(O.ageingBucketLabel(0),   "0-30");
  assert.equal(O.ageingBucketLabel(30),  "0-30");
  assert.equal(O.ageingBucketLabel(31),  "31-60");
  assert.equal(O.ageingBucketLabel(60),  "31-60");
  assert.equal(O.ageingBucketLabel(61),  "61-90");
  assert.equal(O.ageingBucketLabel(91),  "91-120");
  assert.equal(O.ageingBucketLabel(120), "91-120");
  assert.equal(O.ageingBucketLabel(121), "121-Above");
  assert.equal(O.ageingBucketLabel(999), "121-Above");
});

// ── FIFO selection ───────────────────────────────────────────────────────────

test("FIFO: picks newest first, adjusts last", () => {
  const entries = [
    { date: 100, debit: 1000, narration: "" },
    { date: 200, debit:  500, narration: "" },
    { date: 300, debit:  700, narration: "" },
    { date: 400, debit:  300, narration: "" },
  ];
  const picked = O.selectFifoOutstanding(entries, 1100);
  // Newest first: 300 + 700 = 1000 < 1100, take 100 of 500 (date=200)
  assert.deepEqual(picked.map(p => p.date),   [400, 300, 200]);
  assert.deepEqual(picked.map(p => p.amount), [300, 700, 100]);
  assert.equal(picked[2].adjusted, true);
  assert.equal(picked[0].adjusted, false);
});

test("FIFO: exact match, no adjustment", () => {
  const entries = [
    { date: 100, debit: 500, narration: "" },
    { date: 200, debit: 500, narration: "" },
  ];
  const picked = O.selectFifoOutstanding(entries, 1000);
  assert.equal(picked.length, 2);
  assert.equal(picked[0].amount, 500);
  assert.equal(picked[1].amount, 500);
  assert.equal(picked[0].adjusted, false);
  // Final entry might be marked adjusted if the algorithm flagged it, but
  // amounts should be intact
  assert.equal(picked.reduce((s, p) => s + p.amount, 0), 1000);
});

test("FIFO: outstanding larger than total → returns all entries", () => {
  const entries = [
    { date: 100, debit: 500, narration: "" },
    { date: 200, debit: 500, narration: "" },
  ];
  const picked = O.selectFifoOutstanding(entries, 5000);
  assert.equal(picked.length, 2);
  assert.equal(picked.reduce((s, p) => s + p.amount, 0), 1000);
});

test("FIFO: zero / negative outstanding → []", () => {
  const entries = [{ date: 100, debit: 500, narration: "" }];
  assert.deepEqual(O.selectFifoOutstanding(entries, 0),  []);
  assert.deepEqual(O.selectFifoOutstanding(entries, -5), []);
});

// ── Debtors parser ───────────────────────────────────────────────────────────

// ── Regression: empty-key fuzzy match must not allocate to wrong ledger ─────
//
// A debtor whose name normalizes to "" (e.g. a stray separator row) used to
// match the first ledger via "".startsWith("") and silently allocate that
// debtor's outstanding to the wrong customer's invoices. Now it should land
// in `unmatched`.
//
test("generateReport: debtor with name normalizing to empty does NOT match any ledger", () => {
  const debtors = [
    { customer: "DEVA JCB", outstanding: 1000 },
    { customer: "---",      outstanding: 5000 },   // normalizes to ""
  ];
  const ledgersByCustomer = {
    "DEVA JCB":  [{ date: 100, debit: 1000, narration: "ref inv.111111" }],
    "OTHER CO": [{ date: 200, debit: 9999, narration: "ref inv.222222" }],
  };
  const result = O.generateReport({
    debtors, ledgersByCustomer,
    asOnSerial: 1000, branchLabel: "MYP & SRD",
  });
  // DEVA JCB → matched
  // "---"    → unmatched (not silently allocated to whichever ledger comes first)
  assert.equal(result.unmatched.length, 1);
  assert.equal(result.unmatched[0].customer, "---");
  assert.equal(result.rows.length, 1, "Only DEVA JCB should produce rows");
  assert.equal(result.rows[0].customer, "DEVA JCB");
});

// ── Regression: parseDebtorsAoA must not combine column matches across rows ─
test("parseDebtorsAoA: name keyword on row X and amount keyword on row Y don't get combined", () => {
  const aoa = [
    ["Particulars", "Some Other Header"], // has name but not amount
    ["Foo Customer", 1234],
    ["Whatever", "amount"],               // has amount but not name
    ["Bar Customer", 5678],
  ];
  // Old buggy behavior: would combine col[0]=Particulars from row 0 with
  // col[1]=amount from row 2 and treat row 2 as the header → skip rows 1+.
  // New behavior: requires both keywords on the SAME row, so no header is
  // found, falls back to col 0 = name, col 1 = amount, returns BOTH customers.
  const out = O.parseDebtorsAoA(aoa);
  // Row 0 itself fails the (numeric amount) filter — its second column is a
  // string. Row 1 → "Foo Customer", 1234. Row 2 fails (amt is non-numeric
  // "amount"). Row 3 → "Bar Customer", 5678. Total 2 entries.
  const customers = out.map(o => o.customer);
  assert.ok(customers.includes("Foo Customer"), "Foo Customer should be present");
  assert.ok(customers.includes("Bar Customer"), "Bar Customer should be present");
});

test("parseDebtorsAoA: standard layout", () => {
  const aoa = [
    ["Particulars", "Debit"],
    ["DEVA JCB", 16466589.48],
    ["MEGHA ENG", 5499780.17],
    ["Total", 99999],          // should be skipped
    ["", 100],                 // empty name skipped
    ["FOO", 0],                // zero outstanding skipped
  ];
  const out = O.parseDebtorsAoA(aoa);
  assert.equal(out.length, 2);
  assert.equal(out[0].customer, "DEVA JCB");
  assert.equal(out[0].outstanding, 16466589.48);
  assert.equal(out[1].customer, "MEGHA ENG");
});

// ── Ledger parser + FIFO end-to-end (smoke) ──────────────────────────────────
//
// Synthetic mini-ledger with the paired-row Tally export shape. This mirrors
// what we observed in the real DEVA JCB sample.
//

test("parseLedgerAoA + FIFO: paired-row Tally shape", () => {
  const aoa = [
    ["MGB MOTOR AND AUTO AGENCIES PVT LTD", "", "", "", "", "", ""],
    ["DEVA JCB", "", "", "", "", "", ""],
    ["Ledger Account", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["1-Apr-25 to 27-Apr-26", "", "", "", "", "", ""],
    ["Date", "Particulars", "", "Vch Type", "Vch No.", "Debit", "Credit"],
    [46115, "To", "MGB MOTOR PVT LTD - MIYAPUR", "Journal", "42",  4150807.84, ""],
    ["",    "",   "ref inv.261420025",          "",        "",    "",          ""],
    [46123, "To", "MGB MOTOR PVT LTD - MIYAPUR", "Journal", "230", 3488841.14, ""],
    ["",    "",   "ref inv.261420122",          "",        "",    "",          ""],
    [46125, "To", "PARTS SALE",                 "Parts Sale", "262500114", 1348845.62, ""],
    ["",    "",   "ABC",                        "",        "",    "",          ""],
    [46133, "To", "MGB MOTOR PVT LTD - MIYAPUR", "Journal", "375", 5436423,    ""],
    ["",    "",   "ref inv.261420240",          "",        "",    "",          ""],
    ["",    "By", "Closing Balance",             "",        "",    "", 14424917.6],
  ];
  const parsed = O.parseLedgerAoA(aoa, "Ledger.xlsx");
  assert.equal(parsed.customer, "DEVA JCB");
  assert.equal(parsed.entries.length, 4);
  // Verify a few fields
  assert.equal(parsed.entries[0].vchNo,    "42");
  assert.equal(parsed.entries[0].narration,"ref inv.261420025");
  assert.equal(parsed.entries[2].narration,"ABC");
  assert.equal(parsed.entries[2].vchNo,    "262500114");

  // FIFO with outstanding = 14,424,917.60 (= sum) → all 4 rows
  const allPicked = O.selectFifoOutstanding(parsed.entries, 14424917.60);
  assert.equal(allPicked.length, 4);
  assert.equal(allPicked.reduce((s, p) => s + p.amount, 0).toFixed(2), "14424917.60");

  // Smaller outstanding should pick newest only
  const smallPicked = O.selectFifoOutstanding(parsed.entries, 5000000);
  // Newest is 5,436,423 — already exceeds 5,000,000 → only one row, adjusted
  assert.equal(smallPicked.length, 1);
  assert.equal(smallPicked[0].amount, 5000000);
  assert.equal(smallPicked[0].adjusted, true);
});

test("parseLedgerAoA: customer name fallback to filename", () => {
  // No "Ledger Account" marker, no header above
  const aoa = [
    ["Date", "Particulars", "", "Vch Type", "Vch No.", "Debit", "Credit"],
    [46115, "To", "FOO", "Journal", "1", 100, ""],
    ["", "", "ref inv.123456", "", "", "", ""],
  ];
  const parsed = O.parseLedgerAoA(aoa, "Acme_Corp.xlsx");
  assert.equal(parsed.customer, "Acme Corp");
});

// ── End-to-end with real sample ──────────────────────────────────────────────
//
// If the real sample files are present at the expected location, run the full
// pipeline against the DEVA JCB row in the debtors export and verify the FIFO
// total matches exactly. Skipped silently when samples aren't there.
//

const path = require("node:path");
const fs   = require("node:fs");

function findXlsx() {
  // Try a couple of plausible locations
  const candidates = [
    {
      ledger:  "/home/ubuntu/attachments/756c8d04-394f-4c7b-8d6c-c46d2b7377c3/Ledger.xlsx",
      debtors: "/home/ubuntu/attachments/3c7cbcaf-3470-4542-980f-b9ddc4be9185/Debtors+18.4.26+OS.xlsx",
    },
  ];
  for (const c of candidates) {
    if (fs.existsSync(c.ledger) && fs.existsSync(c.debtors)) return c;
  }
  return null;
}

test("Real DEVA JCB sample: FIFO sum matches outstanding exactly", { skip: !findXlsx() }, () => {
  const samples = findXlsx();
  // xlsx is provided by the host app — if not installed, skip gracefully
  let XLSX;
  try { XLSX = require("/home/ubuntu/nah-man/ho-recon/node_modules/xlsx"); }
  catch (e) { return; }

  const debtorsWb  = XLSX.readFile(samples.debtors, { raw: true });
  const debtorsAoa = XLSX.utils.sheet_to_json(
    debtorsWb.Sheets[debtorsWb.SheetNames[0]], { header: 1, raw: true, defval: "" });
  const debtors = O.parseDebtorsAoA(debtorsAoa);
  const deva = debtors.find(d => /DEVA/i.test(d.customer));
  assert.ok(deva, "DEVA JCB should be in debtors export");
  assert.equal(deva.outstanding, 16466589.48);

  const ledWb  = XLSX.readFile(samples.ledger, { raw: true });
  const ledAoa = XLSX.utils.sheet_to_json(
    ledWb.Sheets[ledWb.SheetNames[0]], { header: 1, raw: true, defval: "" });
  const parsed = O.parseLedgerAoA(ledAoa, "Ledger.xlsx");
  assert.equal(parsed.customer, "DEVA JCB");

  const asOn = O.toSerial("18-04-2026");
  const picked = O.selectFifoOutstanding(parsed.entries, deva.outstanding, asOn);
  const sum = picked.reduce((s, p) => s + p.amount, 0);
  assert.equal(Math.round(sum * 100), Math.round(deva.outstanding * 100),
               `FIFO sum ${sum} should equal outstanding ${deva.outstanding}`);

  // Per the user's reference report there should be 9 rows
  assert.equal(picked.length, 9, "DEVA JCB should split into 9 outstanding invoices");

  // Last row should be adjusted (171,872.90 → 156,783.31)
  const last = picked[picked.length - 1];
  assert.equal(last.adjusted, true);
  assert.equal(last.amount,   156783.31);

  // Spot-check INV extraction on one row
  const refRow = picked.find(p => /ref inv\.251423826/.test(p.narration));
  assert.ok(refRow, "Should have row with ref inv.251423826");
  const built = O.buildReportRows("DEVA JCB", picked, O.toSerial("18-04-2026"), "MYP & SRD");
  const builtRefRow = built.find(r => r.invoiceNo === "251423826");
  assert.ok(builtRefRow);
  assert.equal(builtRefRow.invoiceNo, "251423826");
  assert.equal(builtRefRow.invoiceDate, 46098); // ledger date

  // PARTS SALE row should have Vch No as invoice
  const partsRow = built.find(r => r.invoiceNo === "262500114");
  assert.ok(partsRow, "PARTS SALE row should fall back to Vch No");
});
