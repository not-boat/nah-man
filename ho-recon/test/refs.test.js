"use strict";
// Run with:  node --test test/refs.test.js
// Tests the reference / UTR extractor against realistic Indian banking
// narration patterns seen in Tally exports.

const test = require("node:test");
const assert = require("node:assert/strict");
const Refs = require("../src/refs.js");

function values(refs) { return refs.map(r => r.value); }
function kinds(refs)  { return refs.map(r => r.kind); }

// ── Bare UPI 12-digit ────────────────────────────────────────────────────────
test("UPI 12-digit bare", () => {
  const r = Refs.extractRefs("upi 412345678901 by mr ramesh");
  assert.equal(r[0].kind, "UPI");
  assert.equal(r[0].value, "412345678901");
});

test("UPI inside cluttered narration with date+amount+name", () => {
  const r = Refs.extractRefs(
    "upi/412345678901/sree durga earth movers/04-12-2024/50000.00/payment for inv 255320402",
    50000
  );
  // Should pull both the UPI ref and the invoice number, NOT the amount or date
  const vals = values(r);
  assert.ok(vals.includes("412345678901"), "UPI ref present");
  assert.ok(vals.includes("INV255320402") || vals.includes("255320402"), "Invoice ref present");
  assert.ok(!vals.includes("50000"),  "amount filtered");
  assert.ok(!vals.includes("04122024"), "packed date not picked");
});

// ── NEFT with HDFC IFSC prefix ──────────────────────────────────────────────
test("NEFT HDFCNxxxx with bank prefix", () => {
  const r = Refs.extractRefs("by neft HDFCN24070112345678 from acme corp");
  assert.equal(r[0].kind, "NEFT");
  assert.equal(r[0].value, "HDFCN24070112345678");
});

test("RTGS with UTIBR prefix", () => {
  const r = Refs.extractRefs("rtgs ref UTIBR21345678901234");
  assert.equal(r[0].kind, "RTGS");
  assert.equal(r[0].value, "UTIBR21345678901234");
});

test("IMPS with HDFCH prefix", () => {
  const r = Refs.extractRefs("imps txn HDFCH240701567890");
  assert.equal(r[0].kind, "IMPS");
});

// ── Labeled refs ─────────────────────────────────────────────────────────────
test("Ref id with explicit label", () => {
  const r = Refs.extractRefs("Ref id - 678901234567 by transfer");
  assert.equal(r[0].value, "678901234567");
  // Could be classified as UPI (12-digit) or LABELED — both acceptable
  assert.ok(["UPI","LABELED","NUMERIC"].includes(r[0].kind));
});

test("Txn ID label", () => {
  const r = Refs.extractRefs("Txn ID: ABC1234567890");
  const vals = values(r);
  assert.ok(vals.some(v => v.includes("ABC1234567890") || v === "1234567890"));
});

test("UTR No. label (mixed punctuation)", () => {
  const r = Refs.extractRefs("UTR No.: 901234567890 — payment received");
  assert.equal(r[0].value, "901234567890");
});

// ── CMS ──────────────────────────────────────────────────────────────────────
test("CMS prefix", () => {
  const r = Refs.extractRefs("CMS123456789012 collection");
  assert.equal(r[0].kind, "CMS");
});

// ── Invoice ──────────────────────────────────────────────────────────────────
test("Invoice number 'inv 255320402'", () => {
  const r = Refs.extractRefs("payment for inv 255320402 sree durga");
  const vals = values(r);
  assert.ok(vals.includes("INV255320402") || vals.includes("255320402"));
});

test("Invoice number 'inv.255320402' (no space)", () => {
  const r = Refs.extractRefs("inv.255320402");
  const vals = values(r);
  assert.ok(vals.includes("INV255320402") || vals.includes("255320402"));
});

// ── Cheque ───────────────────────────────────────────────────────────────────
test("Cheque number 'chq 000123'", () => {
  const r = Refs.extractRefs("chq no 000123 deposit");
  const vals = values(r);
  assert.ok(vals.length > 0, "extracted at least one ref");
});

// ── Noise rejection ──────────────────────────────────────────────────────────
test("Indian mobile 10-digit (starts 9) rejected as ref", () => {
  const r = Refs.extractRefs("call vendor 9876543210 for confirmation");
  // 9876543210 is 10 digits → would match NUMERIC pattern, but isLikelyMobile rejects it
  const vals = values(r);
  assert.ok(!vals.includes("9876543210"), "mobile not treated as ref");
});

test("4-digit year not a ref", () => {
  const r = Refs.extractRefs("year 2024 settlement");
  assert.equal(r.length, 0);
});

test("GSTIN not a ref", () => {
  const r = Refs.extractRefs("gst 27AAAPL1234C1Z5 invoice");
  const vals = values(r);
  assert.ok(!vals.includes("27AAAPL1234C1Z5"));
});

test("Entry's own amount not picked as ref", () => {
  // ₹50000 inside narration must not be treated as a ref when amount=50000
  const r = Refs.extractRefs("by transfer 50000 to ho payment", 50000);
  const vals = values(r);
  assert.ok(!vals.includes("50000"), "own amount filtered");
});

// ── Multi-ref + de-duplication ───────────────────────────────────────────────
test("Multiple refs in one narration are all extracted", () => {
  const r = Refs.extractRefs(
    "neft HDFCN24070112345 utr 412345678901 inv 999888777 — settlement"
  );
  const vals = values(r);
  assert.ok(vals.includes("HDFCN24070112345"));
  assert.ok(vals.includes("412345678901"));
  assert.ok(vals.some(v => v.includes("999888777")));
});

// ── Just-a-reference (sparse narration) ──────────────────────────────────────
test("Sparse narration: just a UTR", () => {
  const r = Refs.extractRefs("412345678901");
  assert.equal(r[0].value, "412345678901");
  assert.equal(r[0].kind, "UPI");
});

test("Sparse narration: just date+ref", () => {
  const r = Refs.extractRefs("04/12/2024 HDFCN24070112345678");
  const vals = values(r);
  assert.ok(vals.includes("HDFCN24070112345678"));
});

// ── refsMatchScore ───────────────────────────────────────────────────────────
test("refsMatchScore: exact UPI match scores high", () => {
  const a = Refs.extractRefs("upi 412345678901 sree durga");
  const b = Refs.extractRefs("412345678901 by transfer");
  const r = Refs.refsMatchScore(a, b);
  assert.ok(r);
  assert.ok(r.score >= 80, `expected ≥80, got ${r.score}`);
});

test("refsMatchScore: cross-kind same digits also matches", () => {
  // Branch logged it as labeled "ref id 412345678901"
  // HO logged it as a bare UPI 12-digit
  const a = Refs.extractRefs("Ref id - 412345678901");
  const b = Refs.extractRefs("upi 412345678901");
  const r = Refs.refsMatchScore(a, b);
  assert.ok(r);
  assert.ok(r.score >= 80, `expected ≥80 cross-kind exact, got ${r.score}`);
});

test("refsMatchScore: fuzzy ~1 typo on long NEFT ref still matches", () => {
  // OCR / typo: one digit off
  const a = Refs.extractRefs("HDFCN24070112345678");
  const b = Refs.extractRefs("HDFCN24070112345679"); // last digit differs
  const r = Refs.refsMatchScore(a, b);
  assert.ok(r);
  assert.ok(r.score >= 50, `expected ≥50 (1-typo fuzzy), got ${r.score}`);
});

test("refsMatchScore: last-8 partial on same kind", () => {
  const a = Refs.extractRefs("HDFCN24070198765432");
  const b = Refs.extractRefs("ICICN23010198765432"); // different bank, same trailing 8
  const r = Refs.refsMatchScore(a, b);
  assert.ok(r);
  assert.ok(r.score >= 30, `expected ≥30 (last-8), got ${r.score}`);
});

test("refsMatchScore: completely different refs return null or low score", () => {
  const a = Refs.extractRefs("upi 111111111111");
  const b = Refs.extractRefs("upi 222222222222");
  const r = Refs.refsMatchScore(a, b);
  assert.ok(!r || r.score < 30);
});

test("refsMatchScore: multi-ref boost when 2+ refs match", () => {
  const a = Refs.extractRefs("upi 412345678901 inv 255320402");
  const b = Refs.extractRefs("412345678901 invoice 255320402 sree durga");
  const r = Refs.refsMatchScore(a, b);
  assert.ok(r);
  // Both UPI & invoice match — should reach into the 90s
  assert.ok(r.score >= 85, `expected ≥85 multi-match, got ${r.score}`);
});

// ── jaroWinkler ──────────────────────────────────────────────────────────────
test("jaroWinkler: typo'd company name", () => {
  const sim = Refs.jaroWinkler("Sree Durga Earth Movers", "SREE DURGHA EARTHMOVERS");
  assert.ok(sim >= 0.85, `expected ≥0.85, got ${sim}`);
});

test("jaroWinkler: completely different names", () => {
  const sim = Refs.jaroWinkler("Acme Corp", "Globex Industries");
  assert.ok(sim < 0.6);
});

// ── Levenshtein ──────────────────────────────────────────────────────────────
test("levenshtein: 1-character difference", () => {
  assert.equal(Refs.levenshtein("HDFCN24070112345678", "HDFCN24070112345679"), 1);
});
test("levenshtein: identical", () => {
  assert.equal(Refs.levenshtein("abc", "abc"), 0);
});

// ── `--<amount>` divider — must not be concatenated onto the UTR ────────────
test("UTR with --amount divider: bank-prefix ref stays clean (Sai Enterprises)", () => {
  const r = Refs.extractRefs("inv 252702397 sai enterprises ref no UBINH25305820666--568695/-", 568695);
  const vals = values(r);
  assert.ok(vals.includes("UBINH25305820666"), `UBIN ref must be clean (16 chars). got=${JSON.stringify(vals)}`);
  assert.ok(!vals.includes("UBINH25305820666568695"), "amount must not be concatenated onto UTR");
  // The amount must also be filtered out as a standalone numeric ref
  assert.ok(!vals.includes("568695"), "amount filtered as own-amount");
});

test("UTR with --amount divider: NEFT HDFC ref stays clean", () => {
  const r = Refs.extractRefs("inv 252810639 ref no BKIDN62025111132642415--37147-inv amt 28000", 28000);
  const vals = values(r);
  assert.ok(vals.includes("BKIDN62025111132642415"), `BKID ref must be clean. got=${JSON.stringify(vals)}`);
  assert.ok(!vals.includes("BKIDN6202511113264241537147"), "amount must not be concatenated");
});

test("UTR with --amount divider: 12-digit UPI not affected (already worked)", () => {
  const r = Refs.extractRefs("inv 252702450 utr no 985494211238--12850 Gayathri Infra concrete", 12850);
  const vals = values(r);
  assert.ok(vals.includes("985494211238"), "12-digit UPI extracted cleanly");
  assert.ok(!vals.includes("98549421123812850"), "amount must not be concatenated");
});

// ── Cross-kind same-bank-prefix partial match ───────────────────────────────
test("Cross-kind same-bank-prefix: IMPS UBINH vs RTGS UBINR with same last-6", () => {
  const a = Refs.extractRefs("ref no UBINH25305820666", 568695);
  const b = Refs.extractRefs("RTGS CR-UBIN0814776-UBINR22025110101820666", 568695);
  const m = Refs.refsMatchScore(a, b);
  assert.ok(m, "expected a match");
  assert.ok(/cross-kind/.test(m.method), `expected cross-kind method, got ${m.method}`);
  assert.ok(m.score >= 30, `expected score >= 30, got ${m.score}`);
  assert.equal(m.exact, false, "cross-kind partial must require amount corroboration");
});

test("Cross-kind ONLY fires for known bank prefixes", () => {
  // ZZZZH / ZZZZR — unknown prefix, should NOT cross-kind match
  const a = [{ kind: "IMPS", value: "ZZZZH25305820666", strength: 80 }];
  const b = [{ kind: "RTGS", value: "ZZZZR22025110101820666", strength: 88 }];
  const m = Refs.refsMatchScore(a, b);
  // No suffix endsWith, no same-kind last-6, unknown prefix → null
  assert.equal(m, null, "unknown bank prefix must not cross-kind match");
});

// ── End-to-end: BR vs HO with --amount divider must reconcile via refs ─────
test("End-to-end: BR `utr no 985494211238--12850` matches HO `UPI/985494211238/...` exactly", () => {
  const br = Refs.extractRefs("inv 252702450 utr no 985494211238--12850 Gayathri Infra concrete", 12850);
  const ho = Refs.extractRefs("UPI/985494211238/JCB tyre paymen/kalakondakumara/State Bank Of I/AXLcb6f26145da64fda817c8b23d5f22856", 12850);
  const m = Refs.refsMatchScore(br, ho);
  assert.ok(m, "should match");
  assert.equal(m.exact, true, "must be exact UTR match");
  assert.ok(m.score >= 80, `expected high score, got ${m.score}`);
  assert.equal(m.method, "UPI Exact");
});

// ── NEFT/RTGS/IMPS labels in the LABELED regex ──────────────────────────────
test("Labeled regex recognises `neft` / `rtgs` / `imps` / `upi` keywords", () => {
  const cases = [
    ["neft HDFCN24070112345", "HDFCN24070112345"],
    ["rtgs UTIBR21345678",     "UTIBR21345678"],
    ["imps 123412341234",      "123412341234"],
    ["upi 412345678901",       "412345678901"],
  ];
  for (const [txt, want] of cases) {
    const r = Refs.extractRefs(txt);
    assert.ok(values(r).includes(want), `${txt} → expected ${want}, got ${JSON.stringify(values(r))}`);
  }
});
