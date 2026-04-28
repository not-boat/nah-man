"use strict";
// ═══════════════════════════════════════════════════════════════════════════════
// O/S REPORT — pure logic
// ═══════════════════════════════════════════════════════════════════════════════
//
// All parsing / FIFO / report-building logic. No Electron, no fs. Takes raw
// AoA (array-of-arrays) from sheets in, returns structured data out — so it
// can be unit-tested with `node --test`.
//
// ── Data flow ──────────────────────────────────────────────────────────────
//   parseDebtorsAoA(aoa)         → [{ customer, outstanding }, ...]
//   parseLedgerAoA(aoa, fname)   → { customer, entries: [{ date, particulars,
//                                     vchType, vchNo, debit, narration }, ...] }
//   selectFifoOutstanding(entries, outstanding)
//                                → [{ ...entry, amount, adjusted }, ...]  newest→older
//   buildReportRows(customer, picked, asOnSerial)
//                                → MYP&SRD-template rows
//   ageingBucketLabel(days)      → "0-30" | "31-60" | "61-90" | "91-120" | "121-Above"
//
// ═══════════════════════════════════════════════════════════════════════════════

(function (global) {
  // ── Date helpers ────────────────────────────────────────────────────────────

  // Convert any reasonable date input → Excel serial number (whole days).
  // Accepts: Excel serial (number), JS Date, "DD-MMM-YY[YY]", "DD/MM/YYYY",
  //   "DD-MM-YYYY", "YYYY-MM-DD", "YYYYMMDD", "DDMMYYYY".
  function toSerial(d) {
    if (d == null || d === "") return null;
    if (typeof d === "number" && isFinite(d)) return Math.floor(d);
    if (d instanceof Date && !isNaN(d)) return jsDateToSerial(d);

    const s = String(d).trim();
    if (!s) return null;

    // 03-Apr-25 / 03-Apr-2025
    const tally = s.match(/^(\d{1,2})[\/\-]([A-Za-z]{3})[\/\-](\d{2,4})$/);
    if (tally) {
      const months = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };
      const m = months[tally[2].toLowerCase()];
      if (!m) return null;
      let y = parseInt(tally[3], 10);
      if (y < 100) y = (y > 50 ? 1900 : 2000) + y;
      return jsDateToSerial(new Date(Date.UTC(y, m - 1, parseInt(tally[1], 10))));
    }

    // DD/MM/YYYY or DD-MM-YYYY
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m) {
      return jsDateToSerial(new Date(Date.UTC(+m[3], +m[2] - 1, +m[1])));
    }
    // YYYY-MM-DD or YYYY/MM/DD
    m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) {
      return jsDateToSerial(new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])));
    }
    // YYYYMMDD or DDMMYYYY (8-digit packed)
    m = s.match(/^(\d{8})$/);
    if (m) {
      const x = m[1];
      if (parseInt(x.slice(0, 4), 10) > 1900) {
        return jsDateToSerial(new Date(Date.UTC(+x.slice(0,4), +x.slice(4,6) - 1, +x.slice(6,8))));
      }
      return jsDateToSerial(new Date(Date.UTC(+x.slice(4,8), +x.slice(2,4) - 1, +x.slice(0,2))));
    }
    return null;
  }

  // JS Date → Excel serial. Excel epoch = 1899-12-30 (works around 1900 bug).
  function jsDateToSerial(jsDate) {
    const ms = Date.UTC(jsDate.getUTCFullYear(), jsDate.getUTCMonth(), jsDate.getUTCDate());
    const epoch = Date.UTC(1899, 11, 30);
    return Math.round((ms - epoch) / 86400000);
  }

  // Excel serial → "DD-MM-YYYY"
  function serialToDDMMYYYY(serial) {
    if (serial == null) return "";
    const ms = Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000;
    const d = new Date(ms);
    const dd = String(d.getUTCDate()).padStart(2, "0");
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    const yy = d.getUTCFullYear();
    return `${dd}-${mm}-${yy}`;
  }

  // ── Invoice extractor ───────────────────────────────────────────────────────
  //
  // Tries (in order):
  //   1. "ref inv. <digits>"      → user's existing pattern
  //   2. "ref inv <digits>"
  //   3. "Invoice: <digits>"      / "Invoice no <digits>" / "inv <digits>"
  //   4. "REF <digits>"           → uppercase tag, no "inv"
  //   5. Bare 8+ digit run inside narration
  //
  // Always strips whitespace / CRLF / trailing date suffix like "  on 25-01-2025".
  // If nothing matches, returns "".
  const INV_PATTERNS = [
    /\bref\s*inv\.?\s*[#:]?\s*([\d\-\/]{6,})/i,
    /\binvoice\s*[#:no.]*\s*([\d\-\/]{6,})/i,
    /\binv\s*[#:.no]*\s*([\d\-\/]{6,})/i,
    /\bref\s+([\d\-\/]{6,})/i,
    /\b(\d{8,})\b/,
  ];

  function extractInvFromNarration(narration) {
    if (!narration) return "";
    const s = String(narration).replace(/[\r\n\t]+/g, " ").trim();
    for (const re of INV_PATTERNS) {
      const m = s.match(re);
      if (m) return String(m[1]).trim();
    }
    return "";
  }

  // ── Debtors parser ──────────────────────────────────────────────────────────
  //
  // Expected shape (from Tally "Outstandings → Group → Sundry Debtors → Excel"):
  //   row 0: ["Particulars", "Debit"]   ← header
  //   row 1+: [Customer, OutstandingAmount]
  // Amount is in the "Debit" column (Tally convention: receivable = debit balance).
  // Customers with non-positive or empty outstanding are dropped.
  //
  function parseDebtorsAoA(aoa) {
    if (!Array.isArray(aoa) || !aoa.length) return [];
    let headerIdx = -1, colName = -1, colAmt = -1;
    for (let i = 0; i < Math.min(aoa.length, 10); i++) {
      const row = aoa[i] || [];
      for (let c = 0; c < row.length; c++) {
        const v = String(row[c] || "").toLowerCase().trim();
        if (v === "particulars" || v === "name" || v === "customer" || v === "ledger") colName = c;
        if (v === "debit" || v === "amount" || v === "outstanding" || v === "balance"
            || v === "closing balance" || v === "closing") colAmt = c;
      }
      if (colName !== -1 && colAmt !== -1) { headerIdx = i; break; }
    }
    if (headerIdx === -1) {
      // Fallback: assume col 0 = name, col 1 = amount, no header
      colName = 0; colAmt = 1; headerIdx = -1;
    }
    const out = [];
    for (let i = headerIdx + 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const name = String(row[colName] || "").trim();
      const amt = Number(row[colAmt]);
      if (!name || !isFinite(amt) || amt <= 0) continue;
      // Skip "Grand Total" / "Total" footer rows
      if (/^(grand\s+)?total$/i.test(name)) continue;
      out.push({ customer: name, outstanding: amt });
    }
    return out;
  }

  // ── Ledger parser ───────────────────────────────────────────────────────────
  //
  // Tally ledger export shape (paired-row format):
  //   row 0: company name  (e.g. "MGB MOTOR AND AUTO AGENCIES PVT LTD")
  //   row 1..k: address lines
  //   row k: customer ledger name (e.g. "DEVA JCB")        ← we capture this
  //   row k+1: "Ledger Account"
  //   row k+...: blank, date range
  //   row h: header  Date | Particulars | (blank) | Vch Type | Vch No. | Debit | Credit
  //   row h+1+: pairs of:
  //     main row     [date, "To"|"By", particulars, vchType, vchNo, debit, credit]
  //     narration row[ "" ,    ""    , narration  ,    ""  ,   ""   ,   "" ,   "" ]
  //   Last rows: closing-balance summary.
  //
  // We only capture entries with side === "To" (debit side). All netting has
  // already happened upstream — credit side here is only the closing-balance
  // line, never an actual payment we care about.
  //
  function parseLedgerAoA(aoa, fname) {
    const out = { customer: "", entries: [], warnings: [] };
    if (!Array.isArray(aoa) || !aoa.length) return out;

    // 1. Find customer name: row immediately above "Ledger Account"
    let ledgerAccountIdx = -1;
    for (let i = 0; i < Math.min(aoa.length, 30); i++) {
      const cell = String((aoa[i] || [])[0] || "").toLowerCase().trim();
      if (cell === "ledger account") { ledgerAccountIdx = i; break; }
    }
    if (ledgerAccountIdx > 0) {
      // Walk backward to the nearest non-empty cell — that's the customer name
      for (let i = ledgerAccountIdx - 1; i >= 0; i--) {
        const cell = String((aoa[i] || [])[0] || "").trim();
        if (cell && !/^\s*\d/.test(cell)) { out.customer = cell; break; }
      }
    }
    if (!out.customer && fname) {
      // Fall back to filename stem
      out.customer = String(fname).replace(/\.[^.]+$/, "").replace(/[_\-]+/g, " ").trim();
    }

    // 2. Find header row
    let headerIdx = -1;
    let colDate = -1, colSide = -1, colParticulars = -1, colVchType = -1,
        colVchNo = -1, colDebit = -1, colCredit = -1;

    for (let i = 0; i < Math.min(aoa.length, 30); i++) {
      const row = aoa[i] || [];
      let foundDate = -1, foundParticulars = -1, foundDebit = -1, foundCredit = -1,
          foundVchType = -1, foundVchNo = -1;
      for (let c = 0; c < row.length; c++) {
        const v = String(row[c] || "").toLowerCase().trim();
        if (v === "date") foundDate = c;
        else if (v === "particulars") foundParticulars = c;
        else if (v === "debit") foundDebit = c;
        else if (v === "credit") foundCredit = c;
        else if (v === "vch type" || v === "vchtype") foundVchType = c;
        else if (v === "vch no" || v === "vch no." || v === "vchno") foundVchNo = c;
      }
      if (foundDate !== -1 && foundParticulars !== -1
          && foundDebit !== -1 && foundCredit !== -1) {
        headerIdx     = i;
        colDate       = foundDate;
        colParticulars= foundParticulars;
        colDebit      = foundDebit;
        colCredit     = foundCredit;
        colVchType    = foundVchType;
        colVchNo      = foundVchNo;
        break;
      }
    }
    if (headerIdx === -1) {
      out.warnings.push("No header row found (expected Date / Particulars / Debit / Credit)");
      return out;
    }

    // 2b. Detect the "side" column. Tally exports the "Particulars" header as
    // a single cell that visually spans two columns: the LEFT cell holds the
    // side prefix ("To" / "By") and the RIGHT cell holds the actual party /
    // narration text. Sample first few data rows to figure out which is which.
    {
      let toByCount = 0, sampled = 0;
      for (let i = headerIdx + 1; i < aoa.length && sampled < 10; i++) {
        const row = aoa[i] || [];
        const v = String(row[colParticulars] || "").trim();
        if (!v) continue;
        sampled++;
        if (v === "To" || v === "By" || v === "to" || v === "by") toByCount++;
      }
      if (sampled >= 3 && toByCount >= sampled - 1) {
        // colParticulars is actually the side column; the real particulars sit
        // one column to the right.
        colSide        = colParticulars;
        colParticulars = colParticulars + 1;
      } else {
        colSide = -1;
      }
    }

    // 3. Walk paired rows
    for (let i = headerIdx + 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const dateRaw = row[colDate];
      const dateSerial = toSerial(dateRaw);

      // A "main" row has a date AND a debit or credit number
      const debit  = Number(row[colDebit])  || 0;
      const credit = Number(row[colCredit]) || 0;
      const side   = colSide >= 0 ? String(row[colSide] || "").trim() : "";

      if (dateSerial && (debit > 0 || credit > 0)) {
        // Skip closing-balance line (side="By" + debit==0)
        if (side === "By" || credit > 0) continue;
        // It's a debit entry. Pull narration from next row's particulars cell.
        let narration = "";
        if (i + 1 < aoa.length) {
          const next = aoa[i + 1] || [];
          const isNextMain = toSerial(next[colDate]) != null
                          || Number(next[colDebit])  > 0
                          || Number(next[colCredit]) > 0;
          if (!isNextMain) {
            narration = String(next[colParticulars] || "").trim();
          }
        }
        out.entries.push({
          date:        dateSerial,
          particulars: String(row[colParticulars] || "").trim(),
          vchType:     colVchType >= 0 ? String(row[colVchType] || "").trim() : "",
          vchNo:       colVchNo   >= 0 ? String(row[colVchNo]   || "").trim() : "",
          debit:       debit,
          narration:   narration,
        });
      }
    }
    return out;
  }

  // ── FIFO selection ──────────────────────────────────────────────────────────
  //
  // Walk debit entries newest-first; accumulate until cumulative ≥ outstanding;
  // adjust the final entry by (cumulative - outstanding). Returns the picked
  // entries with an extra `amount` field — full debit, except the last.
  //
  // If the entries don't sum to ≥ outstanding (shouldn't happen with a real
  // ledger, but possible for partially-exported / incomplete data), returns
  // everything and the caller can flag it via the totals.
  //
  function selectFifoOutstanding(entries, outstanding, asOnSerial) {
    const cutoff = (asOnSerial != null && isFinite(asOnSerial)) ? asOnSerial : Infinity;
    const sorted = entries
      .filter(e => e.debit > 0 && e.date <= cutoff)
      .slice()
      // Newest first. For equal dates, keep the original (input) order so
      // that the FIFO walk is deterministic and matches the user's eyeballed
      // ordering of same-day entries.
      .sort((a, b) => b.date - a.date);
    const picked = [];
    let cum = 0;
    const target = Number(outstanding);
    if (!isFinite(target) || target <= 0) return picked;
    for (const e of sorted) {
      const remaining = target - cum;
      if (remaining <= 0) break;
      if (e.debit >= remaining - 0.005) {
        // Last (oldest in picked set) — adjust to exactly hit the target
        picked.push(Object.assign({}, e, { amount: round2(remaining), adjusted: e.debit !== remaining }));
        cum = target;
        break;
      }
      picked.push(Object.assign({}, e, { amount: round2(e.debit), adjusted: false }));
      cum += e.debit;
    }
    return picked;
  }

  function round2(n) { return Math.round(n * 100) / 100; }

  // ── Ageing buckets ──────────────────────────────────────────────────────────
  //
  // Buckets per the user's MYP&SRD template:
  //   0-30, 31-60, 61-90, 91-120, 121-Above
  // Days = max(0, asOnSerial - invoiceDateSerial)
  //
  function ageingBucketLabel(days) {
    if (days <= 30)  return "0-30";
    if (days <= 60)  return "31-60";
    if (days <= 90)  return "61-90";
    if (days <= 120) return "91-120";
    return "121-Above";
  }

  // ── Report row builder ──────────────────────────────────────────────────────
  //
  // Output columns (mirrors the user's MYP&SRD sheet):
  //   Branches | Customer Name | Invoice No | Invoice Date | Due Days
  //   | 0-30 | 31-60 | 61-90 | 91-120 | 121-Above | Grand Total
  //
  // Invoice No: extracted from narration if `ref inv.` / `Invoice:` etc; else
  //   falls back to Vch No (which IS the invoice for direct-sale rows like
  //   PARTS SALE / OILS SALE).
  //
  function buildReportRows(customer, picked, asOnSerial, branchLabel = "MYP & SRD") {
    return picked.map(p => {
      const inv = extractInvFromNarration(p.narration) || p.vchNo || "";
      const days = Math.max(0, (asOnSerial || 0) - p.date);
      const bucket = ageingBucketLabel(days);
      const buckets = { "0-30": "", "31-60": "", "61-90": "", "91-120": "", "121-Above": "" };
      buckets[bucket] = round2(p.amount);
      return {
        branches:    branchLabel,
        customer:    customer,
        invoiceNo:   inv,
        invoiceDate: p.date,
        dueDays:     days,
        b0_30:       buckets["0-30"],
        b31_60:      buckets["31-60"],
        b61_90:      buckets["61-90"],
        b91_120:     buckets["91-120"],
        b121:        buckets["121-Above"],
        grandTotal:  round2(p.amount),
        adjusted:    !!p.adjusted,
        narration:   p.narration,
        particulars: p.particulars,
      };
    });
  }

  // ── End-to-end orchestrator ────────────────────────────────────────────────
  //
  // Given parsed debtors + parsed ledgers map (customer-key → entries[]),
  // returns:
  //   { rows: [...], reconciliation: [...], unmatched: [...] }
  //
  function generateReport({ debtors, ledgersByCustomer, asOnSerial, branchLabel }) {
    const rows = [];
    const reconciliation = [];
    const unmatched = [];
    const ledgerKeys = Object.keys(ledgersByCustomer);
    const normMap = new Map();
    for (const k of ledgerKeys) normMap.set(normalizeName(k), k);

    for (const d of debtors) {
      const key = normalizeName(d.customer);
      const ledgerKey = normMap.get(key)
        || ledgerKeys.find(k => normalizeName(k).startsWith(key) || key.startsWith(normalizeName(k)));
      if (!ledgerKey) {
        unmatched.push({ customer: d.customer, outstanding: d.outstanding,
                          reason: "No ledger file found for this customer" });
        reconciliation.push({ customer: d.customer, outstanding: d.outstanding,
                              picked: 0, status: "missing-ledger", variance: d.outstanding });
        continue;
      }
      const entries = ledgersByCustomer[ledgerKey];
      const picked = selectFifoOutstanding(entries, d.outstanding, asOnSerial);
      const pickedSum = picked.reduce((s, p) => s + p.amount, 0);
      const variance = round2(d.outstanding - pickedSum);
      const built = buildReportRows(d.customer, picked, asOnSerial, branchLabel);
      rows.push(...built);
      reconciliation.push({
        customer:     d.customer,
        outstanding:  d.outstanding,
        picked:       picked.length,
        pickedSum:    round2(pickedSum),
        variance:     variance,
        status:       Math.abs(variance) < 1 ? "ok" : "short",
      });
    }
    return { rows, reconciliation, unmatched };
  }

  function normalizeName(s) {
    return String(s || "")
      .toUpperCase()
      .replace(/[^A-Z0-9]/g, "");
  }

  // ── Exports ─────────────────────────────────────────────────────────────────
  const api = {
    toSerial,
    serialToDDMMYYYY,
    extractInvFromNarration,
    parseDebtorsAoA,
    parseLedgerAoA,
    selectFifoOutstanding,
    ageingBucketLabel,
    buildReportRows,
    generateReport,
    normalizeName,
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  if (global) global.OSReport = api;
})(typeof window !== "undefined" ? window : (typeof globalThis !== "undefined" ? globalThis : null));
