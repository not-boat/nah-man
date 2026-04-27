"use strict";
// ═══════════════════════════════════════════════════════════════════════════════
// REFERENCE / UTR EXTRACTION & MATCHING — v3
// ═══════════════════════════════════════════════════════════════════════════════
//
// Extracts every plausible payment reference (UPI / NEFT / RTGS / IMPS / CMS /
// cheque / invoice / labeled / bare numeric) out of a free-form narration string,
// rejects "noise" tokens (the entry's own amount, dates, mobile numbers, GSTINs,
// PANs, years), and exposes utilities to score how well two ref-arrays match —
// including fuzzy match (Levenshtein) and same-suffix partial match.
//
// Loaded both from main.js (Node) via require, and from index.html as a script
// tag — so it sets globals on window when no module system is present.
// ═══════════════════════════════════════════════════════════════════════════════

(function (global) {
  // ── Bank IFSC prefixes (first 4 chars of NEFT/RTGS/IMPS UTRs) ───────────────
  // Top ~40 Indian banks. Add more as needed; matching falls back to a generic
  // "[A-Z]{4}" rule if the prefix isn't in this list, but listed prefixes get
  // a confidence bump.
  const BANK_PREFIXES = new Set([
    "HDFC","ICIC","SBIN","UTIB","AXIS","KKBK","YESB","BARB","PUNB","CITI",
    "BKID","IDIB","IDFB","INDB","IOBA","MAHB","ORBC","SYNB","UCBA","UBIN",
    "ALLA","ANDB","CBIN","CNRB","CORP","DENA","DLXB","FDRL","KARB","KVBL",
    "LAVB","PSIB","RATN","SCBL","SIBL","SRCB","TMBL","VIJB","DBSS","DEUT",
    "HSBC","JAKA","BACB","ESFB","AUBL","BANK","UJVN","NSPB","FINO","PYTM",
  ]);

  // Ref kinds, in descending native strength
  const KIND_STRENGTH = {
    UPI:     90,
    NEFT:    88,
    RTGS:    88,
    IMPS:    80,
    UTR:     80,   // generic "UTR / TXN" labeled with no bank prefix
    CMS:     75,
    LABELED: 65,
    CHEQUE:  60,
    INVOICE: 55,
    NUMERIC: 25,   // bare digit blob fallback — must be corroborated
  };

  // ── Helpers ─────────────────────────────────────────────────────────────────
  function normRef(s) {
    return String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  }

  function isLikelyMobile(s) {
    return /^[6-9]\d{9}$/.test(s);
  }
  function isLikelyYear(s) {
    return /^(19|20)\d{2}$/.test(s);
  }
  function isLikelyDate(s) {
    // DDMMYYYY / YYYYMMDD packed dates that show up as "ddmmyyyy" inside text
    if (!/^\d{8}$/.test(s)) return false;
    const a = parseInt(s.slice(0, 4), 10);
    if (a > 1900 && a < 2100) {
      const m = parseInt(s.slice(4, 6), 10);
      const d = parseInt(s.slice(6, 8), 10);
      return m >= 1 && m <= 12 && d >= 1 && d <= 31;
    }
    const dd = parseInt(s.slice(0, 2), 10);
    const mm = parseInt(s.slice(2, 4), 10);
    const yy = parseInt(s.slice(4, 8), 10);
    return dd >= 1 && dd <= 31 && mm >= 1 && mm <= 12 && yy >= 1900 && yy < 2100;
  }
  function isGSTIN(s) {
    return /^[0-9]{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]{3}$/.test(s);
  }
  function isPAN(s) {
    return /^[A-Z]{5}\d{4}[A-Z]$/.test(s);
  }

  // Should this raw token be ignored as "noise" given the entry's own amount?
  function isNoise(token, ownAmount) {
    if (!token) return true;
    if (isLikelyMobile(token))               return true;
    if (isLikelyYear(token))                 return true;
    if (isLikelyDate(token))                 return true;
    if (isGSTIN(token))                      return true;
    if (isPAN(token))                        return true;
    if (ownAmount != null && ownAmount > 0) {
      const amt = Math.round(parseFloat(ownAmount));
      if (!isNaN(amt) && amt > 0) {
        const s = String(amt);
        if (token === s || token === s + "00") return true;        // ₹50000 or 5000000 paise
      }
    }
    return false;
  }

  // ── Classify a normalized token ─────────────────────────────────────────────
  // Returns { kind, value } or null if the token isn't a plausible ref.
  function classify(tok) {
    if (!tok) return null;

    // 4-letter bank prefix + digits → NEFT/RTGS/IMPS family
    const bankM = tok.match(/^([A-Z]{4})([A-Z0-9]*\d{6,})$/);
    if (bankM) {
      const prefix = bankM[1];
      // RTGS / NEFT / IMPS hint chars sometimes appear right after the IFSC prefix:
      //   HDFCN.... = NEFT  (N)
      //   HDFCR.... = RTGS  (R)
      //   HDFCH.... = IMPS  / FT (H)
      const hint = bankM[2][0];
      let kind = "NEFT"; // default — most common for IFSC-prefixed refs
      if (hint === "R") kind = "RTGS";
      else if (hint === "H" || hint === "I") kind = "IMPS";
      const known = BANK_PREFIXES.has(prefix);
      if (!known && tok.length < 10) return null; // unknown 4-letter + short digits → too weak
      return { kind, value: tok };
    }

    // Pure 12-digit UPI UTR
    if (/^\d{12}$/.test(tok)) return { kind: "UPI", value: tok };

    // Generic "UTR" / "Nxxxx" / RTGS leading patterns
    if (/^UTR\d{6,}$/.test(tok))    return { kind: "UTR", value: tok };
    if (/^N\d{8,}$/.test(tok))      return { kind: "NEFT", value: tok };
    if (/^RTGS\d{6,}$/.test(tok))   return { kind: "RTGS", value: tok };
    if (/^IMPS\d{6,}$/.test(tok))   return { kind: "IMPS", value: tok };
    if (/^CMS\d{6,}$/.test(tok))    return { kind: "CMS", value: tok };

    // INV prefix → invoice reference (e.g. INV255320402)
    if (/^INV\d{4,}$/.test(tok)) return { kind: "INVOICE", value: tok };
    if (/^BILL\d{4,}$/.test(tok)) return { kind: "INVOICE", value: tok };

    // Cheque-style (CHQ / CHEQUE prefix or pure 6-7 digit when explicitly cheque)
    if (/^CHQ\d{4,}$/.test(tok))    return { kind: "CHEQUE", value: tok };
    if (/^CHEQUE\d{4,}$/.test(tok)) return { kind: "CHEQUE", value: tok };

    // Pure long digit blob 9-18 chars → bare numeric (low strength)
    if (/^\d{9,18}$/.test(tok)) return { kind: "NUMERIC", value: tok };

    return null;
  }

  // ── Pull "labeled" refs out of free text ────────────────────────────────────
  // Captures things like:
  //   utr-12345...  utr no: 12345  ref id - 12345  ref no 12345
  //   txn id 12345  trans no.12345 transaction reference 12345
  //   payment id 12345
  function extractLabeled(text, ownAmount) {
    const out = [];
    if (!text) return out;
    const re =
      /(?:utr|ref(?:erence)?(?:[\s\.\-]*(?:id|no))?|txn(?:[\s\.\-]*(?:id|no))?|trans(?:action)?(?:[\s\.\-]*(?:id|no|ref))?|pay(?:ment)?(?:[\s\.\-]*(?:id|no))?|chq(?:[\s\.\-]*no)?|cheque(?:[\s\.\-]*no)?|neft|rtgs|imps|upi)\s*[\s:\-#\.]\s*([A-Z]{0,5}\d[A-Z0-9\-]{4,30})/gi;
    let m;
    while ((m = re.exec(text)) !== null) {
      // Truncate the captured value at the first `--` divider — Tally narrations
      // commonly use `<UTR>--<AMOUNT>--<NAME>` and the amount/name are NOT
      // part of the reference. Without this, normRef's strip-non-alphanumeric
      // would concatenate the amount onto the UTR (e.g. `UBINH25305820666--568695`
      // → `UBINH25305820666568695`), producing a corrupt 22-char ref that no
      // longer matches the same UTR on the HO side.
      let raw = String(m[1]);
      const dividerIdx = raw.search(/--/);
      if (dividerIdx >= 0) raw = raw.slice(0, dividerIdx);
      const tok = normRef(raw);
      if (!tok || tok.length < 6) continue;
      if (isNoise(tok, ownAmount)) continue;
      const cls = classify(tok);
      if (cls) out.push({ ...cls, raw: m[0] });
      else     out.push({ kind: "LABELED", value: tok, raw: m[0] });
    }
    return out;
  }

  // ── Pull bare/free tokens ───────────────────────────────────────────────────
  function extractFree(text, ownAmount) {
    const out = [];
    if (!text) return out;
    // Split on anything that isn't a-z0-9; keep tokens that are mixed alphanumeric
    const tokens = String(text).toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean);
    for (const t of tokens) {
      if (t.length < 6) continue;
      if (isNoise(t, ownAmount)) continue;
      const cls = classify(t);
      if (cls) out.push({ ...cls, raw: t });
    }
    return out;
  }

  // ── Public: extract ALL refs from one text field ────────────────────────────
  // ownAmount (number, optional) — the entry's own amount. Tokens equal to this
  // amount are filtered out so an amount inside the narration is not misread
  // as a reference.
  function extractRefs(text, ownAmount) {
    if (!text) return [];
    const seen = new Map(); // value → ref (de-dupe; keep highest-strength kind)
    const all = [
      ...extractLabeled(text, ownAmount),
      ...extractFree(text, ownAmount),
    ];
    for (const r of all) {
      const cur = seen.get(r.value);
      if (!cur || (KIND_STRENGTH[r.kind] || 0) > (KIND_STRENGTH[cur.kind] || 0)) {
        seen.set(r.value, r);
      }
    }
    // Sort by strength descending
    return Array.from(seen.values()).map(r => ({
      kind: r.kind,
      value: r.value,
      strength: KIND_STRENGTH[r.kind] || 0,
      raw: r.raw,
    })).sort((a, b) => b.strength - a.strength);
  }

  // ── Pick the single "best" ref string (back-compat for entry.utr) ───────────
  function bestRef(refs) {
    if (!refs || !refs.length) return "";
    return refs[0].value;
  }

  // ── Levenshtein distance (capped O(n·m)) ────────────────────────────────────
  function levenshtein(a, b) {
    a = String(a); b = String(b);
    if (a === b) return 0;
    const al = a.length, bl = b.length;
    if (!al) return bl;
    if (!bl) return al;
    if (Math.abs(al - bl) > 4) return Math.abs(al - bl); // early exit
    const dp = new Array(bl + 1);
    for (let j = 0; j <= bl; j++) dp[j] = j;
    for (let i = 1; i <= al; i++) {
      let prev = dp[0];
      dp[0] = i;
      for (let j = 1; j <= bl; j++) {
        const tmp = dp[j];
        if (a.charCodeAt(i - 1) === b.charCodeAt(j - 1)) {
          dp[j] = prev;
        } else {
          dp[j] = 1 + Math.min(prev, dp[j], dp[j - 1]);
        }
        prev = tmp;
      }
    }
    return dp[bl];
  }

  // ── Jaro-Winkler similarity (0..1) — for fuzzy NAME comparison ──────────────
  function jaroWinkler(a, b) {
    a = String(a || "").toLowerCase();
    b = String(b || "").toLowerCase();
    if (!a.length || !b.length) return 0;
    if (a === b) return 1;
    const range = Math.max(0, Math.floor(Math.max(a.length, b.length) / 2) - 1);
    const aMatches = new Array(a.length).fill(false);
    const bMatches = new Array(b.length).fill(false);
    let matches = 0;
    for (let i = 0; i < a.length; i++) {
      const lo = Math.max(0, i - range);
      const hi = Math.min(b.length - 1, i + range);
      for (let j = lo; j <= hi; j++) {
        if (bMatches[j]) continue;
        if (a[i] !== b[j]) continue;
        aMatches[i] = true;
        bMatches[j] = true;
        matches++;
        break;
      }
    }
    if (!matches) return 0;
    let t = 0, k = 0;
    for (let i = 0; i < a.length; i++) {
      if (!aMatches[i]) continue;
      while (!bMatches[k]) k++;
      if (a[i] !== b[k]) t++;
      k++;
    }
    t /= 2;
    const m = matches;
    const jaro = (m / a.length + m / b.length + (m - t) / m) / 3;
    // Winkler boost for common prefix up to 4 chars
    let l = 0;
    while (l < Math.min(4, a.length, b.length) && a[l] === b[l]) l++;
    return jaro + l * 0.1 * (1 - jaro);
  }

  // ── Score a single pair of refs ─────────────────────────────────────────────
  // Returns {score, method, exact} or null. `exact === true` means the two
  // refs share the same full normalised value (same- or cross-kind) — this is
  // the only situation that should bypass amount corroboration in the
  // higher-level scorer. Everything else (suffix / fuzzy / last-N / bare
  // short numeric) sets exact=false and must be combined with an amount match.
  function scoreSinglePair(ra, rb) {
    if (ra.value === rb.value) {
      const base = Math.min(ra.strength, rb.strength);
      // Bare short NUMERIC matches (<10 digits) are NOT trustworthy on their
      // own — could be an account number, invoice number, anything. Treat as
      // partial so the higher-level scorer requires amount corroboration.
      if (ra.kind === "NUMERIC" && rb.kind === "NUMERIC" && ra.value.length < 10) {
        return { score: 35, method: `Bare#${ra.value.slice(-6)}`, exact: false };
      }
      return { score: Math.min(90, base), method: `${ra.kind} Exact`, exact: true };
    }

    // Different lengths but one ends with the other (suffix match) — common
    // when one side has a prefix the other doesn't.
    const a = ra.value, b = rb.value;
    const minLen = Math.min(a.length, b.length);
    if (minLen >= 6 && (a.endsWith(b) || b.endsWith(a))) {
      const base = Math.min(ra.strength, rb.strength) * 0.7;
      return {
        score: Math.round(base),
        method: `${ra.kind} suffix=${a.slice(-Math.min(8, minLen))}`,
        exact: false,
      };
    }

    // Last-6 / last-8 partial match for same-kind refs
    if (ra.kind === rb.kind && minLen >= 6 && a.slice(-6) === b.slice(-6)) {
      const last8 = minLen >= 8 && a.slice(-8) === b.slice(-8);
      const score = last8 ? 45 : 28;
      return { score, method: `${ra.kind} last-${last8 ? 8 : 6}`, exact: false };
    }

    // Cross-kind partial match for IFSC bank-prefixed refs:
    // The same UTR is sometimes booked as different kinds on the two sides
    // (NEFT / RTGS / IMPS — H/R/I hint letters differ but the trailing UTR
    // sequence is identical). Require both refs to start with the SAME 4-letter
    // bank prefix from our known list, and the last 6+ digits to match.
    const aPref = a.match(/^([A-Z]{4})/);
    const bPref = b.match(/^([A-Z]{4})/);
    if (aPref && bPref && aPref[1] === bPref[1] && BANK_PREFIXES.has(aPref[1])
        && ra.kind !== rb.kind && minLen >= 8) {
      const aDigits = a.replace(/[^0-9]/g, "");
      const bDigits = b.replace(/[^0-9]/g, "");
      if (aDigits.length >= 6 && bDigits.length >= 6 && aDigits.slice(-6) === bDigits.slice(-6)) {
        const last8 = aDigits.length >= 8 && bDigits.length >= 8 && aDigits.slice(-8) === bDigits.slice(-8);
        return {
          score: last8 ? 50 : 32,
          method: `${aPref[1]} cross-kind last-${last8 ? 8 : 6}`,
          exact: false,
        };
      }
    }

    // Levenshtein fuzzy on long, strong refs (catches OCR / 1-2 char typos)
    if (ra.strength >= 70 && rb.strength >= 70 && Math.min(a.length, b.length) >= 10) {
      const d = levenshtein(a, b);
      if (d === 1) return { score: 65, method: `${ra.kind} ~1typo`, exact: false };
      if (d === 2) return { score: 45, method: `${ra.kind} ~2typo`, exact: false };
    }

    return null;
  }

  // ── Score TWO ref arrays (best pairing wins, multi-match boosts) ────────────
  // Returns { score, method, exact, exactCount, a, b } or null.
  //   exact      — true iff the chosen best pair is a full-value match
  //   exactCount — how many independent exact matches were found between the
  //                two arrays (powers a multi-ref boost)
  function refsMatchScore(refsA, refsB) {
    if (!refsA || !refsB || !refsA.length || !refsB.length) return null;
    let best = null;
    let matchCount = 0;   // any kind of match (incl. partial)
    let exactCount = 0;   // strict exact-value matches only
    for (const ra of refsA) {
      for (const rb of refsB) {
        const r = scoreSinglePair(ra, rb);
        if (!r) continue;
        if (r.score >= 25) matchCount++;
        if (r.exact) exactCount++;
        // Prefer exact pairings over equally-scoring partials. We do this by
        // adding a tiny bias when comparing — this keeps the existing score
        // values clean while ensuring `best.exact` reflects the strongest
        // signal available.
        const cmp = r.score + (r.exact ? 0.5 : 0);
        const bestCmp = best ? best.score + (best.exact ? 0.5 : 0) : -1;
        if (!best || cmp > bestCmp) best = { ...r, a: ra, b: rb };
      }
    }
    if (!best) return null;
    // Multi-ref boost — only when 2+ EXACT matches (suffix/fuzzy multi-matches
    // are not trustworthy enough to compound).
    if (exactCount >= 2) {
      best.score = Math.min(95, best.score + Math.min(10, (exactCount - 1) * 5));
      best.method += ` (+${exactCount - 1} exact)`;
    }
    best.exactCount = exactCount;
    return best;
  }

  // ── Public API ──────────────────────────────────────────────────────────────
  const api = {
    extractRefs,
    bestRef,
    refsMatchScore,
    scoreSinglePair,
    levenshtein,
    jaroWinkler,
    normRef,
    classify,
    BANK_PREFIXES,
    KIND_STRENGTH,
    isNoise,
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  } else {
    Object.assign(global, api);
    global.RefsLib = api;
  }
})(typeof window !== "undefined" ? window : globalThis);
