"use strict";
// ═══════════════════════════════════════════════════════════════════════════════
// RECONCILIATION ENGINE  v3 — trust-first design
// ═══════════════════════════════════════════════════════════════════════════════
//
// Confidence tiers:
//   95-100  AUTO-RECONCILED     (high trust, zero manual check needed)
//   80-94   RECONCILED          (good match, spot-check recommended)
//   70-79   PROBABLE            (shown in results, user confirms)
//   50-69   MANUAL REVIEW       (never auto-booked)
//   <50     NO MATCH            (goes to suspense/wrong-ledger path)
//
// Key design decisions:
//   - Partial UTR (last-4, last-5) NEVER used alone — always needs amount corroboration
//   - Amount + Date is the PRIMARY non-UTR signal
//   - Amount uniqueness is computed across the FULL dataset — a ₹60,000 that appears
//     5 times is treated as a round/common amount even if not technically round
//   - Multiple candidates at similar confidence → CONFLICT, sent to manual review
//   - Minimum gap between best and second-best must be ≥15 pts to auto-classify
// ═══════════════════════════════════════════════════════════════════════════════

// ── Round / common amount detection ──────────────────────────────────────────
const ALWAYS_ROUND = new Set([
  100,200,250,300,400,500,600,700,750,800,900,
  1000,1250,1500,2000,2500,3000,4000,5000,7500,
  10000,12500,15000,20000,25000,30000,40000,50000,
  75000,100000,125000,150000,200000,250000,500000,1000000
]);

function buildAmountFrequencyMap(lists) {
  // Count how many times each amount appears across all entry lists combined
  const freq = new Map();
  for (const list of lists) {
    for (const e of list) {
      const a = Math.round(parseFloat(e.amount) * 100); // paise to avoid float issues
      freq.set(a, (freq.get(a) || 0) + 1);
    }
  }
  return freq;
}

function isCommonAmount(amount, freqMap) {
  const a = Math.round(parseFloat(amount) * 100);
  if (ALWAYS_ROUND.has(parseFloat(amount))) return true;
  if (parseFloat(amount) >= 1000 && parseFloat(amount) % 1000 === 0) return true;
  if (parseFloat(amount) >= 100  && parseFloat(amount) % 100  === 0) return true;
  // If this exact amount appears 3+ times across all ledgers, treat as non-unique
  if (freqMap && (freqMap.get(a) || 0) >= 3) return true;
  return false;
}

// ── UTR utilities ─────────────────────────────────────────────────────────────
function normUTR(s) { return String(s||"").replace(/[\s\-]/g,"").toUpperCase(); }

function classifyUTR(u) {
  const s = normUTR(u);
  if (!s) return "none";
  if (/^\d{12}$/.test(s))     return "upi";
  if (/^N\d|NEFT/i.test(s))   return "neft";
  if (/^IMPS/i.test(s))       return "imps";
  if (/^RTGS|^UT\d/i.test(s)) return "rtgs";
  if (/^CMS/i.test(s))        return "cms";
  return "other";
}

// Score UTR pair — returns { score, method } or null if no UTR signal
function scoreUTR(aUTR, bUTR) {
  const a = normUTR(aUTR), b = normUTR(bUTR);
  if (!a || !b) return null;
  if (a === b) return { score: 60, method: "UTR Exact" }; // base — boosted by other signals below

  const aType = classifyUTR(a);

  if (aType === "upi") {
    // 4-4-4 chunk matching: all 3 chunks match = very strong partial
    const ch = s => [s.slice(0,4), s.slice(4,8), s.slice(8,12)];
    const ac = ch(a), bc = ch(b);
    const hits = ac.filter((x,j) => x && bc[j] && x === bc[j]).length;
    if (hits === 3) return { score: 55, method: "UPI 4-4-4 (3/3)" }; // near-exact
    if (hits === 2) return { score: 35, method: "UPI 4-4-4 (2/3)" }; // needs corroboration
    // Last-4 only: too weak for independent signal, return low score
    if (a.slice(-4) === b.slice(-4)) return { score: 20, method: "UPI Last-4 only" };
  } else if (aType !== "none") {
    // NEFT/IMPS/RTGS/CMS last-5: weak alone
    if (a.length >= 5 && b.length >= 5 && a.slice(-5) === b.slice(-5))
      return { score: 22, method: `${aType.toUpperCase()} Last-5 only` };
  }
  return null;
}

// ── Date utilities ────────────────────────────────────────────────────────────
function toDate(d) {
  if (!d) return null;
  const s = String(d).replace(/[\/\-\.]/g,"");
  if (s.length !== 8) return null;
  let y, m, dd;
  if (parseInt(s.slice(0,4)) > 1900) { y=s.slice(0,4); m=s.slice(4,6); dd=s.slice(6,8); }
  else                               { dd=s.slice(0,2); m=s.slice(2,4); y=s.slice(4,8); }
  const dt = new Date(`${y}-${m}-${dd}`);
  return isNaN(dt.getTime()) ? null : dt;
}

function daysDiff(a, b) {
  const da = toDate(a), db = toDate(b);
  if (!da || !db) return 999;
  return Math.abs((da - db) / 86400000);
}

// ── Amount match ──────────────────────────────────────────────────────────────
function amtEq(a, b) { return Math.abs(parseFloat(a) - parseFloat(b)) < 0.5; }

// ── Narration token similarity ────────────────────────────────────────────────
// Focuses on numeric tokens (account numbers, amounts, refs) which are more
// reliable than name words — especially since Branch=customer name, HO=bank name
function narrSim(a, b) {
  if (!a && !b) return 0;
  const tokens = s => String(s||"").toLowerCase()
    .split(/[\s\/\-_,\(\)]+/)
    .filter(t => t.length >= 3);
  const ta = new Set(tokens(a)), tb = new Set(tokens(b));
  if (!ta.size || !tb.size) return 0;
  let inter = 0;
  ta.forEach(t => { if (tb.has(t)) inter++; });
  return inter / Math.min(ta.size, tb.size);
}

// Specifically look for shared numeric sequences (UTRs, account numbers)
// that appear in both narrations — very strong signal
function sharedNumericRef(narrA, narrB) {
  const nums = s => (String(s||"").match(/\d{6,}/g) || []).map(n => n.replace(/^0+/,""));
  const na = new Set(nums(narrA)), nb = new Set(nums(narrB));
  for (const n of na) { if (nb.has(n) && n.length >= 6) return n; }
  return null;
}

// ═══════════════════════════════════════════════════════════════════════════════
// CORE SCORER
// Scores a single candidate pair (src, candidate).
// Returns { score, method, signals[] } — score 0 means no meaningful match.
// ═══════════════════════════════════════════════════════════════════════════════
function scorePair(src, cand, freqMap) {
  let score = 0;
  const signals = [];
  let method = "";

  const amountMatches = amtEq(src.amount, cand.amount);
  const diff = daysDiff(src.date, cand.date);
  const dateKnown = diff < 999;
  const sameDay   = diff === 0;
  const nearDay   = diff <= 2;

  // ── Signal A: UTR ──────────────────────────────────────────────────────────
  const utrResult = scoreUTR(src.utr, cand.utr);
  let utrScore = 0;
  if (utrResult) {
    utrScore = utrResult.score;
    signals.push(utrResult.method);
  }

  // ── Signal B: Shared numeric ref in narrations ─────────────────────────────
  // (catches cases where UTR format differs payer vs receiver side)
  const sharedRef = sharedNumericRef(src.narration, cand.narration);
  let sharedRefScore = 0;
  if (sharedRef && sharedRef.length >= 9) {
    sharedRefScore = sharedRef.length >= 12 ? 55 : 40;
    signals.push(`SharedRef(${sharedRef.slice(-6)})`);
  }

  // Take best of UTR match and shared narration ref
  const utrSignal = Math.max(utrScore, sharedRefScore);

  // ── Signal C: Amount ───────────────────────────────────────────────────────
  let amtScore = 0;
  if (amountMatches) {
    const common = isCommonAmount(src.amount, freqMap);
    amtScore = common ? 20 : 35; // unique amounts worth more
    signals.push(common ? "Amt(common)" : "Amt(unique)");
  }

  // ── Signal D: Date ─────────────────────────────────────────────────────────
  let dateScore = 0;
  if (amountMatches && dateKnown) {
    if (sameDay)       { dateScore = 30; signals.push("Date=exact"); }
    else if (diff===1) { dateScore = 22; signals.push("Date±1d"); }
    else if (diff===2) { dateScore = 14; signals.push("Date±2d"); }
    else if (diff<=5)  { dateScore = 6;  signals.push(`Date±${diff}d`); }
  }

  // ── Signal E: Narration similarity (secondary, not primary) ────────────────
  let narrScore = 0;
  if (amountMatches) {
    const ns = narrSim(src.narration, cand.narration);
    if (ns >= 0.6)      { narrScore = 12; signals.push(`Narr(${Math.round(ns*100)}%)`); }
    else if (ns >= 0.35){ narrScore = 6;  signals.push(`Narr(${Math.round(ns*100)}%)`); }
  }

  // ── Combine signals ────────────────────────────────────────────────────────
  // The key rule: partial UTR alone (score < 40) ONLY counts if amount also matches
  // Full UTR exact gets massive boost from any corroborating signal

  if (utrSignal >= 60) {
    // Full UTR exact match — primary anchor, boost with corroborating signals
    score = utrSignal + amtScore + dateScore + narrScore;
    // Cap at 100, but exact UTR + amount + date can reach ~100 naturally
    score = Math.min(100, score);
    method = utrResult?.method === "UTR Exact" ? "UTR Exact" : signals.join(" + ");
  } else if (utrSignal >= 35 && amountMatches) {
    // Strong partial UTR (3/3 chunks or long shared ref) + amount = trustworthy
    score = utrSignal + amtScore + dateScore + narrScore;
    method = signals.join(" + ");
  } else if (utrSignal > 0 && amountMatches) {
    // Weak partial UTR (last-4, last-5, 2/3 chunks) — only worth something WITH amount
    score = utrSignal + amtScore + dateScore;
    method = signals.join(" + ") + " ⚠partial-UTR";
  } else if (!utrSignal && amountMatches) {
    // No UTR signal at all — rely purely on amount + date + narration
    score = amtScore + dateScore + narrScore;
    method = signals.join(" + ");
  }
  // If no amount match and no full UTR: score stays 0

  return { score: Math.round(score), method, signals };
}

// ═══════════════════════════════════════════════════════════════════════════════
// FIND BEST MATCH — with conflict detection
// Returns { entry, idx, score, method, conflicted } or null
// ═══════════════════════════════════════════════════════════════════════════════
function findBestMatch(src, candidates, usedSet, freqMap) {
  const scored = [];

  for (let i = 0; i < candidates.length; i++) {
    if (usedSet.has(i)) continue;
    const result = scorePair(src, candidates[i], freqMap);
    if (result.score > 0) {
      scored.push({ idx: i, entry: candidates[i], ...result });
    }
  }

  if (!scored.length) return null;

  // Sort by score descending
  scored.sort((a, b) => b.score - a.score);

  const best = scored[0];
  const second = scored[1];

  // Conflict detection: if second candidate is within 15 pts of best,
  // the match is ambiguous — flag it
  const conflicted = second && (best.score - second.score) < 15 && best.score < 90;

  return { ...best, conflicted, alternatives: conflicted ? scored.slice(1, 3) : [] };
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN RECONCILIATION FUNCTION
// ═══════════════════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════════════════
// BILLS TRANSFER LAYER
// Handles "Transfer to HO Bills" entries — invoice-based, no UTR
// Branch narration: "to ho", "ho transfer", "ref mallikarjun", "ref sandeep" etc.
// HO narration: "inv.255320402", "inv 255320473" etc.
// Party in HO: customer name (Sree Durga Earth Movers)
// ═══════════════════════════════════════════════════════════════════════════════

// Configurable — updated from UI state before reconcile runs
let BILLS_AUTHORISER_NAMES = ["mallikarjun", "malli", "sandeep"];
let BILLS_KEYWORDS = ["to ho", "ho transfer", "transfer to ho", "bill transfer", "ho bill"];

function isBillsTransferEntry(entry) {
  const narr = String(entry.narration || "").toLowerCase();
  const party = String(entry.party || "").toLowerCase();
  const combined = narr + " " + party;

  // Keyword match
  if (BILLS_KEYWORDS.some(k => combined.includes(k))) return true;

  // Authoriser name match — only valid when preceded by "ref" or "ref." or standalone
  // "sir" alone is NOT used — only with ref keyword context
  for (const name of BILLS_AUTHORISER_NAMES) {
    // "ref mallikarjun", "ref. malli", "ref:malli"
    if (new RegExp(`ref[\\s\\.\\:]*${name}`,"i").test(combined)) return true;
    // Name alone as full word (not "sir")
    if (name !== "sir" && name.length >= 5) {
      if (new RegExp(`\\b${name}`,"i").test(combined)) return true;
    }
  }

  // HO side: has invoice reference — also qualifies as bills type
  if (/inv[\s\.\-]*\d{5,}/i.test(combined)) return true;

  return false;
}

function extractInvoiceNumber(text) {
  if (!text) return null;
  // "inv.255320402", "inv 255320402", "inv-255320402", "invoice no 255320402"
  const m = String(text).match(/inv(?:oice)?[\s\.\-\:no#]*(\d{5,})/i);
  if (m) return m[1];
  return null;
}

// Score a bills-transfer pair: Branch entry vs HO entry
function scoreBillsPair(br, ho, freqMap) {
  let score = 0;
  const signals = [];

  // Amount must match (±₹1) — non-negotiable for bills transfers
  if (!amtEq(br.amount, ho.amount)) return { score: 0, method: "", signals: [] };
  score += 35;
  signals.push("Amt");

  // Invoice number match — primary anchor
  const brInv = extractInvoiceNumber(br.narration) || extractInvoiceNumber(br.party);
  const hoInv = extractInvoiceNumber(ho.narration) || extractInvoiceNumber(ho.party);
  if (brInv && hoInv && brInv === hoInv) {
    score += 55; signals.push(`Inv#${brInv}`);
  } else if (brInv && hoInv && brInv.slice(-5) === hoInv.slice(-5)) {
    // Last 5 digits of invoice number — suffix sometimes differs by prefix
    score += 35; signals.push(`Inv~${brInv.slice(-5)}`);
  }

  // Customer name in HO matches branch party name
  const nameSim = narrSim(ho.party, br.party);
  if (nameSim >= 0.6) { score += 20; signals.push(`Name(${Math.round(nameSim*100)}%)`); }
  else if (nameSim >= 0.35) { score += 10; signals.push(`Name~`); }

  // Also check HO narration vs branch party (HO narration sometimes has customer name)
  const narrNameSim = narrSim(ho.narration, br.party);
  if (narrNameSim >= 0.5 && score < 90) { score += 12; signals.push(`NarrName`); }

  // Date within 7 days (bills transfer booking lag)
  const diff = daysDiff(br.date, ho.date);
  if (diff === 0)      { score += 15; signals.push("Date=exact"); }
  else if (diff <= 2)  { score += 10; signals.push(`Date±${diff}d`); }
  else if (diff <= 7)  { score += 5;  signals.push(`Date±${diff}d`); }

  return { score: Math.min(100, Math.round(score)), method: "BILLS: " + signals.join("+"), signals };
}

// Match bills-transfer entries as a separate pool before normal reconciliation
function reconcileBillsTransfers(branchEntries, hoEntries, freqMap) {
  const billsBranch = branchEntries.map((e,i) => ({ e, i, bills: isBillsTransferEntry(e) })).filter(x => x.bills);
  const billsHO     = hoEntries.map((e,i) => ({ e, i, bills: isBillsTransferEntry(e) })).filter(x => x.bills);

  const usedBranchIdx = new Set();
  const usedHOIdx     = new Set();
  const matched       = [];
  const unmatchedBr   = [];
  const unmatchedHO   = [];

  // Match each bills-branch entry
  for (const { e: br, i: bi } of billsBranch) {
    let best = null, bestScore = 0, bestMethod = "";

    for (const { e: ho, i: hi } of billsHO) {
      if (usedHOIdx.has(hi)) continue;
      const { score, method } = scoreBillsPair(br, ho, freqMap);
      if (score > bestScore) { bestScore = score; best = { ho, hi }; bestMethod = method; }
    }

    if (best && bestScore >= 80) {
      usedBranchIdx.add(bi);
      usedHOIdx.add(best.hi);
      matched.push({ branchIdx: bi, hoIdx: best.hi,
        branch: br, ho: best.ho, score: bestScore, method: bestMethod,
        net: Math.abs(parseFloat(br.amount) - parseFloat(best.ho.amount)).toFixed(2),
        autoConfirmed: bestScore >= 90, isBillsTransfer: true });
    } else if (best && bestScore >= 55) {
      usedBranchIdx.add(bi);
      // Goes to manual review — returned as unmatched with best candidate attached
      unmatchedBr.push({ branchIdx: bi, branch: br,
        candidate: best.ho, candidateScore: bestScore, candidateMethod: bestMethod });
    } else {
      usedBranchIdx.add(bi);
      unmatchedBr.push({ branchIdx: bi, branch: br, candidate: null, candidateScore: 0 });
    }
  }

  // HO bills entries with no match
  for (const { e: ho, i: hi } of billsHO) {
    if (!usedHOIdx.has(hi)) unmatchedHO.push({ hoIdx: hi, ho });
  }

  return { matched, unmatchedBr, unmatchedHO,
    usedBranchIdxSet: usedBranchIdx, usedHOIdxSet: usedHOIdx };
}

function reconcile(branchEntries, hoEntries, suspenseEntries) {
  const reconciled   = [];
  const toSuspense   = [];
  const fromSuspense = [];
  const wrongLedger  = [];
  const manualReview = [];

  const usedHO   = new Set();
  const usedSusp = new Set();

  // Build amount frequency map across all three ledgers
  const freqMap = buildAmountFrequencyMap([branchEntries, hoEntries, suspenseEntries]);

  // ── Step 0: Bills Transfer pre-pass ─────────────────────────────────────────
  const bills = reconcileBillsTransfers(branchEntries, hoEntries, freqMap);

  // Add confirmed bills matches to reconciled
  for (const m of bills.matched) {
    usedHO.add(m.hoIdx);
    reconciled.push(m);
  }
  // Bills manual review candidates
  for (const u of bills.unmatchedBr) {
    if (u.candidate) {
      manualReview.push({
        source: "Bills↔HO", branch: u.branch, ho: u.candidate,
        score: u.candidateScore, method: u.candidateMethod,
        reason: "Bills transfer — low confidence invoice/name match, verify",
        alternatives: [], isBillsTransfer: true,
      });
    } else {
      wrongLedger.push({
        branch: u.branch,
        reason: "Bills transfer entry — no matching HO credit found",
        suggestion: "Check if HO has booked against customer or different branch",
        isBillsTransfer: true,
      });
    }
  }
  // HO bills entries with no branch match → suspense
  for (const u of bills.unmatchedHO) {
    usedHO.add(u.hoIdx);
    toSuspense.push({
      ho: u.ho,
      action: "Rebook: Bills HO Credit → Suspense",
      reason: "HO has bills credit but no matching Branch debit found",
      isBillsTransfer: true,
      tallyEntry: {
        date: u.ho.date, fromLedger: "Branch – Credit to Branch",
        toLedger: "Suspense Account", amount: u.ho.amount,
        narration: `Bills recon: suspense - ${u.ho.narration?.slice(0,40) || u.ho.amount}`,
        utr: u.ho.utr,
      },
    });
  }

  // Mark all bills-classified branch entries as used so normal engine skips them
  const usedBranch = bills.usedBranchIdxSet;

  // ── Step 1: Match each Branch entry against HO entries (skip bills entries) ──
  for (let bi = 0; bi < branchEntries.length; bi++) {
    if (usedBranch.has(bi)) continue; // already handled by bills layer
    const br = branchEntries[bi];
    const m = findBestMatch(br, hoEntries, usedHO, freqMap);

    // Determine tier
    const score = m ? m.score : 0;
    const tier =
      score >= 95 ? "auto" :
      score >= 80 ? "reconciled" :
      score >= 70 ? "probable" :
      score >= 50 ? "review" : "none";

    if (m && (tier === "auto" || tier === "reconciled") && !m.conflicted) {
      usedHO.add(m.idx);
      const net = Math.abs(parseFloat(br.amount) - parseFloat(m.entry.amount));
      reconciled.push({
        branch: br, ho: m.entry,
        score: m.score, method: m.method,
        net: net.toFixed(2),
        autoConfirmed: tier === "auto",
      });

    } else if (m && (tier === "probable" || tier === "review" || m.conflicted)) {
      // Medium confidence or conflicted — manual review
      let reason = "";
      if (m.conflicted)    reason = `Conflicted — ${m.alternatives.length+1} candidates within 15 pts`;
      else if (tier === "review")   reason = "Low confidence — verify date and narration";
      else                          reason = "Probable match — confirm before booking";

      manualReview.push({
        source: "Branch↔HO",
        branch: br, ho: m.entry,
        score: m.score, method: m.method,
        reason,
        alternatives: m.alternatives || [],
      });

    } else {
      // No HO match — check Suspense with enhanced customer name matching
      const sm = findBestMatch(br, suspenseEntries, usedSusp, freqMap);

      // Also try customer name match against suspense narration
      // (payments booked in suspense sometimes use customer name in narration)
      let bestSusp = sm;
      if (!sm || sm.score < 80) {
        for (let si = 0; si < suspenseEntries.length; si++) {
          if (usedSusp.has(si)) continue;
          const se = suspenseEntries[si];
          if (!amtEq(br.amount, se.amount)) continue;
          // Check branch party name against suspense narration
          const nameInNarr = narrSim(br.party, se.narration);
          const nameInParty = narrSim(br.party, se.party);
          const bestNameSim = Math.max(nameInNarr, nameInParty);
          if (bestNameSim >= 0.45) {
            const diff = daysDiff(br.date, se.date);
            let nameScore = 35 + bestNameSim * 30; // 35–65 base
            if (diff === 0)     nameScore += 20;
            else if (diff <= 2) nameScore += 12;
            else if (diff <= 7) nameScore += 5;
            nameScore = Math.round(Math.min(nameScore, 85));
            const nameMethod = `Amt+CustName(${Math.round(bestNameSim*100)}%)${nameInNarr>nameInParty?"+Narr":""}`;
            if (!bestSusp || nameScore > bestSusp.score) {
              bestSusp = { entry: se, idx: si, score: nameScore, method: nameMethod, conflicted: false, alternatives: [] };
            }
          }
        }
      }
      const sScore = bestSusp ? bestSusp.score : 0;

      if (bestSusp && sScore >= 70 && !bestSusp.conflicted) {
        usedSusp.add(bestSusp.idx);
        fromSuspense.push({
          branch: br, suspense: bestSusp.entry,
          score: bestSusp.score, method: bestSusp.method,
          action: "Rebook: Suspense → Branch (Debit to HO)",
          tallyEntry: {
            date:       bestSusp.entry.date || br.date,
            fromLedger: "Suspense Account",
            toLedger:   "Branch – Debit to HO",
            amount:     br.amount,
            narration:  `Recon adj: ${br.utr || br.narration?.slice(0,40) || br.amount}`,
            utr:        br.utr || bestSusp.entry.utr,
          },
        });
      } else if (bestSusp && sScore >= 50) {
        // Low-confidence suspense match → manual
        manualReview.push({
          source: "Branch↔Suspense",
          branch: br, suspense: bestSusp.entry,
          score: bestSusp.score, method: bestSusp.method,
          reason: "Possible suspense match — low confidence, verify manually",
          alternatives: bestSusp.alternatives || [],
        });
      } else {
        wrongLedger.push({
          branch: br,
          reason: "Not found in HO ledger or Suspense",
          suggestion: isCommonAmount(br.amount, freqMap)
            ? "Common/round amount — check if booked to different branch or wrong ledger"
            : "Unique amount — likely wrong ledger or branch, investigate urgently",
        });
      }
    }
  }

  // ── Step 2: HO entries with no Branch match → park in Suspense ──────────────
  hoEntries.forEach((ho, i) => {
    if (usedHO.has(i)) return;
    const inReview = manualReview.some(r => r.ho === ho);
    if (inReview) return;

    toSuspense.push({
      ho,
      action: "Rebook: Credit to Branch → Suspense",
      reason: "HO has this credit but no matching Branch debit found",
      tallyEntry: {
        date:       ho.date,
        fromLedger: "Branch – Credit to Branch",
        toLedger:   "Suspense Account",
        amount:     ho.amount,
        narration:  `Recon: Park in suspense - ${ho.utr || ho.narration?.slice(0,40) || ho.amount}`,
        utr:        ho.utr,
      },
    });
  });

  // ── Step 3: Unmatched suspense entries ──────────────────────────────────────
  suspenseEntries.forEach((s, i) => {
    if (usedSusp.has(i)) return;
    manualReview.push({
      source: "Suspense-Unmatched",
      suspense: s,
      reason: "Suspense entry — no matching Branch debit found, verify manually",
    });
  });

  return { reconciled, toSuspense, fromSuspense, wrongLedger, manualReview };
}

// ── FY month utilities ────────────────────────────────────────────────────────
// Indian FY: Apr=1 ... Mar=12
function fyMonth(dateStr) {
  const d = toDate(dateStr);
  if (!d) return null;
  const m = d.getMonth() + 1;
  const y = d.getFullYear();
  const fyM = m >= 4 ? m - 3 : m + 9;
  const fy  = m >= 4 ? `${y}-${String(y+1).slice(-2)}` : `${y-1}-${String(y).slice(-2)}`;
  const mon = d.toLocaleString("en-IN", { month: "short" });
  return { fy, fyMonth: fyM, monthName: `${mon} ${m >= 4 ? y : y}` };
}

function sameFYMonth(a, b) {
  const fa = fyMonth(a), fb = fyMonth(b);
  if (!fa || !fb) return true; // unknown dates — don't flag
  return fa.fy === fb.fy && fa.fyMonth === fb.fyMonth;
}

// ── Split reconciled → clean + dateMismatch ───────────────────────────────────
function splitDateMismatches(results) {
  const clean = [];
  const dateMismatch = [];
  for (const r of results.reconciled) {
    const receiptDate = r.ho?.date;
    const journalDate = r.branch?.date;
    if (receiptDate && journalDate && !sameFYMonth(journalDate, receiptDate)) {
      const hoFY = fyMonth(receiptDate);
      const brFY = fyMonth(journalDate);
      dateMismatch.push({
        ...r,
        receiptDate,
        journalDate,
        receiptFYMonth: hoFY?.monthName || receiptDate,
        journalFYMonth: brFY?.monthName || journalDate,
        approved: false,
        tallyAlterEntry: {
          voucherRef:  r.branch?.ref || "",
          partyLedger: r.branch?.party || "",
          newDate:     receiptDate,
          amount:      r.branch?.amount,
          narration:   r.branch?.narration || "",
        },
      });
    } else {
      clean.push(r);
    }
  }
  return { ...results, reconciled: clean, dateMismatch };
}

// ── Bank ledger cross-check ───────────────────────────────────────────────────
// Run AFTER reconcile(). Takes wrongLedger + manualReview entries and looks for
// where the payment was actually booked in the Tally bank account ledger.
// Annotates each entry with bankHint: { party, date, amount, narration }
function crossCheckBankLedger(results, bankEntries) {
  if (!bankEntries || !bankEntries.length) return results;

  const annotate = (entry) => {
    const br = entry.branch || entry.ho || entry.suspense;
    if (!br) return entry;

    // Find bank ledger entry matching by amount + date (±3 days) or UTR
    let bestHint = null, bestScore = 0;
    for (const be of bankEntries) {
      if (!amtEq(br.amount, be.amount)) continue;
      let score = 30;
      const diff = daysDiff(br.date, be.date);
      if (diff === 0)     score += 30;
      else if (diff <= 2) score += 20;
      else if (diff <= 5) score += 10;
      else continue; // too far apart for amount-only

      // UTR match boosts significantly
      if (br.utr && be.utr && normUTR(br.utr) === normUTR(be.utr)) score += 40;
      else {
        const shared = sharedNumericRef(br.narration, be.narration);
        if (shared) score += 25;
      }

      if (score > bestScore) { bestScore = score; bestHint = be; }
    }

    if (bestHint && bestScore >= 50) {
      return { ...entry, bankHint: {
        party:     bestHint.party,
        date:      bestHint.date,
        amount:    bestHint.amount,
        narration: bestHint.narration,
        score:     bestScore,
      }};
    }
    return entry;
  };

  return {
    ...results,
    wrongLedger:  results.wrongLedger.map(annotate),
    manualReview: results.manualReview.map(annotate),
    toSuspense:   results.toSuspense.map(annotate),
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════════════════════════════════════
const S = {
  tab: "upload",
  branch: { files: [], entries: [] },
  ho:     { files: [], entries: [] },
  susp:   { files: [], entries: [] },
  bank:   { files: [], entries: [] },   // Bank account ledger from Tally (optional)
  branchLedgerName:   "Debit to HO",
  hoLedgerName:       "Credit to Branch",
  suspenseLedgerName: "Suspense Account",
  billsAuthorisers: "mallikarjun, sandeep",
  billsKeywords:    "to ho, ho transfer, transfer to ho",
  results: null,
  resultsTab: "reconciled",
  statusMsg: "",
  processing: false,
  reviewedSet: new Set(),
  searchQuery: "",
  sortBy: null, sortDir: 1,
};

// ═══════════════════════════════════════════════════════════════════════════════
// DOM HELPERS
// ═══════════════════════════════════════════════════════════════════════════════
const ROOT = document.getElementById("root");
const E = (tag, props={}, ...kids) => {
  const e = document.createElement(tag);
  for (const [k,v] of Object.entries(props)) {
    if (k==="style" && typeof v==="object") Object.assign(e.style, v);
    else if (k.startsWith("on"))           e.addEventListener(k.slice(2), v);
    else if (k==="class")                  e.className = v;
    else                                   e.setAttribute(k, v);
  }
  kids.forEach(c => { if (c==null) return; e.appendChild(typeof c==="string" ? document.createTextNode(c) : c); });
  return e;
};
const DIV  = (p,...k) => E("div",p,...k);
const SPAN = (p,...k) => E("span",p,...k);
const BTN  = (p,...k) => E("button",p,...k);

// ═══════════════════════════════════════════════════════════════════════════════
// RENDER
// ═══════════════════════════════════════════════════════════════════════════════
function render() {
  ROOT.innerHTML = "";

  // ── Titlebar ────────────────────────────────────────────────────────────────
  const tb = DIV({style:{display:"flex",alignItems:"center",justifyContent:"space-between",
    height:"38px",background:"#040509",padding:"0 14px",
    borderBottom:"1px solid #14151f",WebkitAppRegion:"drag",flexShrink:"0"}});

  const logo = DIV({style:{display:"flex",alignItems:"center",gap:"9px"}});
  const badge = DIV({style:{width:"26px",height:"26px",borderRadius:"6px",
    background:"linear-gradient(135deg,#3a7bd5,#1a4a8a)",
    display:"flex",alignItems:"center",justifyContent:"center",fontSize:"14px"}});
  badge.textContent = "⇄";
  const title = SPAN({style:{fontSize:"11px",fontFamily:"'IBM Plex Mono',monospace",
    fontWeight:"600",color:"#6070a0",letterSpacing:"2px"}});
  title.textContent = "HO–BRANCH RECONCILIATION";
  logo.append(badge, title);

  const wc = DIV({style:{display:"flex",gap:"5px",WebkitAppRegion:"no-drag"}});
  [["–","#f97316","minimize"],["□","#fbbf24","maximize"],["✕","#ef4444","close"]].forEach(([t,c,a])=>{
    const b = BTN({style:{background:"transparent",border:"none",color:c,cursor:"pointer",
      width:"26px",height:"26px",borderRadius:"5px",fontSize:"12px",
      fontFamily:"'IBM Plex Mono',monospace",transition:"background 0.15s"},
      onmouseenter:()=>b.style.background="#1a1b2e",
      onmouseleave:()=>b.style.background="transparent",
      onclick:()=>window.api[a]()});
    b.textContent = t;
    wc.appendChild(b);
  });
  tb.append(logo, wc);

  // ── Tab bar ─────────────────────────────────────────────────────────────────
  const tabs = DIV({style:{display:"flex",borderBottom:"1px solid #13141e",
    background:"#040509",padding:"0 18px",flexShrink:"0"}});

  const tabDefs = [
    ["upload","① UPLOAD & RUN"],
    ["results","② RESULTS" + (S.results ? ` (${(S.results.reconciled||[]).length+
      (S.results.dateMismatch||[]).length+
      (S.results.toSuspense||[]).length+(S.results.fromSuspense||[]).length+
      (S.results.wrongLedger||[]).length+(S.results.manualReview||[]).length})` : "")],
    ["guide","GUIDE"],
  ];
  tabDefs.forEach(([id, label]) => {
    const t = BTN({
      style:{padding:"10px 18px",background:"transparent",border:"none",
        borderBottom: S.tab===id ? "2px solid #3a7bd5" : "2px solid transparent",
        color: S.tab===id ? "#7aa8e0" : "#3a3d52",
        cursor:"pointer",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",
        fontWeight:"600",letterSpacing:"1.5px",transition:"all 0.15s"},
      onclick:()=>{ if(id==="results"&&!S.results)return; S.tab=id; render(); },
      onmouseenter:(ev)=>{ if(S.tab!==id) ev.target.style.color="#6070a0"; },
      onmouseleave:(ev)=>{ if(S.tab!==id) ev.target.style.color="#3a3d52"; },
    });
    t.textContent = label;
    tabs.appendChild(t);
  });

  // ── Main ─────────────────────────────────────────────────────────────────────
  const main = DIV({style:{flex:"1",overflowY:"auto",padding:"22px 24px"}});

  if (S.tab === "upload")  renderUploadTab(main);
  if (S.tab === "results") renderResultsTab(main);
  if (S.tab === "guide")   renderGuideTab(main);

  ROOT.append(tb, tabs, main);
}

// ─── UPLOAD TAB ───────────────────────────────────────────────────────────────
function renderUploadTab(container) {
  const grid = DIV({style:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"14px",marginBottom:"20px"}});

  grid.appendChild(makeUploadCard("BRANCH LEDGER","Debit to HO","📘","#3a7bd5","branch",
    "Export: Branch books → Ledger → 'Debit to HO' → Alt+E → Excel"));
  grid.appendChild(makeUploadCard("HO LEDGER","Credit to Branch","📗","#2a9d6a","ho",
    "Export: HO books → Ledger → 'Credit to Branch (your branch)' → Alt+E → Excel"));
  grid.appendChild(makeUploadCard("SUSPENSE LEDGER","Suspense Account","📙","#c07a20","susp",
    "Export: HO books → Ledger → Suspense Account → Alt+E → Excel"));
  grid.appendChild(makeUploadCard("BANK ACCOUNT LEDGER","Optional — cross-check only","🏦","#4a90a0","bank",
    "Export: HO or Branch → Bank Account ledger → Alt+E → Excel. Used to find where wrong/unmatched payments were actually booked."));

  container.appendChild(grid);

  // Ledger name config
  const cfgCard = DIV({style:{background:"#0c0d16",border:"1px solid #1a1b28",
    borderRadius:"10px",padding:"16px 20px",marginBottom:"18px"}});
  const cfgTitle = DIV({style:{fontSize:"9px",color:"#3a7bd5",fontFamily:"'IBM Plex Mono',monospace",
    fontWeight:"600",letterSpacing:"2px",marginBottom:"12px"}});
  cfgTitle.textContent = "TALLY LEDGER NAMES (for XML export — must match exactly)";
  cfgCard.appendChild(cfgTitle);

  const cfgRow = DIV({style:{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"12px"}});
  [
    ["branchLedgerName","Branch: Debit to HO ledger name","#3a7bd5"],
    ["hoLedgerName","HO: Credit to Branch ledger name","#2a9d6a"],
    ["suspenseLedgerName","Suspense ledger name","#c07a20"],
  ].forEach(([key, placeholder, color]) => {
    const wrap = DIV({});
    const lbl = DIV({style:{fontSize:"9px",color:color,fontFamily:"'IBM Plex Mono',monospace",
      letterSpacing:"1.5px",marginBottom:"5px",fontWeight:"600"}});
    lbl.textContent = placeholder.toUpperCase();
    const inp = E("input",{style:{width:"100%",fontSize:"11px"},
      value: S[key],
      oninput: ev => { S[key] = ev.target.value; }});
    inp.value = S[key];
    wrap.append(lbl, inp);
    cfgRow.appendChild(wrap);
  });
  cfgCard.appendChild(cfgRow);

  // Bills transfer config row
  const billsRow = DIV({style:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px",marginTop:"12px",
    paddingTop:"12px",borderTop:"1px solid #1a1b28"}});
  const billsHdr = DIV({style:{gridColumn:"1/-1",fontSize:"9px",color:"#8060c0",
    fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",letterSpacing:"2px",marginBottom:"4px"}});
  billsHdr.textContent = "BILLS TRANSFER CONFIG";
  billsRow.appendChild(billsHdr);
  [
    ["billsAuthorisers","Authoriser names (comma-separated)","#8060c0"],
    ["billsKeywords","Transfer keywords (comma-separated)","#8060c0"],
  ].forEach(([key, placeholder, color]) => {
    const wrap = DIV({});
    const lbl = DIV({style:{fontSize:"9px",color:color,fontFamily:"'IBM Plex Mono',monospace",
      letterSpacing:"1.5px",marginBottom:"5px",fontWeight:"600"}});
    lbl.textContent = placeholder.toUpperCase();
    const inp = E("input",{style:{width:"100%",fontSize:"11px"},
      oninput: ev => { S[key] = ev.target.value; }});
    inp.value = S[key];
    wrap.append(lbl, inp);
    billsRow.appendChild(wrap);
  });
  cfgCard.appendChild(billsRow);
  container.appendChild(cfgCard);

  // Run button
  const runRow = DIV({style:{display:"flex",alignItems:"center",gap:"12px",marginBottom:"16px"}});
  const hasData = S.branch.entries.length > 0 && S.ho.entries.length > 0;

  const runBtn = BTN({
    style:{background: hasData ? "linear-gradient(135deg,#3a7bd5,#1a4a8a)" : "#1a1b2e",
      border:"none",color: hasData ? "white" : "#3a3d52",
      padding:"11px 30px",borderRadius:"9px",fontSize:"11px",fontWeight:"600",
      cursor: hasData ? "pointer" : "not-allowed",fontFamily:"'IBM Plex Mono',monospace",
      letterSpacing:"1px",transition:"all 0.2s"},
    onclick: hasData ? runReconciliation : null,
    onmouseenter:(ev)=>{ if(hasData){ev.target.style.transform="translateY(-1px)";ev.target.style.boxShadow="0 6px 20px rgba(58,123,213,0.35)";} },
    onmouseleave:(ev)=>{ ev.target.style.transform="";ev.target.style.boxShadow=""; },
  });
  runBtn.textContent = S.processing ? "⟳  PROCESSING…" : "▶  RUN RECONCILIATION";

  const statusEl = DIV({style:{color:"#5a6080",fontSize:"11px",fontFamily:"'IBM Plex Mono',monospace",flex:"1"}});
  statusEl.textContent = S.statusMsg || (hasData
    ? `Ready: ${S.branch.entries.length} branch · ${S.ho.entries.length} HO · ${S.susp.entries.length} suspense · ${S.bank.entries.length} bank ledger entries`
    : "Upload Branch and HO ledgers to begin (Suspense optional but recommended)");

  runRow.append(runBtn, statusEl);
  container.appendChild(runRow);

  // Entry previews
  if (S.branch.entries.length || S.ho.entries.length || S.susp.entries.length || S.bank.entries.length) {
    const previews = DIV({style:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px"}});
    [
      ["branch","BRANCH ENTRIES","#3a7bd5"],
      ["ho","HO ENTRIES","#2a9d6a"],
      ["susp","SUSPENSE ENTRIES","#c07a20"],
      ["bank","BANK LEDGER ENTRIES","#4a90a0"],
    ].forEach(([key, label, color]) => {
      const entries = S[key].entries;
      if (!entries.length) return;
      const card = DIV({style:{background:"#0a0b14",border:`1px solid ${color}33`,borderRadius:"8px",padding:"12px"}});
      const h = DIV({style:{fontSize:"9px",color,fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",
        letterSpacing:"1.5px",marginBottom:"8px"}});
      h.textContent = `${label} — ${entries.length} entries`;
      card.appendChild(h);
      entries.slice(0,5).forEach(e => {
        const row = DIV({style:{display:"flex",justifyContent:"space-between",padding:"3px 0",
          borderBottom:"1px solid #111220"}});
        const left = SPAN({style:{color:"#7080a0",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",
          flex:"1",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}});
        left.textContent = e.party || e.narration?.slice(0,30) || "—";
        const right = SPAN({style:{color:color,fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",
          marginLeft:"8px",flexShrink:"0"}});
        right.textContent = "₹"+Number(e.amount).toLocaleString("en-IN");
        row.append(left, right);
        card.appendChild(row);
      });
      if (entries.length > 5) {
        const more = DIV({style:{color:"#3a3d52",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",
          marginTop:"5px",textAlign:"center"}});
        more.textContent = `+${entries.length - 5} more`;
        card.appendChild(more);
      }
      previews.appendChild(card);
    });
    container.appendChild(previews);
  }
}

function makeUploadCard(title, subtitle, icon, color, key, hint) {
  const card = DIV({style:{background:"#0c0d16",border:`1px solid ${color}22`,borderRadius:"10px",
    padding:"16px",display:"flex",flexDirection:"column",gap:"10px"}});

  const hdr = DIV({style:{display:"flex",alignItems:"center",gap:"8px"}});
  const ic = SPAN({style:{fontSize:"18px"}}); ic.textContent = icon;
  const titWrap = DIV({});
  const tit = DIV({style:{fontSize:"11px",color,fontFamily:"'IBM Plex Mono',monospace",
    fontWeight:"600",letterSpacing:"1px"}}); tit.textContent = title;
  const sub = DIV({style:{fontSize:"10px",color:"#3a3d52",marginTop:"1px"}}); sub.textContent = subtitle;
  titWrap.append(tit, sub);
  hdr.append(ic, titWrap);
  card.appendChild(hdr);

  const files = S[key].files;
  const zone = DIV({
    style:{border:`2px dashed ${files.length ? color+"55" : "#1c1d2a"}`,borderRadius:"7px",
      padding:"14px 10px",cursor:"pointer",textAlign:"center",
      background: files.length ? color+"08" : "rgba(0,0,0,0.2)",transition:"all 0.2s"},
    onclick: () => pickFiles(key, title),
    onmouseenter:(ev)=>{ev.currentTarget.style.borderColor=color+"77";ev.currentTarget.style.background=color+"0d";},
    onmouseleave:(ev)=>{ev.currentTarget.style.borderColor=files.length?color+"55":"#1c1d2a";ev.currentTarget.style.background=files.length?color+"08":"rgba(0,0,0,0.2)";},
  });

  if (files.length) {
    files.forEach(f => {
      const chip = DIV({style:{color,fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",
        marginBottom:"3px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}});
      chip.textContent = "📄 " + f.name;
      zone.appendChild(chip);
    });
    const cnt = DIV({style:{color:color,fontSize:"9px",letterSpacing:"1px",fontFamily:"'IBM Plex Mono',monospace",
      marginTop:"4px",fontWeight:"600"}});
    cnt.textContent = `${S[key].entries.length} ENTRIES LOADED`;
    zone.appendChild(cnt);
  } else {
    const lbl = DIV({style:{color:"#3a3d52",fontSize:"11px",fontFamily:"'IBM Plex Mono',monospace",
      fontWeight:"600",letterSpacing:"1px"}}); lbl.textContent = "CLICK TO BROWSE";
    const sub2 = DIV({style:{color:"#252635",fontSize:"10px",marginTop:"3px"}}); sub2.textContent = "XLS · XLSX · CSV · XML";
    zone.append(lbl, sub2);
  }
  card.appendChild(zone);

  if (files.length) {
    const clr = BTN({style:{background:"transparent",border:"1px solid #1c1d2a",color:"#3a3d52",
      padding:"4px 10px",borderRadius:"5px",cursor:"pointer",fontSize:"9px",
      fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"1px",alignSelf:"flex-start"},
      onclick:(ev)=>{ ev.stopPropagation(); S[key].files=[]; S[key].entries=[]; S.results=null; render(); }});
    clr.textContent = "CLEAR";
    card.appendChild(clr);
  }

  const hintEl = DIV({style:{color:"#252635",fontSize:"10px",lineHeight:"1.4",marginTop:"2px"}});
  hintEl.textContent = hint;
  card.appendChild(hintEl);

  return card;
}

// ─── RESULTS TAB ──────────────────────────────────────────────────────────────
function renderResultsTab(container) {
  const R = S.results;
  if (!R) return;

  const counts = {
    reconciled:   R.reconciled.length,
    dateMismatch: R.dateMismatch.length,
    fromSuspense: R.fromSuspense.length,
    toSuspense:   R.toSuspense.length,
    wrongLedger:  R.wrongLedger.length,
    manualReview: R.manualReview.length,
  };

  const totalAmt = grp => grp.reduce((s, r) => s + parseFloat(r.branch?.amount || r.ho?.amount || r.suspense?.amount || 0), 0);

  // Stats row
  const stats = DIV({style:{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:"10px",marginBottom:"18px"}});
  [
    { key:"reconciled",   label:"RECONCILED",     color:"#2a9d6a", count:counts.reconciled,   amt:totalAmt(R.reconciled) },
    { key:"dateMismatch", label:"DATE MISMATCH",  color:"#e08030", count:counts.dateMismatch, amt:totalAmt(R.dateMismatch) },
    { key:"fromSuspense", label:"SUSP → BRANCH",  color:"#f59e0b", count:counts.fromSuspense, amt:totalAmt(R.fromSuspense) },
    { key:"toSuspense",   label:"HO → SUSPENSE",  color:"#e07030", count:counts.toSuspense,   amt:totalAmt(R.toSuspense) },
    { key:"wrongLedger",  label:"WRONG LEDGER",   color:"#e05050", count:counts.wrongLedger,  amt:totalAmt(R.wrongLedger) },
    { key:"manualReview", label:"MANUAL REVIEW",  color:"#8060c0", count:counts.manualReview, amt:totalAmt(R.manualReview) },
  ].forEach(s => {
    const active = S.resultsTab === s.key;
    const card = DIV({style:{background: active ? "#0e1020" : "rgba(255,255,255,0.02)",
      border:`1px solid ${active ? s.color+"66" : "#1a1b28"}`,borderRadius:"9px",
      padding:"12px 14px",cursor:"pointer",textAlign:"center",transition:"all 0.15s"},
      onclick:()=>{ S.resultsTab = s.key; render(); },
      onmouseenter:(ev)=>{ if(!active){ev.currentTarget.style.borderColor=s.color+"44";} },
      onmouseleave:(ev)=>{ if(!active){ev.currentTarget.style.borderColor="#1a1b28";} },
    });
    const num = DIV({style:{fontSize:"24px",fontWeight:"600",color:s.color,fontFamily:"'IBM Plex Mono',monospace"}});
    num.textContent = s.count;
    const lbl = DIV({style:{fontSize:"8px",color:"#3a3d52",letterSpacing:"1.5px",fontFamily:"'IBM Plex Mono',monospace",marginTop:"2px"}});
    lbl.textContent = s.label;
    const amtEl = DIV({style:{fontSize:"9px",color:s.color+"99",fontFamily:"'IBM Plex Mono',monospace",marginTop:"3px"}});
    amtEl.textContent = "₹"+Number(s.amt).toLocaleString("en-IN",{maximumFractionDigits:0});
    card.append(num, lbl, amtEl);
    stats.appendChild(card);
  });
  container.appendChild(stats);

  // Export buttons
  const exportRow = DIV({style:{display:"flex",gap:"10px",marginBottom:"16px",alignItems:"center"}});

  const exXls = BTN({style:{background:"#0c1a0c",border:"1px solid #2a9d6a44",color:"#2a9d6a",
    padding:"8px 18px",borderRadius:"7px",cursor:"pointer",fontSize:"10px",
    fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",letterSpacing:"1px"},
    onclick: doExportXlsx });
  exXls.textContent = "⬇  EXPORT EXCEL REPORT";

  const hasTallyExports = R.toSuspense.length > 0 || R.fromSuspense.length > 0;
  const exXml = BTN({style:{background: hasTallyExports ? "#0c0e1c" : "#0a0b12",
    border:`1px solid ${hasTallyExports ? "#3a7bd544" : "#1a1b28"}`,
    color: hasTallyExports ? "#3a7bd5" : "#3a3d52",
    padding:"8px 18px",borderRadius:"7px",
    cursor: hasTallyExports ? "pointer" : "not-allowed",fontSize:"10px",
    fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",letterSpacing:"1px"},
    onclick: hasTallyExports ? doExportTallyXml : null });
  exXml.textContent = "⬇  EXPORT TALLY XML";

  const approvedMismatches = (R.dateMismatch || []).filter(r => r.approved);
  const hasMismatchExport  = approvedMismatches.length > 0;

  const exDateFix = BTN({style:{
    background: hasMismatchExport ? "#1a0e08" : "#0a0b12",
    border: `1px solid ${hasMismatchExport ? "#e0803044" : "#1a1b28"}`,
    color: hasMismatchExport ? "#e08030" : "#3a3d52",
    padding:"8px 18px",borderRadius:"7px",
    cursor: hasMismatchExport ? "pointer" : "not-allowed",
    fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",letterSpacing:"1px"},
    onclick: hasMismatchExport ? doExportDateCorrections : null });
  exDateFix.textContent = `⬇  DATE CORRECTIONS XML (${approvedMismatches.length})`;

  const note = SPAN({style:{color:"#2a2d3e",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",flex:"1"}});
  note.textContent = hasMismatchExport
    ? `${approvedMismatches.length} date corrections approved — export XML to fix journal dates`
    : hasTallyExports
    ? `Tally XML: ${R.toSuspense.length} suspense moves + ${R.fromSuspense.length} suspense adjustments`
    : "No Tally adjustments to export";
  exportRow.append(exXls, exXml, exDateFix, note);
  container.appendChild(exportRow);

  // Search bar
  const searchRow = DIV({style:{display:"flex",gap:"10px",alignItems:"center",marginBottom:"12px"}});
  const searchInp = E("input",{
    style:{flex:"1",padding:"7px 12px",fontSize:"11px",background:"#0c0d16",
      border:"1px solid #1a1b28",borderRadius:"7px",color:"#c0c8e0",
      fontFamily:"'IBM Plex Mono',monospace",outline:"none"},
    placeholder:"Search party, amount, UTR, narration…",
    oninput: ev => { S.searchQuery = ev.target.value; refreshTable(); }
  });
  searchInp.value = S.searchQuery || "";
  const clearSearch = BTN({style:{background:"transparent",border:"1px solid #1a1b28",color:"#3a3d52",
    padding:"6px 12px",borderRadius:"7px",cursor:"pointer",fontSize:"10px",
    fontFamily:"'IBM Plex Mono',monospace"},
    onclick:()=>{ S.searchQuery=""; searchInp.value=""; refreshTable(); }});
  clearSearch.textContent = "✕ CLEAR";
  searchRow.append(searchInp, clearSearch);

  // Approve All button — only shown on dateMismatch tab
  if (tab === "dateMismatch" && (R.dateMismatch||[]).length > 0) {
    const pending = (R.dateMismatch||[]).filter(r => !r.approved);
    const approveAll = BTN({style:{
      background:"linear-gradient(135deg,#2a6a3a,#1a4a2a)",
      border:"1px solid #2a9d6a88",color:"#2a9d6a",
      padding:"7px 18px",borderRadius:"7px",cursor:pending.length?"pointer":"not-allowed",
      fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",fontWeight:"700",letterSpacing:"0.5px",
      opacity: pending.length ? "1" : "0.5",transition:"all 0.15s"},
      onclick: () => {
        if (!pending.length) return;
        (R.dateMismatch||[]).forEach(r => { r.approved = true; });
        render();
      }});
    approveAll.textContent = pending.length
      ? `✓ APPROVE ALL (${pending.length} pending)`
      : "✓ ALL APPROVED";
    searchRow.append(approveAll);

    // Info note
    const infoNote = DIV({style:{padding:"8px 14px",background:"#100e06",border:"1px solid #3a3010",
      borderRadius:"7px",marginBottom:"12px",fontSize:"11px",color:"#7a6a30",fontFamily:"'IBM Plex Mono',monospace"}});
    infoNote.textContent = "Review each entry below. Journal date (orange) will change to receipt date (green) in Tally. Approve individually or use Approve All, then export XML.";
    container.appendChild(infoNote);
  }

  container.appendChild(searchRow);

  // Results table — rendered into a slot so search can swap body without full re-render
  const tab = S.resultsTab;
  const tableSlot = DIV({});
  container.appendChild(tableSlot);

  function refreshTable() {
    tableSlot.innerHTML = "";
    tableSlot.appendChild(renderTable(tab, R));
  }
  refreshTable();
}

function filterRows(rows, query) {
  if (!query || !query.trim()) return rows;
  const q = query.toLowerCase();
  return rows.filter(r => {
    const fields = [
      r.branch?.party, r.branch?.narration, r.branch?.utr, String(r.branch?.amount||""),
      r.ho?.party, r.ho?.narration, r.ho?.utr, String(r.ho?.amount||""),
      r.suspense?.party, r.suspense?.narration, r.suspense?.utr,
      r.reason, r.method,
    ];
    return fields.some(f => f && String(f).toLowerCase().includes(q));
  });
}

function renderTable(tab, R) {
  const COLS_CFG = {
    reconciled:   { cols:"28px 90px 1fr 80px 1fr 80px 110px 80px 32px", heads:["#","TYPE","BRANCH","AMOUNT","HO ENTRY","AMOUNT","METHOD [SCORE]","NET",""] },
    dateMismatch: { cols:"28px 1fr 80px 110px 110px 110px 100px 32px",  heads:["#","ENTRY","AMOUNT","JOURNAL MONTH","RECEIPT MONTH","APPROVED?","ACTION",""] },
    fromSuspense: { cols:"28px 1fr 80px 1fr 80px 120px 32px",           heads:["#","BRANCH ENTRY","AMOUNT","SUSPENSE","AMOUNT","ACTION",""] },
    toSuspense:   { cols:"28px 70px 1fr 80px 120px 32px",               heads:["#","TYPE","HO ENTRY","AMOUNT","ACTION",""] },
    wrongLedger:  { cols:"28px 70px 1fr 80px 1fr 32px",                 heads:["#","TYPE","BRANCH ENTRY","AMOUNT","REASON / SUGGESTION",""] },
    manualReview: { cols:"28px 80px 1fr 80px 1fr 1fr 30px 32px",        heads:["#","SOURCE","ENTRY","AMOUNT","CANDIDATE","REASON","✓",""] },
  };
  const STATUS_COLOR = {
    reconciled:"#2a9d6a", fromSuspense:"#f59e0b",
    toSuspense:"#e07030", wrongLedger:"#e05050", manualReview:"#8060c0"
  };

  const cfg = COLS_CFG[tab];
  const color = STATUS_COLOR[tab];
  const wrap = DIV({style:{background:"#09091280",border:"1px solid #13141e",borderRadius:"10px",overflow:"hidden"}});

  // Sticky header
  const thead = DIV({style:{display:"grid",gridTemplateColumns:cfg.cols,gap:"8px",
    padding:"9px 12px",background:"#0a0b14",borderBottom:"1px solid #13141e",position:"sticky",top:"0",zIndex:"2"}});
  cfg.heads.forEach(h => {
    const c = DIV({style:{fontSize:"8px",color:"#2a2d3e",letterSpacing:"1.5px",
      fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600"}}); c.textContent = h;
    thead.appendChild(c);
  });
  wrap.appendChild(thead);

  const tbody = DIV({style:{maxHeight:"440px",overflowY:"auto"}});
  const allRows = R[tab] || [];
  const rows = filterRows(allRows, S.searchQuery);

  if (!rows.length) {
    const empty = DIV({style:{padding:"40px",textAlign:"center",color:"#2a2d3e",fontFamily:"'IBM Plex Mono',monospace",fontSize:"12px"}});
    empty.textContent = S.searchQuery ? "No entries match your search" : "No entries in this category";
    tbody.appendChild(empty);
  }

  // Type badge helper
  const typeBadge = (isBills) => {
    const b = DIV({style:{display:"inline-block",padding:"2px 6px",borderRadius:"8px",fontSize:"8px",
      fontFamily:"'IBM Plex Mono',monospace",fontWeight:"700",
      background:isBills?"#1a0f2a":"#0d1a14",
      border:`1px solid ${isBills?"#8060c044":"#2a9d6a33"}`,
      color:isBills?"#8060c0":"#2a9d6a"}});
    b.textContent = isBills ? "BILLS" : "BANK";
    return b;
  };

  rows.forEach((r, i) => {
    const rowWrap = DIV({style:{borderBottom:"1px solid #0c0d16"}});
    const rowGrid = DIV({style:{display:"grid",gridTemplateColumns:cfg.cols,gap:"8px",
      padding:"9px 12px",alignItems:"center",cursor:"pointer",
      background:i%2===0?"rgba(255,255,255,0.01)":"transparent",
      transition:"background 0.1s",
      borderLeft:r.isBillsTransfer?"3px solid #8060c044":"3px solid transparent"}});
    rowWrap.addEventListener("mouseenter",()=>rowGrid.style.background="rgba(255,255,255,0.03)");
    rowWrap.addEventListener("mouseleave",()=>rowGrid.style.background=i%2===0?"rgba(255,255,255,0.01)":"transparent");

    const num = DIV({style:{color:"#252635",fontSize:"10px",fontFamily:"monospace"}});
    num.textContent = i+1;
    const toggle = DIV({style:{color:"#252635",fontSize:"11px",textAlign:"center",padding:"2px 4px"}});
    toggle.textContent = "▼";

    // Right-click context menu
    const mainEntry = r.branch||r.ho||r.suspense;
    const matchEntry = r.ho||r.suspense;
    rowGrid.addEventListener("contextmenu", ev => showContextMenu(ev, mainEntry, matchEntry!==mainEntry?matchEntry:null));

    if (tab === "reconciled") {
      const confBadge = DIV({style:{display:"inline-block",padding:"2px 6px",borderRadius:"8px",fontSize:"8px",
        fontFamily:"'IBM Plex Mono',monospace",fontWeight:"700",
        background:r.autoConfirmed?"#0d2b1e":"#1a1808",
        border:`1px solid ${r.autoConfirmed?"#2a9d6a55":"#7a6a1a55"}`,
        color:r.autoConfirmed?"#2a9d6a":"#f59e0b"}});
      confBadge.textContent = r.autoConfirmed ? "AUTO ✓" : "CONFIRM";
      const typeCell = DIV({style:{display:"flex",flexDirection:"column",gap:"3px"}});
      typeCell.append(typeBadge(r.isBillsTransfer), confBadge);
      rowGrid.append(num, typeCell, entryCell(r.branch,color), amtCell(r.branch?.amount,color),
        entryCell(r.ho,"#5080b0"), amtCell(r.ho?.amount,"#5080b0"),
        methodCell(r.method,r.score), netCell(r.net), toggle);

    } else if (tab === "dateMismatch") {
      // Month mismatch row — shows journal month vs receipt month, approve button
      const isApproved = r.approved;
      const jmonth = DIV({style:{fontFamily:"'IBM Plex Mono',monospace",fontSize:"11px",
        color:"#e08030",fontWeight:"600"}});
      jmonth.textContent = r.journalFYMonth || r.journalDate || "—";
      const rmonth = DIV({style:{fontFamily:"'IBM Plex Mono',monospace",fontSize:"11px",
        color:"#2a9d6a",fontWeight:"600"}});
      rmonth.textContent = r.receiptFYMonth || r.receiptDate || "—";

      const statusEl = DIV({style:{
        display:"inline-flex",alignItems:"center",gap:"5px",
        padding:"3px 10px",borderRadius:"12px",fontSize:"10px",
        fontFamily:"'IBM Plex Mono',monospace",fontWeight:"700",
        background: isApproved ? "#0d2b1e" : "#1a1a0a",
        border: `1px solid ${isApproved ? "#2a9d6a55" : "#5a5a2a55"}`,
        color: isApproved ? "#2a9d6a" : "#a0a020",
      }});
      statusEl.textContent = isApproved ? "✓ APPROVED" : "PENDING";

      const approveBtn = BTN({style:{
        background: isApproved ? "#0a1a0a" : "linear-gradient(135deg,#2a6a3a,#1a4a2a)",
        border:`1px solid ${isApproved?"#2a9d6a33":"#2a9d6a88"}`,
        color: isApproved ? "#2a9d6a66" : "#2a9d6a",
        padding:"5px 14px",borderRadius:"6px",cursor: isApproved?"default":"pointer",
        fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",
        letterSpacing:"0.5px",transition:"all 0.15s",
      }});
      approveBtn.textContent = isApproved ? "✓ Approved" : "APPROVE";
      if (!isApproved) {
        approveBtn.addEventListener("mouseenter",()=>{approveBtn.style.background="linear-gradient(135deg,#3a8a4a,#2a6a3a)";});
        approveBtn.addEventListener("mouseleave",()=>{approveBtn.style.background="linear-gradient(135deg,#2a6a3a,#1a4a2a)";});
        approveBtn.addEventListener("click", ev => {
          ev.stopPropagation();
          r.approved = true;
          render(); // re-render to update export button state
        });
      }

      rowGrid.style.background = isApproved
        ? (i%2===0?"rgba(42,157,106,0.05)":"rgba(42,157,106,0.03)")
        : (i%2===0?"rgba(255,255,255,0.01)":"transparent");

      rowGrid.append(num, entryCell(r.branch, "#e08030"), amtCell(r.branch?.amount, "#e08030"),
        jmonth, rmonth, statusEl, approveBtn, toggle);
      rowGrid.append(num, entryCell(r.branch,color), amtCell(r.branch?.amount,color),
        entryCell(r.suspense,"#7060a0"), amtCell(r.suspense?.amount,"#7060a0"),
        actionCell(r.action,color), toggle);

    } else if (tab === "toSuspense") {
      rowGrid.append(num, typeBadge(r.isBillsTransfer), entryCell(r.ho,color),
        amtCell(r.ho?.amount,color), actionCell(r.action,color), toggle);

    } else if (tab === "wrongLedger") {
      const bankDot = r.bankHint ? SPAN({style:{display:"inline-block",width:"7px",height:"7px",
        borderRadius:"50%",background:"#4a90a0",marginLeft:"4px",flexShrink:"0",title:"Payment found in bank ledger"}}) : null;
      const reasonWrap = DIV({style:{display:"flex",alignItems:"center",gap:"6px",overflow:"hidden"}});
      reasonWrap.appendChild(reasonCell(r.reason, r.bankHint
        ? `🏦 Found in bank ledger → booked to: ${r.bankHint.party || "unknown"}`
        : r.suggestion));
      if (bankDot) reasonWrap.insertBefore(bankDot, reasonWrap.firstChild);
      rowGrid.append(num, typeBadge(r.isBillsTransfer), entryCell(r.branch,color),
        amtCell(r.branch?.amount,color), reasonWrap, toggle);

    } else if (tab === "manualReview") {
      const src = r.source==="Suspense-Unmatched"?r.suspense:(r.branch||r.ho);
      const mat = r.ho||r.suspense;
      const isConflict = r.reason?.startsWith("Conflicted");
      const srcLbl = DIV({style:{color:isConflict?"#e05050":r.isBillsTransfer?"#8060c0":color,
        fontSize:"9px",fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600"}});
      srcLbl.textContent = isConflict?"⚡CONFLICT":r.isBillsTransfer?"BILLS":
        (r.source==="Suspense-Unmatched"?"SUSP":r.source);

      // Mark-reviewed checkbox
      const globalIdx = (R.manualReview||[]).indexOf(r);
      const isReviewed = S.reviewedSet.has(globalIdx);
      const rvBox = DIV({style:{width:"16px",height:"16px",borderRadius:"3px",cursor:"pointer",flexShrink:"0",
        border:`1px solid ${isReviewed?"#2a9d6a":"#2a2d3e"}`,
        background:isReviewed?"#0d2b1e":"transparent",
        display:"flex",alignItems:"center",justifyContent:"center",fontSize:"10px",color:"#2a9d6a"}});
      rvBox.textContent = isReviewed ? "✓" : "";
      rvBox.addEventListener("click", ev => {
        ev.stopPropagation();
        if(S.reviewedSet.has(globalIdx)) S.reviewedSet.delete(globalIdx);
        else S.reviewedSet.add(globalIdx);
        const rev = S.reviewedSet.has(globalIdx);
        rvBox.style.background = rev?"#0d2b1e":"transparent";
        rvBox.style.borderColor = rev?"#2a9d6a":"#2a2d3e";
        rvBox.textContent = rev?"✓":"";
        rowGrid.style.opacity = rev?"0.45":"1";
      });
      if (isReviewed) rowGrid.style.opacity = "0.45";

      rowGrid.append(num, srcLbl, entryCell(src,isConflict?"#e05050":color),
        amtCell(src?.amount,isConflict?"#e05050":color),
        entryCell(mat,"#5060a0"), reasonCell(r.reason), rvBox, toggle);
    }

    rowWrap.appendChild(rowGrid);

    // Expandable detail — text selectable, values clickable to copy
    let expanded = false;
    const detail = DIV({style:{display:"none",padding:"12px 18px 14px",
      background:"#060710",borderTop:"1px solid #10111c",userSelect:"text",WebkitUserSelect:"text"}});

    const buildDetail = () => {
      detail.innerHTML = "";
      const grid2 = DIV({style:{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"14px"}});
      const SHOW_KEYS = ["date","party","amount","utr","narration","ref","sourceFile","debitCredit"];
      const addSide = (entry, lbl, clr) => {
        if (!entry) return;
        const side = DIV({});
        const h = DIV({style:{color:clr,fontSize:"9px",letterSpacing:"2px",fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",marginBottom:"8px"}});
        h.textContent = lbl; side.appendChild(h);
        Object.entries(entry).filter(([k])=>SHOW_KEYS.includes(k)).forEach(([k,v])=>{
          const row = DIV({style:{display:"flex",gap:"8px",marginBottom:"3px"}});
          const kk = SPAN({style:{color:"#2a2d3e",fontSize:"10px",minWidth:"80px",fontFamily:"'IBM Plex Mono',monospace",flexShrink:"0"}});
          kk.textContent = k+":";
          const vv = SPAN({style:{color:"#8090b0",fontSize:"10px",wordBreak:"break-all",cursor:"pointer",userSelect:"text"}});
          vv.textContent = String(v);
          vv.title = "Click to copy";
          vv.addEventListener("click",ev=>{ev.stopPropagation();navigator.clipboard.writeText(String(v)).catch(()=>{});vv.style.color="#2a9d6a";setTimeout(()=>vv.style.color="#8090b0",700);});
          row.append(kk,vv); side.appendChild(row);
        });
        grid2.appendChild(side);
      };
      if(tab==="reconciled")  {addSide(r.branch,"BRANCH","#3a7bd5");addSide(r.ho,"HO","#2a9d6a");}
      if(tab==="dateMismatch"){addSide(r.branch,"BRANCH (journal to change)","#e08030");addSide(r.ho,"HO RECEIPT (source of truth)","#2a9d6a");}
      if(tab==="fromSuspense"){addSide(r.branch,"BRANCH","#3a7bd5");addSide(r.suspense,"SUSPENSE","#c07a20");}
      if(tab==="toSuspense")  {addSide(r.ho,"HO ENTRY","#2a9d6a");}
      if(tab==="wrongLedger") {addSide(r.branch,"BRANCH","#e05050");}
      if(tab==="manualReview"){
        const src=r.source==="Suspense-Unmatched"?r.suspense:(r.branch||r.ho);
        addSide(src,"ENTRY","#8060c0"); addSide(r.ho||r.suspense,"BEST CANDIDATE","#5060a0");
      }
      detail.appendChild(grid2);

      // Date mismatch: show the change preview
      if (tab==="dateMismatch") {
        const preview = DIV({style:{marginTop:"10px",padding:"10px 14px",background:"#0e0e06",
          borderRadius:"6px",border:"1px solid #3a3a10",display:"flex",gap:"24px",alignItems:"center",flexWrap:"wrap"}});
        const arrow = (label, fromVal, toVal, color) => {
          const g = DIV({style:{display:"flex",flexDirection:"column",gap:"3px"}});
          const lbl = DIV({style:{color:"#3a3a20",fontSize:"9px",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"1.5px"}});
          lbl.textContent = label;
          const vals = DIV({style:{display:"flex",alignItems:"center",gap:"8px"}});
          const fv = SPAN({style:{color:"#e08030",fontSize:"12px",fontFamily:"'IBM Plex Mono',monospace",fontWeight:"700",textDecoration:"line-through"}});
          fv.textContent = fromVal;
          const arr = SPAN({style:{color:"#3a3a20",fontSize:"14px"}}); arr.textContent = "→";
          const tv = SPAN({style:{color:"#2a9d6a",fontSize:"12px",fontFamily:"'IBM Plex Mono',monospace",fontWeight:"700"}});
          tv.textContent = toVal;
          vals.append(fv, arr, tv);
          g.append(lbl, vals);
          return g;
        };
        preview.appendChild(arrow("JOURNAL DATE CHANGE", r.journalDate||"—", r.receiptDate||"—"));
        const vref = DIV({style:{display:"flex",flexDirection:"column",gap:"3px"}});
        const vlbl = DIV({style:{color:"#3a3a20",fontSize:"9px",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"1.5px"}});
        vlbl.textContent = "VOUCHER REF";
        const vval = SPAN({style:{color:"#7080a0",fontSize:"11px",fontFamily:"'IBM Plex Mono',monospace",cursor:"pointer"}});
        vval.textContent = r.tallyAlterEntry?.voucherRef || "(no voucher number — set manually in Tally)";
        vval.addEventListener("click",()=>{navigator.clipboard.writeText(r.tallyAlterEntry?.voucherRef||"").catch(()=>{});});
        vref.append(vlbl, vval);
        preview.appendChild(vref);
        detail.appendChild(preview);
      }

      // Bank ledger cross-check hint
      if (r.bankHint) {
        const hint = DIV({style:{marginTop:"10px",padding:"8px 12px",background:"#061418",
          borderRadius:"6px",borderLeft:"3px solid #4a90a0",display:"flex",gap:"14px",flexWrap:"wrap"}});
        const htitle = DIV({style:{color:"#4a90a0",fontSize:"9px",fontFamily:"'IBM Plex Mono',monospace",
          fontWeight:"700",letterSpacing:"2px",marginBottom:"4px",width:"100%"}});
        htitle.textContent = "🏦 BANK LEDGER — PAYMENT FOUND HERE";
        const fields = [
          ["Booked to", r.bankHint.party],
          ["Date",      r.bankHint.date],
          ["Amount",    "₹"+Number(r.bankHint.amount).toLocaleString("en-IN")],
          ["Narration", r.bankHint.narration],
        ];
        fields.forEach(([k,v]) => {
          const f = DIV({style:{display:"flex",gap:"6px",alignItems:"baseline"}});
          const fk = SPAN({style:{color:"#2a3d42",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",minWidth:"70px"}});
          fk.textContent = k+":";
          const fv = SPAN({style:{color:"#7ac0cc",fontSize:"11px",fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",cursor:"pointer"}});
          fv.textContent = v||"—";
          fv.title = "Click to copy";
          fv.addEventListener("click",ev=>{ev.stopPropagation();navigator.clipboard.writeText(String(v||"")).catch(()=>{});fv.style.color="#2a9d6a";setTimeout(()=>fv.style.color="#7ac0cc",700);});
          f.append(fk,fv); hint.appendChild(f);
        });
        hint.insertBefore(htitle, hint.firstChild);
        detail.appendChild(hint);
      }

      if (r.score > 0 && r.method) {
        const sc = r.score>=90?"#2a9d6a":r.score>=80?"#4a9d5a":r.score>=70?"#f59e0b":"#c07a20";
        const conf = DIV({style:{marginTop:"10px",padding:"8px 12px",background:"#0c0d18",
          borderRadius:"6px",borderLeft:`3px solid ${sc}`,display:"flex",alignItems:"center",gap:"16px"}});
        conf.innerHTML = `<span style="color:#2a2d3e;font-size:10px;font-family:'IBM Plex Mono',monospace">Score: </span>`
          +`<span style="color:${sc};font-size:12px;font-weight:700;font-family:'IBM Plex Mono',monospace">${Math.round(r.score)}</span>`
          +`<span style="color:#4a5070;font-size:10px;font-family:'IBM Plex Mono',monospace;flex:1"> Signals: ${r.method||"—"}</span>`;
        detail.appendChild(conf);
      }

      if (r.alternatives?.length > 0) {
        const altHdr = DIV({style:{marginTop:"10px",fontSize:"9px",color:"#e05050",
          fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",letterSpacing:"1.5px",marginBottom:"6px"}});
        altHdr.textContent = `⚡ ${r.alternatives.length} CONFLICTING CANDIDATE(S)`;
        detail.appendChild(altHdr);
        r.alternatives.forEach(alt=>{
          const ar = DIV({style:{padding:"6px 10px",background:"#0e0a12",borderRadius:"5px",
            border:"1px solid #2a1a3a",marginBottom:"4px",display:"flex",gap:"12px",alignItems:"center"}});
          const as_ = DIV({style:{color:"#8060c0",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",minWidth:"40px"}});
          as_.textContent=`[${Math.round(alt.score)}]`;
          const an = DIV({style:{color:"#6070a0",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",flex:"1"}});
          an.textContent=(alt.entry?.party||alt.entry?.narration?.slice(0,35)||"—")+" · "+(alt.entry?.date||"")+" · ₹"+Number(alt.entry?.amount||0).toLocaleString("en-IN");
          const am = DIV({style:{color:"#3a3d52",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace"}});
          am.textContent=alt.method||"";
          ar.append(as_,an,am); detail.appendChild(ar);
        });
      }
    };

    rowGrid.addEventListener("click", ev => {
      if (ev.button===2) return;
      expanded = !expanded;
      if (expanded && !detail.children.length) buildDetail();
      detail.style.display = expanded ? "block" : "none";
      toggle.textContent = expanded ? "▲" : "▼";
    });
    rowWrap.appendChild(detail);
    tbody.appendChild(rowWrap);
  });

  wrap.appendChild(tbody);

  // Count + total amount footer
  const totalAmt = allRows.reduce((s,r)=>s+parseFloat(r.branch?.amount||r.ho?.amount||r.suspense?.amount||0),0);
  const footer = DIV({style:{padding:"7px 14px",background:"#0a0a12",borderTop:"1px solid #13141e",
    display:"flex",justifyContent:"space-between",alignItems:"center"}});
  const fl = DIV({style:{color:"#2a2d3e",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace"}});
  fl.textContent = `${rows.length}${rows.length<allRows.length?" (filtered)":""} of ${allRows.length} entries`;
  const fr = DIV({style:{color:"#3a4060",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace"}});
  fr.textContent = "Total: ₹"+totalAmt.toLocaleString("en-IN",{maximumFractionDigits:0});
  footer.append(fl,fr);
  wrap.appendChild(footer);

  return wrap;
}

// Cell helpers
function entryCell(e, color) {
  const cell = DIV({style:{overflow:"hidden"}});
  if (!e) { const na=DIV({style:{color:"#252635",fontSize:"11px"}}); na.textContent="—"; cell.appendChild(na); return cell; }
  const name = DIV({style:{color:color||"#8090b0",fontSize:"11px",fontFamily:"'IBM Plex Mono',monospace",
    whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}});
  name.textContent = e.party || e.narration?.slice(0,40) || "—";
  const meta = DIV({style:{color:"#2a2d3e",fontSize:"10px",marginTop:"2px",
    whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}});
  meta.textContent = [e.date, e.utr||"no UTR"].filter(Boolean).join(" · ");
  cell.append(name, meta);
  return cell;
}
function amtCell(amt, color) {
  const d = DIV({style:{color:color||"#8090b0",fontSize:"11px",fontFamily:"'IBM Plex Mono',monospace",
    fontWeight:"600",whiteSpace:"nowrap"}});
  d.textContent = amt ? "₹"+Number(amt).toLocaleString("en-IN",{maximumFractionDigits:2}) : "—";
  return d;
}
function methodCell(method, score, badgeEl) {
  const wrap = DIV({});
  if (badgeEl) wrap.appendChild(badgeEl);
  const scoreColor = score>=90?"#2a9d6a": score>=80?"#4a9d5a": score>=70?"#f59e0b":"#c07a20";
  const d = DIV({style:{fontSize:"9px",fontFamily:"'IBM Plex Mono',monospace",
    color:scoreColor,letterSpacing:"0.3px",marginTop: badgeEl ? "4px" : "0"}});
  // Show score alongside method
  d.textContent = (method || "—") + (score ? ` [${score}]` : "");
  wrap.appendChild(d);
  return wrap;
}
function netCell(net) {
  const n = parseFloat(net||0);
  const d = DIV({style:{fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",fontWeight:"600",
    color: n===0 ? "#2a9d6a" : "#e07030"}});
  d.textContent = n === 0 ? "✓ 0.00" : "Δ "+n.toFixed(2);
  return d;
}
function actionCell(action, color) {
  const d = DIV({style:{color:color,fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace",
    fontWeight:"600",lineHeight:"1.4"}});
  d.textContent = action || "—";
  return d;
}
function reasonCell(reason, suggestion) {
  const wrap = DIV({});
  const r = DIV({style:{color:"#e05050",fontSize:"10px",fontFamily:"'IBM Plex Mono',monospace"}});
  r.textContent = reason || "—";
  wrap.appendChild(r);
  if (suggestion) {
    const s = DIV({style:{color:"#5060a0",fontSize:"10px",marginTop:"2px"}});
    s.textContent = suggestion;
    wrap.appendChild(s);
  }
  return wrap;
}

// ─── GUIDE TAB ────────────────────────────────────────────────────────────────
function renderGuideTab(container) {
  const sections = [
    { title:"How This Tool Works", color:"#3a7bd5", items:[
      "Upload 3 ledger exports: ① Branch 'Debit to HO' ② HO 'Credit to Branch' ③ Suspense Account.",
      "Step 1 — Each Branch debit is matched against HO credits. If matched (net = 0) → RECONCILED.",
      "Step 2 — Branch debits not in HO → checked against Suspense. Found → flagged 'Suspense → Branch' for rebook.",
      "Step 3 — Branch debits not in HO and not in Suspense → flagged 'Wrong Ledger/Branch' for investigation.",
      "Step 4 — HO credits with no matching Branch debit → flagged 'HO → Suspense' to be moved to suspense.",
      "Export: Excel report (all 5 categories) + Tally XML (journal entries ready to import into Tally).",
    ]},
    { title:"Tally Export Steps", color:"#2a9d6a", items:[
      "Gateway of Tally → Display More Reports → Account Books → Ledger.",
      "Select the ledger (e.g. 'Debit to HO') → set date range → press Alt+E → Export.",
      "Choose Excel format (.xlsx). Column names like Date, Particulars, Debit, Credit will be auto-detected.",
      "Repeat for HO ledger 'Credit to [Branch Name]' and Suspense Account.",
      "No formatting changes needed — export as-is from Tally.",
    ]},
    { title:"Matching Logic (v3 — Trust-First)", color:"#f59e0b", items:[
      "PRIMARY: Amount + Date. Same day + same amount = strong signal (30+35 pts). Within ±2 days = good signal.",
      "SECONDARY: Full UTR exact match = 60 base pts, boosted to 90-100 with amount + date corroboration.",
      "Partial UTR (Last-4, Last-5, chunk): NEVER used alone. Only adds points when amount already matches.",
      "Shared numeric ref in narration (6+ digits matching both sides) = strong fallback when UTR differs payer vs receiver.",
      "Amount uniqueness: if same amount appears 3+ times across all ledgers, treated as common — never auto-matched on amount alone.",
      "Conflict detection: if two candidates score within 15 pts of each other, both are flagged as CONFLICT → manual review.",
      "Tiers: Score ≥95 = AUTO ✓ (no check needed) · 80-94 = Reconciled (spot-check) · 70-79 = Probable · <70 = Manual Review.",
    ]},
    { title:"Tally XML Import", color:"#8060c0", items:[
      "The exported XML contains Journal vouchers — one per adjustment needed.",
      "In Tally: Gateway → Import → XML Data → select the exported file.",
      "Before importing, verify ledger names in the config section match exactly with your Tally ledger names.",
      "The XML uses the ledger names you set in the 'Tally Ledger Names' fields on the Upload tab.",
      "After import, vouchers will appear as Journal entries for your review before finalising.",
    ]},
  ];

  sections.forEach(sec => {
    const card = DIV({style:{background:"#0a0b14",border:"1px solid #13141e",borderRadius:"10px",
      overflow:"hidden",marginBottom:"14px"}});
    const hdr = DIV({style:{padding:"11px 16px",borderBottom:"1px solid #13141e",
      borderLeft:`3px solid ${sec.color}`}});
    const ht = SPAN({style:{fontFamily:"'IBM Plex Mono',monospace",fontSize:"11px",
      fontWeight:"600",color:sec.color,letterSpacing:"0.5px"}});
    ht.textContent = sec.title;
    hdr.appendChild(ht);
    card.appendChild(hdr);

    sec.items.forEach((item, ii) => {
      const row = DIV({style:{padding:"8px 16px 8px 20px",
        borderBottom: ii<sec.items.length-1 ? "1px solid #0c0d16" : "none",
        display:"flex",gap:"10px",alignItems:"flex-start"}});
      const arr = SPAN({style:{color:sec.color+"55",fontSize:"12px",flexShrink:"0",marginTop:"1px"}});
      arr.textContent = "›";
      const txt = SPAN({style:{color:"#6070a0",fontSize:"12px",lineHeight:"1.6"}});
      txt.textContent = item;
      row.append(arr, txt);
      card.appendChild(row);
    });
    container.appendChild(card);
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// ACTIONS
// ═══════════════════════════════════════════════════════════════════════════════
async function pickFiles(key, label) {
  const paths = await window.api.openFiles(label);
  if (!paths.length) return;

  S.statusMsg = `Parsing ${paths.length} file(s)…`;
  render();

  let allEntries = [];
  for (const p of paths) {
    S.statusMsg = `Parsing: ${p.split(/[\\\/]/).pop()}…`;
    const r = document.getElementById("root");
    // update status inline without full re-render
    const entries = await window.api.parseFile(p);
    if (entries.error) { alert(`Error parsing file:\n${entries.error}`); continue; }
    allEntries.push(...entries);
  }

  S[key].files = paths.map(p => ({ path: p, name: p.split(/[\\\/]/).pop() }));
  S[key].entries = allEntries;
  S.results = null;
  S.statusMsg = `${label}: ${allEntries.length} entries loaded`;
  render();
}

async function runReconciliation() {
  if (S.processing) return;
  if (!S.branch.entries.length || !S.ho.entries.length) {
    S.statusMsg = "⚠ Branch and HO ledgers are required.";
    render(); return;
  }

  // Sync bills config into engine globals before running
  BILLS_AUTHORISER_NAMES = S.billsAuthorisers
    .split(",").map(s => s.trim().toLowerCase()).filter(Boolean);
  BILLS_KEYWORDS = S.billsKeywords
    .split(",").map(s => s.trim().toLowerCase()).filter(Boolean);

  S.processing = true;
  S.statusMsg = "Running reconciliation…";
  S.reviewedSet = new Set();
  render();

  await new Promise(r => setTimeout(r, 50));

  try {
    S.results = reconcile(S.branch.entries, S.ho.entries, S.susp.entries);
    // Split date-mismatched entries into their own category
    S.results = splitDateMismatches(S.results);
    // Cross-check wrong/unmatched entries against bank ledger if provided
    if (S.bank.entries.length) {
      S.results = crossCheckBankLedger(S.results, S.bank.entries);
    }
    const res = S.results;
    const billsCount = res.reconciled.filter(r => r.isBillsTransfer).length;
    S.statusMsg = `Done — ${res.reconciled.length} reconciled (${billsCount} bills) · `
      + `${res.dateMismatch.length} date mismatch · `
      + `${res.fromSuspense.length} susp→branch · ${res.toSuspense.length} HO→susp · `
      + `${res.wrongLedger.length} wrong ledger · ${res.manualReview.length} manual review`;
    S.tab = "results";
    S.resultsTab = "reconciled";
    S.searchQuery = "";
  } catch(e) {
    S.statusMsg = "Error: " + e.message;
    console.error(e);
  } finally {
    S.processing = false;
    render();
  }
}

// ── Context menu for copy ─────────────────────────────────────────────────────
let ctxMenu = null;
function showContextMenu(e, entry, matchEntry) {
  e.preventDefault();
  e.stopPropagation();
  removeContextMenu();

  const items = [];
  if (entry) {
    items.push(
      { label: `Copy party: ${entry.party||"—"}`,   val: entry.party||"" },
      { label: `Copy amount: ₹${Number(entry.amount||0).toLocaleString("en-IN")}`, val: String(entry.amount||"") },
      { label: `Copy UTR: ${entry.utr||"—"}`,       val: entry.utr||"" },
      { label: `Copy date: ${entry.date||"—"}`,     val: entry.date||"" },
      { label: `Copy narration`,                    val: entry.narration||"" },
    );
  }
  if (matchEntry) {
    items.push({ label: "── matched ──", val: null });
    items.push(
      { label: `Copy matched party: ${matchEntry.party||"—"}`, val: matchEntry.party||"" },
      { label: `Copy matched amount: ₹${Number(matchEntry.amount||0).toLocaleString("en-IN")}`, val: String(matchEntry.amount||"") },
      { label: `Copy matched UTR: ${matchEntry.utr||"—"}`, val: matchEntry.utr||"" },
    );
  }
  items.push({ label: "── row ──", val: null });
  const rowText = [entry, matchEntry].filter(Boolean)
    .map(en => [en.date, en.party, en.amount, en.utr, en.narration].filter(Boolean).join("\t"))
    .join("\n");
  items.push({ label: "Copy full row as text", val: rowText });

  ctxMenu = DIV({style:{
    position:"fixed", top: Math.min(e.clientY, window.innerHeight-200)+"px",
    left: Math.min(e.clientX, window.innerWidth-240)+"px",
    background:"#0e0f1a", border:"1px solid #2a2d3e", borderRadius:"8px",
    padding:"4px", zIndex:"9999", minWidth:"220px",
    boxShadow:"0 8px 32px rgba(0,0,0,0.6)", userSelect:"none",
  }});

  items.forEach(item => {
    if (!item.val && item.val !== "") {
      const sep = DIV({style:{height:"1px",background:"#1a1d2a",margin:"3px 4px"}});
      ctxMenu.appendChild(sep); return;
    }
    const row = DIV({style:{
      padding:"6px 12px", borderRadius:"5px", cursor:"pointer",
      color: item.val ? "#8090b0" : "#3a3d52",
      fontSize:"11px", fontFamily:"'IBM Plex Mono',monospace",
      transition:"background 0.1s", whiteSpace:"nowrap", overflow:"hidden",
      textOverflow:"ellipsis", maxWidth:"220px",
    }});
    row.textContent = item.label;
    if (item.val !== null) {
      row.addEventListener("mouseenter", () => row.style.background = "#1a1d2e");
      row.addEventListener("mouseleave", () => row.style.background = "transparent");
      row.addEventListener("click", () => {
        navigator.clipboard.writeText(item.val).catch(() => {});
        removeContextMenu();
      });
    }
    ctxMenu.appendChild(row);
  });

  document.body.appendChild(ctxMenu);
  setTimeout(() => document.addEventListener("click", removeContextMenu, { once: true }), 0);
}

function removeContextMenu() {
  if (ctxMenu) { ctxMenu.remove(); ctxMenu = null; }
  document.removeEventListener("click", removeContextMenu);
}

async function doExportDateCorrections() {
  const R = S.results;
  const approved = (R.dateMismatch || []).filter(r => r.approved);
  if (!approved.length) return;

  const entries = approved.map(r => ({
    voucherRef:  r.tallyAlterEntry.voucherRef,
    partyLedger: r.tallyAlterEntry.partyLedger,
    newDate:     r.tallyAlterEntry.newDate,
    oldDate:     r.journalDate,
    amount:      r.tallyAlterEntry.amount,
    narration:   r.tallyAlterEntry.narration,
    fromLedger:  S.branchLedgerName,
    toLedger:    S.hoLedgerName,
  }));

  await window.api.exportDateCorrections(entries);
}

async function doExportXlsx() {
  const R = S.results;
  const fmtAmt = a => a ? Number(a).toLocaleString("en-IN",{maximumFractionDigits:2}) : "";
  const rowE = e => e ? [e.date||"", e.party||"", fmtAmt(e.amount), e.utr||"", e.narration||"", e.ref||"", e.sourceFile||""] : Array(7).fill("");

  const sheets = [
    { name:"Reconciled",
      headers:["#","Type","Br Date","Br Party","Br Amt","Br UTR","Br Narr","HO Date","HO Party","HO Amt","HO UTR","HO Narr","Method","Score","Net","Auto?"],
      rows: R.reconciled.map((r,i) => [i+1, r.isBillsTransfer?"Bills":"Normal",
        ...rowE(r.branch).slice(0,5), ...rowE(r.ho).slice(0,5),
        r.method, r.score, r.net, r.autoConfirmed?"AUTO":"CONFIRM"]) },
    { name:"Susp→Branch",
      headers:["#","Br Date","Br Party","Br Amt","Br UTR","Br Narr","Susp Date","Susp Party","Susp Amt","Susp UTR","Susp Narr","Action","Method","Score"],
      rows: R.fromSuspense.map((r,i) => [i+1, ...rowE(r.branch).slice(0,5), ...rowE(r.suspense).slice(0,5), r.action, r.method, r.score]) },
    { name:"HO→Suspense",
      headers:["#","Type","HO Date","HO Party","HO Amt","HO UTR","HO Narr","Action","Reason"],
      rows: R.toSuspense.map((r,i) => [i+1, r.isBillsTransfer?"Bills":"Normal", ...rowE(r.ho).slice(0,5), r.action, r.reason]) },
    { name:"Wrong Ledger",
      headers:["#","Type","Br Date","Br Party","Br Amt","Br UTR","Br Narr","Reason","Suggestion"],
      rows: R.wrongLedger.map((r,i) => [i+1, r.isBillsTransfer?"Bills":"Normal", ...rowE(r.branch).slice(0,5), r.reason, r.suggestion]) },
    { name:"Date Mismatch",
      headers:["#","Status","Branch Party","Branch Amount","Journal Date","Receipt Date","Journal Month","Receipt Month","Voucher Ref","Match Method","Score"],
      rows: (R.dateMismatch||[]).map((r,i) => [
        i+1, r.approved?"APPROVED":"PENDING",
        r.branch?.party||"", fmtAmt(r.branch?.amount),
        r.journalDate||"", r.receiptDate||"",
        r.journalFYMonth||"", r.receiptFYMonth||"",
        r.tallyAlterEntry?.voucherRef||"", r.method||"", r.score||""
      ]) },
    { name:"Manual Review",
      headers:["#","Source","Date","Party","Amt","UTR","Narr","Matched Party","Matched Amt","Score","Reason"],
      rows: R.manualReview.map((r,i) => {
        const src = r.source==="Suspense-Unmatched" ? r.suspense : (r.branch||r.ho);
        const mat = r.ho||r.suspense;
        return [i+1, r.source||"", src?.date||"", src?.party||"", fmtAmt(src?.amount),
          src?.utr||"", src?.narration||"", mat?.party||"", fmtAmt(mat?.amount), r.score||"", r.reason||""];
      }),
    },
  ];
  await window.api.exportXlsx(sheets);
}

async function doExportTallyXml() {
  const R = S.results;
  const entries = [
    ...R.toSuspense.map(r => ({ ...r.tallyEntry, fromLedger: S.hoLedgerName, toLedger: S.suspenseLedgerName })),
    ...R.fromSuspense.map(r => ({ ...r.tallyEntry, fromLedger: S.suspenseLedgerName, toLedger: S.branchLedgerName })),
  ];
  await window.api.exportTallyXml(entries);
}

// ═══════════════════════════════════════════════════════════════════════════════
// BOOT
// ═══════════════════════════════════════════════════════════════════════════════
render();
