"use strict";
// ═══════════════════════════════════════════════════════════════════════════════
// Renderer — tiny UI controller around osreport.js
// ═══════════════════════════════════════════════════════════════════════════════

const O = window.OSReport;
const $ = id => document.getElementById(id);

const state = {
  debtorsFile:  null,
  debtors:      [],            // [{ customer, outstanding }]
  ledgersByCustomer: {},       // customer → entries[]
  ledgerFiles:  [],            // strings (just for display)
  asOnSerial:   null,
  asOnLabel:    "",
};

// ── Init: as-on date defaults to today ────────────────────────────────────────
{
  const d = new Date();
  const v = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
  $("asOnDate").value = v;
  state.asOnSerial = O.toSerial(v);
  state.asOnLabel  = O.serialToDDMMYYYY(state.asOnSerial);
}
$("asOnDate").addEventListener("change", () => {
  const v = $("asOnDate").value;
  state.asOnSerial = O.toSerial(v);
  state.asOnLabel  = O.serialToDDMMYYYY(state.asOnSerial);
  refreshPreview();
});

// ── File pickers ──────────────────────────────────────────────────────────────

$("btnDebtors").addEventListener("click", async () => {
  const filePath = await window.api.pickDebtors();
  if (!filePath) return;
  $("debtorsStatus").textContent = "Loading…";
  const r = await window.api.parseDebtorsFile(filePath);
  if (!r.ok) {
    $("debtorsStatus").innerHTML = `<span class="err">Error:</span> ${r.error}`;
    state.debtors = [];
  } else {
    state.debtorsFile = filePath;
    state.debtors = r.debtors;
    $("debtorsStatus").innerHTML = `<span class="ok">${r.fileName}</span> · ${r.debtors.length} debtors loaded`;
  }
  refreshPreview();
});

$("btnLedgers").addEventListener("click", async () => {
  const files = await window.api.pickLedgers();
  if (!files || !files.length) return;
  $("ledgersStatus").textContent = `Loading ${files.length} file(s)…`;
  state.ledgersByCustomer = {};
  state.ledgerFiles = [];
  let errors = 0;
  for (const f of files) {
    const r = await window.api.parseLedgerFile(f);
    if (!r.ok) { errors++; continue; }
    if (!r.ledger.customer) { errors++; continue; }
    state.ledgersByCustomer[r.ledger.customer] = r.ledger.entries;
    state.ledgerFiles.push(r.fileName);
  }
  $("ledgersStatus").innerHTML = `<span class="ok">${state.ledgerFiles.length} ledger(s) loaded</span>`
    + (errors ? ` · <span class="warn">${errors} skipped</span>` : "");
  refreshPreview();
});

$("btnLedgerFolder").addEventListener("click", async () => {
  const folder = await window.api.pickLedgerFolder();
  if (!folder) return;
  $("ledgersStatus").textContent = "Scanning folder…";
  const r = await window.api.parseLedgerFolder(folder);
  if (!r.ok) {
    $("ledgersStatus").innerHTML = `<span class="err">Error:</span> ${r.error}`;
    return;
  }
  state.ledgersByCustomer = {};
  state.ledgerFiles = [];
  for (const item of r.ledgers) {
    state.ledgersByCustomer[item.ledger.customer] = item.ledger.entries;
    state.ledgerFiles.push(item.fileName);
  }
  let msg = `<span class="ok">${state.ledgerFiles.length} ledger(s)</span>`;
  if (r.errors && r.errors.length) msg += ` · <span class="warn">${r.errors.length} skipped</span>`;
  $("ledgersStatus").innerHTML = msg;
  refreshPreview();
});

// ── Preview ───────────────────────────────────────────────────────────────────

function refreshPreview() {
  const ready = state.debtors.length > 0 && Object.keys(state.ledgersByCustomer).length > 0;
  if (!ready) {
    $("summaryCard").style.display = "none";
    $("btnGenerate").disabled = true;
    return;
  }
  // Run the generator in dry-run mode to show preview
  const result = O.generateReport({
    debtors:           state.debtors,
    ledgersByCustomer: state.ledgersByCustomer,
    asOnSerial:        state.asOnSerial,
    branchLabel:       $("branchLabel").value || "MYP & SRD",
  });
  state.lastResult = result;

  // Counts
  const okCount      = result.reconciliation.filter(r => r.status === "ok").length;
  const shortCount   = result.reconciliation.filter(r => r.status === "short").length;
  const missingCount = result.reconciliation.filter(r => r.status === "missing-ledger").length;
  $("previewCounts").innerHTML =
    `· ${result.rows.length} rows · `
    + `<span class="ok">${okCount} matched</span>`
    + (shortCount   ? ` · <span class="warn">${shortCount} short</span>` : "")
    + (missingCount ? ` · <span class="err">${missingCount} missing ledger</span>` : "");

  // Table
  const tbody = $("previewTable").querySelector("tbody");
  tbody.innerHTML = "";
  for (const r of result.reconciliation) {
    const tr = document.createElement("tr");
    const pillCls = r.status === "ok" ? "ok" : r.status === "missing-ledger" ? "missing" : "short";
    const pillTxt = r.status === "ok" ? "ok" : r.status === "missing-ledger" ? "no ledger" : "short";
    tr.innerHTML = `
      <td>${escapeHtml(r.customer)}</td>
      <td class="num">${formatNum(r.outstanding)}</td>
      <td class="num">${r.picked || ""}</td>
      <td class="num">${r.pickedSum != null ? formatNum(r.pickedSum) : ""}</td>
      <td class="num">${formatNum(r.variance)}</td>
      <td><span class="pill ${pillCls}">${pillTxt}</span></td>
    `;
    tbody.appendChild(tr);
  }
  $("summaryCard").style.display = "";
  $("btnGenerate").disabled = false;
}

$("branchLabel").addEventListener("input", refreshPreview);

// ── Generate / export ─────────────────────────────────────────────────────────

$("btnGenerate").addEventListener("click", async () => {
  if (!state.lastResult) return;
  $("generateStatus").innerHTML = "<span>Saving…</span>";
  $("btnGenerate").disabled = true;
  const r = await window.api.exportReport({
    rows:           state.lastResult.rows,
    reconciliation: state.lastResult.reconciliation,
    unmatched:      state.lastResult.unmatched,
    asOnSerial:     state.asOnSerial,
    asOnLabel:      state.asOnLabel,
  });
  $("btnGenerate").disabled = false;
  if (r.canceled) { $("generateStatus").innerHTML = "<span>Save canceled.</span>"; return; }
  if (!r.ok)      { $("generateStatus").innerHTML = `<span class="err">Error: ${r.error || "unknown"}</span>`; return; }
  $("generateStatus").innerHTML = `<span class="ok">Saved →</span> ${escapeHtml(r.path)}`;
});

// ── Utilities ─────────────────────────────────────────────────────────────────

function formatNum(n) {
  if (n == null || n === "") return "";
  const x = Number(n);
  if (!isFinite(x)) return "";
  return x.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function escapeHtml(s) {
  return String(s || "").replace(/[&<>"]/g, c =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}
