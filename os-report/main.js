"use strict";
// ═══════════════════════════════════════════════════════════════════════════════
// O/S Report Maker — Electron main process
// ═══════════════════════════════════════════════════════════════════════════════
//
// IPC handlers for:
//   - File / folder pickers (debtors export, ledger files / folder)
//   - Parsing those files (xlsx → AoA → osreport.js)
//   - Writing the final report Excel
//
// All real logic lives in src/osreport.js (pure, unit-tested). This file only
// glues Electron to that module.

const { app, BrowserWindow, ipcMain, dialog, shell } = require("electron");
const path  = require("path");
const fs    = require("fs");
const XLSX  = require("xlsx");
const O     = require("./src/osreport.js");

let win;
function createWindow() {
  win = new BrowserWindow({
    width: 1100, height: 760, minWidth: 900, minHeight: 600,
    backgroundColor: "#07080f",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true, nodeIntegration: false,
    },
  });
  win.removeMenu();
  win.loadFile(path.join(__dirname, "src", "index.html"));
}
app.whenReady().then(createWindow);
app.on("window-all-closed", () => { if (process.platform !== "darwin") app.quit(); });
app.on("activate",          () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });

// ── Helpers ────────────────────────────────────────────────────────────────────

function readSheetAsAoA(filePath, sheetName) {
  const wb = XLSX.readFile(filePath, { cellDates: false, raw: true });
  const name = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(`Sheet "${name}" not found in ${path.basename(filePath)}`);
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
}

function ts() {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}

// ── File pickers ───────────────────────────────────────────────────────────────

ipcMain.handle("pick-debtors", async () => {
  const r = await dialog.showOpenDialog(win, {
    title: "Select Debtors export (Tally Outstandings → Group → Sundry Debtors)",
    properties: ["openFile"],
    filters: [{ name: "Excel / CSV", extensions: ["xlsx", "xls", "csv"] }],
  });
  return r.canceled ? null : r.filePaths[0];
});

ipcMain.handle("pick-ledgers", async () => {
  const r = await dialog.showOpenDialog(win, {
    title: "Select one or more Ledger Excel files",
    properties: ["openFile", "multiSelections"],
    filters: [{ name: "Excel / CSV", extensions: ["xlsx", "xls", "csv"] }],
  });
  return r.canceled ? [] : r.filePaths;
});

ipcMain.handle("pick-ledger-folder", async () => {
  const r = await dialog.showOpenDialog(win, {
    title: "Select folder containing one Ledger Excel per customer",
    properties: ["openDirectory"],
  });
  return r.canceled ? null : r.filePaths[0];
});

// ── Parsers ────────────────────────────────────────────────────────────────────

ipcMain.handle("parse-debtors-file", async (_, filePath) => {
  try {
    const aoa = readSheetAsAoA(filePath);
    const debtors = O.parseDebtorsAoA(aoa);
    return { ok: true, debtors, fileName: path.basename(filePath) };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

ipcMain.handle("parse-ledger-file", async (_, filePath) => {
  try {
    const aoa = readSheetAsAoA(filePath);
    const parsed = O.parseLedgerAoA(aoa, path.basename(filePath));
    return { ok: true, ledger: parsed, fileName: path.basename(filePath) };
  } catch (e) {
    return { ok: false, error: e.message, fileName: path.basename(filePath) };
  }
});

ipcMain.handle("parse-ledger-folder", async (_, folderPath) => {
  try {
    const stat = fs.statSync(folderPath);
    if (!stat.isDirectory()) return { ok: false, error: "Not a directory" };
    const files = fs.readdirSync(folderPath)
      .filter(f => /\.(xlsx|xls|csv)$/i.test(f))
      .map(f => path.join(folderPath, f));
    const ledgers = [];
    const errors  = [];
    for (const f of files) {
      try {
        const aoa = readSheetAsAoA(f);
        const parsed = O.parseLedgerAoA(aoa, path.basename(f));
        if (!parsed.customer) {
          errors.push({ file: path.basename(f), error: "Could not detect customer name" });
          continue;
        }
        ledgers.push({ fileName: path.basename(f), ledger: parsed });
      } catch (e) {
        errors.push({ file: path.basename(f), error: e.message });
      }
    }
    return { ok: true, ledgers, errors };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

// ── Report writer ──────────────────────────────────────────────────────────────
//
// Output workbook layout:
//   Sheet "MYP&SRD"         — main report, mirrors the user's template:
//     row 0: title          "PARTS AND SERVICE CUSTOMER'S OUTSTANDING AS ON DD-MM-YYYY"
//     row 1: header         Branches | Customer Name | Invoice No | Invoice Date
//                            | Due Days | 0-30 | 31-60 | 61-90 | 91-120 | 121-Above
//                            | Grand Total
//     row 2..n: data rows, sorted by ageing bucket then due days
//
//   Sheet "Reconciliation"  — per-customer audit:
//     Customer | Outstanding | Picked Rows | Picked Sum | Variance | Status
//
//   Sheet "Unmatched"       — debtors with no ledger file
//     Customer | Outstanding | Reason
//
ipcMain.handle("export-report", async (_, payload) => {
  const { rows, reconciliation, unmatched, asOnSerial, asOnLabel } = payload;
  const r = await dialog.showSaveDialog(win, {
    title: "Save Outstanding Report",
    defaultPath: `Outstanding_Report_${ts()}.xlsx`,
    filters: [{ name: "Excel", extensions: ["xlsx"] }],
  });
  if (r.canceled) return { ok: false, canceled: true };

  const wb = XLSX.utils.book_new();
  const onLabel = asOnLabel || (asOnSerial ? O.serialToDDMMYYYY(asOnSerial) : "");

  // ── MYP&SRD sheet ────────────────────────────────────────────────────────
  const headers = [
    "Branches", "Customer Name", "Invoice No", "Invoice Date", "Due Days",
    "0-30", "31-60", "61-90", "91-120", "121-Above", "Grand Total",
  ];
  const titleRow = [`PARTS AND SERVICE CUSTOMER'S OUTSTANDING AS ON ${onLabel}`,
                     "", "", "", "", "", "", "", "", "", ""];
  const dataRows = rows.map(r => [
    r.branches, r.customer,
    r.invoiceNo,
    r.invoiceDate,
    r.dueDays,
    r.b0_30, r.b31_60, r.b61_90, r.b91_120, r.b121,
    r.grandTotal,
  ]);
  // Append a totals row at the bottom: blanks except for column-sum totals
  const sumCols = [5, 6, 7, 8, 9, 10];
  const totals = ["", "", "", "", "Total"];
  for (const c of sumCols) {
    let s = 0;
    for (const dr of dataRows) {
      const v = Number(dr[c]);
      if (isFinite(v)) s += v;
    }
    totals.push(s ? Math.round(s * 100) / 100 : "");
  }

  const myp = XLSX.utils.aoa_to_sheet([titleRow, headers, ...dataRows, totals]);
  myp["!cols"] = [
    { wch: 12 }, { wch: 36 }, { wch: 14 }, { wch: 14 }, { wch:  9 },
    { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
  ];
  // Apply Excel-date number-format to the Invoice Date column
  for (let i = 2; i < 2 + dataRows.length; i++) {
    const cellRef = XLSX.utils.encode_cell({ r: i, c: 3 });
    const cell = myp[cellRef];
    if (cell && typeof cell.v === "number") cell.z = "dd-mm-yyyy";
  }
  XLSX.utils.book_append_sheet(wb, myp, "MYP&SRD");

  // ── Reconciliation sheet ─────────────────────────────────────────────────
  const recHeaders = ["Customer", "Outstanding (Debtors)", "Picked Rows",
                       "Picked Sum (FIFO)", "Variance", "Status", "Notes"];
  const recRows = reconciliation.map(r => [
    r.customer, r.outstanding, r.picked, r.pickedSum != null ? r.pickedSum : "",
    r.variance, r.status,
    r.status === "ok"           ? "Matched exactly via FIFO"
    : r.status === "missing-ledger" ? "No ledger file found for this customer"
    : r.status === "short"      ? "FIFO total fell short — ledger may be incomplete or older debits needed"
    : "",
  ]);
  const recWs = XLSX.utils.aoa_to_sheet([recHeaders, ...recRows]);
  recWs["!cols"] = [
    { wch: 36 }, { wch: 16 }, { wch: 11 }, { wch: 16 }, { wch: 14 }, { wch: 16 }, { wch: 50 },
  ];
  XLSX.utils.book_append_sheet(wb, recWs, "Reconciliation");

  // ── Unmatched sheet ──────────────────────────────────────────────────────
  if (unmatched && unmatched.length) {
    const unHeaders = ["Customer", "Outstanding", "Reason"];
    const unRows = unmatched.map(u => [u.customer, u.outstanding, u.reason]);
    const unWs = XLSX.utils.aoa_to_sheet([unHeaders, ...unRows]);
    unWs["!cols"] = [{ wch: 36 }, { wch: 16 }, { wch: 50 }];
    XLSX.utils.book_append_sheet(wb, unWs, "Unmatched");
  }

  XLSX.writeFile(wb, r.filePath);
  shell.showItemInFolder(r.filePath);
  return { ok: true, path: r.filePath };
});
