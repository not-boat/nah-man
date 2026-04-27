const { app, BrowserWindow, ipcMain, dialog, shell } = require("electron");
const path = require("path");
const fs   = require("fs");
const XLSX = require("xlsx");
const Refs = require("./src/refs.js");

let win;

function createWindow() {
  win = new BrowserWindow({
    width: 1380, height: 860, minWidth: 1000, minHeight: 640,
    frame: false, backgroundColor: "#07080f",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true, nodeIntegration: false,
    },
  });
  win.loadFile(path.join(__dirname, "src/index.html"));
}

app.whenReady().then(createWindow);
app.on("window-all-closed", () => { if (process.platform !== "darwin") app.quit(); });

ipcMain.on("win-minimize", () => win.minimize());
ipcMain.on("win-maximize", () => win.isMaximized() ? win.unmaximize() : win.maximize());
ipcMain.on("win-close",    () => win.close());

// ── Open files ────────────────────────────────────────────────────────────────
ipcMain.handle("open-files", async (_, label) => {
  const r = await dialog.showOpenDialog(win, {
    title: `Select: ${label}`,
    properties: ["openFile", "multiSelections"],
    filters: [{ name: "Excel / CSV / XML", extensions: ["xlsx","xls","csv","xml"] }],
  });
  return r.canceled ? [] : r.filePaths;
});

// ── Parse file → entries[] ────────────────────────────────────────────────────
ipcMain.handle("parse-file", async (_, filePath) => {
  const ext = path.extname(filePath).toLowerCase();
  try {
    if (ext === ".xml")
      return parseTallyXML(fs.readFileSync(filePath, "utf8"), filePath);
    if ([".xlsx",".xls",".csv"].includes(ext)) {
      return parseExcelRows(filePath);
    }
  } catch(e) { return { error: e.message }; }
  return [];
});

// ── Export Excel workbook ─────────────────────────────────────────────────────
ipcMain.handle("export-xlsx", async (_, sheets) => {
  const r = await dialog.showSaveDialog(win, {
    title: "Save Reconciliation Report",
    defaultPath: `HO_Branch_Recon_${ts()}.xlsx`,
    filters: [{ name: "Excel", extensions: ["xlsx"] }],
  });
  if (r.canceled) return false;

  const wb = XLSX.utils.book_new();
  for (const { name, headers, rows } of sheets) {
    if (!rows.length) continue;
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws["!cols"] = headers.map(h => ({ wch: Math.max(h.length + 2, 18) }));
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  XLSX.writeFile(wb, r.filePath);
  shell.showItemInFolder(r.filePath);
  return true;
});

// ── Export Tally XML ──────────────────────────────────────────────────────────
ipcMain.handle("export-tally-xml", async (_, entries) => {
  const r = await dialog.showSaveDialog(win, {
    title: "Save Tally Import XML",
    defaultPath: `Tally_Adjustments_${ts()}.xml`,
    filters: [{ name: "XML", extensions: ["xml"] }],
  });
  if (r.canceled) return false;

  const escX = s => String(s||"")
    .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

  const fmtDate = d => {
    if (!d) return "";
    const s = String(d).replace(/[\/\-.]/g,"");
    if (s.length !== 8) return s;
    if (parseInt(s.slice(0,4)) > 1900) return s;           // already YYYYMMDD
    return s.slice(4) + s.slice(2,4) + s.slice(0,2);       // DDMMYYYY → YYYYMMDD
  };

  const vouchers = entries.map((e, i) => `
    <VOUCHER REMOTEID="ADJ-${i+1}" VCHTYPE="Journal" ACTION="Create">
      <DATE>${fmtDate(e.date)}</DATE>
      <NARRATION>${escX(e.narration)}</NARRATION>
      <VOUCHERNUMBER>ADJ-${String(i+1).padStart(4,"0")}</VOUCHERNUMBER>
      <VOUCHERTYPENAME>Journal</VOUCHERTYPENAME>
      <ALLLEDGERENTRIES.LIST>
        <LEDGERNAME>${escX(e.fromLedger)}</LEDGERNAME>
        <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
        <AMOUNT>${Number(e.amount).toFixed(2)}</AMOUNT>
      </ALLLEDGERENTRIES.LIST>
      <ALLLEDGERENTRIES.LIST>
        <LEDGERNAME>${escX(e.toLedger)}</LEDGERNAME>
        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
        <AMOUNT>-${Number(e.amount).toFixed(2)}</AMOUNT>
      </ALLLEDGERENTRIES.LIST>
    </VOUCHER>`).join("");

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<ENVELOPE>
  <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES><SVCURRENTCOMPANY>##SVCURRENTCOMPANY</SVCURRENTCOMPANY></STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">${vouchers}
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>`;

  fs.writeFileSync(r.filePath, xml, "utf8");
  shell.showItemInFolder(r.filePath);
  return true;
});

// ── Export date correction XML ────────────────────────────────────────────────
// Generates Tally ALTER vouchers — changes only the DATE of existing journal
// entries. Uses REMOTEID to target the existing voucher by its number.
// The user must have the voucher number (ref) for Tally to find and alter it.
ipcMain.handle("export-date-corrections", async (_, entries) => {
  const r = await dialog.showSaveDialog(win, {
    title: "Save Date Correction XML",
    defaultPath: `Date_Corrections_${ts()}.xml`,
    filters: [{ name: "XML", extensions: ["xml"] }],
  });
  if (r.canceled) return false;

  const escX = s => String(s||"")
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");

  const fmtDate = d => {
    if (!d) return "";
    const s = String(d).replace(/[\/\-\.]/g,"");
    if (s.length !== 8) return s;
    if (parseInt(s.slice(0,4)) > 1900) return s;           // YYYYMMDD
    return s.slice(4) + s.slice(2,4) + s.slice(0,2);       // DDMMYYYY → YYYYMMDD
  };

  // Tally ALTER: we create a new voucher with ACTION="Alter" and the original
  // VOUCHERNUMBER so Tally finds and updates it. Only DATE changes.
  const vouchers = entries.map((e, i) => {
    const hasRef = e.voucherRef && e.voucherRef.trim();
    return `
    <VOUCHER ${hasRef ? `REMOTEID="${escX(e.voucherRef)}"` : `REMOTEID="DATE-CORR-${i+1}"`} VCHTYPE="Journal" ACTION="${hasRef ? "Alter" : "Create"}">
      <DATE>${fmtDate(e.newDate)}</DATE>
      <NARRATION>${escX(e.narration)}</NARRATION>
      ${hasRef ? `<VOUCHERNUMBER>${escX(e.voucherRef)}</VOUCHERNUMBER>` : `<VOUCHERNUMBER>DATE-CORR-${String(i+1).padStart(4,"0")}</VOUCHERNUMBER>`}
      <VOUCHERTYPENAME>Journal</VOUCHERTYPENAME>
      <ALLLEDGERENTRIES.LIST>
        <LEDGERNAME>${escX(e.fromLedger)}</LEDGERNAME>
        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
        <AMOUNT>-${Number(e.amount).toFixed(2)}</AMOUNT>
      </ALLLEDGERENTRIES.LIST>
      <ALLLEDGERENTRIES.LIST>
        <LEDGERNAME>${escX(e.toLedger)}</LEDGERNAME>
        <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
        <AMOUNT>${Number(e.amount).toFixed(2)}</AMOUNT>
      </ALLLEDGERENTRIES.LIST>
    </VOUCHER>`;
  }).join("\n");

  // Also write a human-readable summary CSV alongside the XML
  const csvLines = [
    ["#","Voucher Ref","Party/Ledger","Amount","Old Date","New Date","Action"],
    ...entries.map((e,i) => [
      i+1,
      e.voucherRef || "(no ref)",
      e.partyLedger || "",
      e.amount,
      e.oldDate || "",
      e.newDate || "",
      e.voucherRef ? "ALTER existing voucher" : "CREATE new (no voucher ref found)",
    ])
  ].map(row => row.map(c => `"${String(c).replace(/"/g,'""')}"`).join(",")).join("\r\n");

  const csvPath = r.filePath.replace(/\.xml$/i, "_summary.csv");
  fs.writeFileSync(csvPath, "\uFEFF" + csvLines, "utf8");

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<!-- HO-Branch Reconciliation: Date Corrections
     Generated: ${new Date().toISOString()}
     Entries: ${entries.length}
     
     IMPORTANT: Import this file into Tally via:
     Gateway of Tally → Import → XML Data
     
     Vouchers with a Voucher Ref will be ALTERED (date changed only).
     Vouchers without a ref will be CREATED — delete the old entry manually.
     
     A summary CSV has been saved alongside this file.
-->
<ENVELOPE>
  <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>##SVCURRENTCOMPANY</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
${vouchers}
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>`;

  fs.writeFileSync(r.filePath, xml, "utf8");
  shell.showItemInFolder(r.filePath);
  return true;
});

// ═══════════════════════════════════════════════════════════════════════════════
// PARSERS
// ═══════════════════════════════════════════════════════════════════════════════

// ── Date normaliser ───────────────────────────────────────────────────────────
function normalizeDate(d) {
  if (!d) return "";
  // Excel serial number
  if (typeof d === "number") {
    try {
      const info = XLSX.SSF.parse_date_code(d);
      if (info) return `${String(info.d).padStart(2,"0")}/${String(info.m).padStart(2,"0")}/${info.y}`;
    } catch(e) {}
  }
  const s = String(d).trim();
  // "03-Dec-25" or "03-Dec-2025"  ← Tally default export format
  const tallyFmt = s.match(/^(\d{1,2})[\/\-]([A-Za-z]{3})[\/\-](\d{2,4})$/);
  if (tallyFmt) {
    const months = {jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",
                    jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"};
    const m = months[tallyFmt[2].toLowerCase()];
    let y = tallyFmt[3];
    if (y.length === 2) y = (parseInt(y) > 50 ? "19" : "20") + y;
    return `${tallyFmt[1].padStart(2,"0")}/${m}/${y}`;
  }
  // DD/MM/YYYY or DD-MM-YYYY
  if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(s)) return s.replace(/-/g,"/");
  // YYYY-MM-DD
  if (/^\d{4}[\/\-]\d{2}[\/\-]\d{2}/.test(s)) {
    const p = s.slice(0,10).split(/[\/\-]/);
    return `${p[2]}/${p[1]}/${p[0]}`;
  }
  // YYYYMMDD or DDMMYYYY
  if (/^\d{8}$/.test(s)) {
    if (parseInt(s.slice(0,4)) > 1900) return `${s.slice(6)}/${s.slice(4,6)}/${s.slice(0,4)}`;
    return `${s.slice(0,2)}/${s.slice(2,4)}/${s.slice(4)}`;
  }
  return s;
}

// ── UTR extractor — wraps the refs library for back-compat ──────────────────
// extractUTR returns a single best-ref string for legacy callers.
// extractRefsForEntry returns the full refs[] array (and the best UTR).
function extractUTR(text, ownAmount) {
  const refs = Refs.extractRefs(text, ownAmount);
  return Refs.bestRef(refs);
}
function extractRefsForEntry(text, ownAmount) {
  return Refs.extractRefs(text, ownAmount);
}

// ── Check if a cell value looks like a narration (vs a ledger/party name) ─────
// Narration: contains digits, UTR patterns, keywords, brackets, slashes
// Ledger name: mostly uppercase letters, spaces, dots
function looksLikeNarration(val) {
  if (!val) return false;
  const s = String(val).trim();
  if (!s) return false;
  // Has digits (likely contains amount ref, UTR, invoice no)
  if (/\d/.test(s)) return true;
  // Has parentheses or dashes (common in narrations)
  if (/[\(\)\-\/]/.test(s)) return true;
  // Lowercase letters mixed in (narrations are usually mixed case)
  if (/[a-z]/.test(s)) return true;
  return false;
}

// ── MAIN PARSER: Tally paired-row Excel format ────────────────────────────────
//
// Tally exports ledgers in this exact pattern (raw sheet rows):
//   Row i  : [Date] | [Party/Ledger Name] | [Debit] | [Credit]   ← main row
//   Row i+1: [    ] | [Narration text   ] | [     ] | [       ]  ← narration row (no date, no amount)
//
// When there is NO narration, row i+1 still exists but shows
// the credited ledger name (e.g. "Cash" or "Bank Account") — same blank date/amount pattern.
// We always read both rows and put row i+1 into .narration regardless.
//
function parseTallyExcel(rawRows, filePath) {
  // rawRows is the output of sheet_to_aoa (array of arrays) — preserves empty cells
  const fname = path.basename(filePath);
  const entries = [];

  // Find which columns contain date, particulars, debit, credit
  // by scanning the header row (first non-empty row)
  let headerRowIdx = -1;
  let colDate = -1, colParticulars = -1, colDebit = -1, colCredit = -1;

  for (let i = 0; i < Math.min(rawRows.length, 20); i++) {
    const row = rawRows[i];
    for (let c = 0; c < row.length; c++) {
      const v = String(row[c]||"").toLowerCase().trim();
      if (v === "date" || v === "dt")                               colDate = c;
      if (v === "particulars" || v === "narration" || v === "name"
          || v === "ledger" || v === "description")                 colParticulars = c;
      if (v === "debit" || v === "dr" || v === "debit amount")     colDebit = c;
      if (v === "credit" || v === "cr" || v === "credit amount")   colCredit = c;
    }
    if (colParticulars !== -1 && (colDebit !== -1 || colCredit !== -1)) {
      headerRowIdx = i;
      break;
    }
  }

  // Fallback: assume standard Tally column positions if header not found
  // Tally default: A=Date, B=Particulars, C=Debit, D=Credit
  if (headerRowIdx === -1) {
    colDate = 0; colParticulars = 1; colDebit = 2; colCredit = 3;
    headerRowIdx = 0;
  }

  // Tally dual-column "Particulars" layout fix:
  // Some Tally exports lay the data out as
  //   [Date, Dr/Cr, Party Name, Vch Type, Vch No., Debit, Credit]
  // where the "Particulars" header labels the Dr/Cr column but the actual
  // party / narration content sits in the column to its RIGHT (an unlabeled
  // column). Detect this by peeking the first data row: if the cell at
  // colParticulars is just "Dr" / "Cr" / "Dr." / "Cr." and the next column
  // has real content, shift colParticulars right by one.
  for (let probe = headerRowIdx + 1; probe < Math.min(headerRowIdx + 6, rawRows.length); probe++) {
    const peek = rawRows[probe] || [];
    const v = String(peek[colParticulars] || "").trim();
    if (/^(Dr|Cr|Dr\.|Cr\.)$/i.test(v) && peek[colParticulars + 1]
        && String(peek[colParticulars + 1]).trim().length > 2) {
      colParticulars = colParticulars + 1;
      break;
    }
    if (v.length > 2) break; // first data row already has real content — layout is fine
  }

  // Process data rows in pairs starting after header
  let i = headerRowIdx + 1;
  while (i < rawRows.length) {
    const mainRow = rawRows[i];
    const narrRow = (i + 1 < rawRows.length) ? rawRows[i + 1] : null;

    const rawDate   = mainRow[colDate];
    const rawName   = mainRow[colParticulars];
    const rawDebit  = mainRow[colDebit];
    const rawCredit = mainRow[colCredit];

    // Skip completely empty rows or totals rows
    const nameStr = String(rawName || "").trim();
    if (!nameStr || /^(total|grand total|closing|opening|balance)/i.test(nameStr)) {
      i++;
      continue;
    }

    const debit  = parseFloat(String(rawDebit  || "").replace(/,/g,"")) || 0;
    const credit = parseFloat(String(rawCredit || "").replace(/,/g,"")) || 0;
    const amount = debit || credit;

    // If no amount on this row it might be the narration row of a previous entry
    // or a section header — skip it individually (we advance in pairs so this
    // case only arises for odd stray rows)
    if (amount <= 0) {
      i++;
      continue;
    }

    // Get narration from row below
    let narration = "";
    let narrRowIsNarration = false;
    if (narrRow) {
      const narrDateVal  = narrRow[colDate];
      const narrAmtDebit = parseFloat(String(narrRow[colDebit]  || "").replace(/,/g,"")) || 0;
      const narrAmtCr    = parseFloat(String(narrRow[colCredit] || "").replace(/,/g,"")) || 0;
      const narrText     = String(narrRow[colParticulars] || "").trim();

      // Row below is a narration/sub-row if: no date AND no amount
      const noDate   = !narrDateVal || String(narrDateVal).trim() === "";
      const noAmount = (narrAmtDebit === 0 && narrAmtCr === 0);

      if (noDate && noAmount && narrText) {
        narration = narrText;
        narrRowIsNarration = true;
      }
    }

    // Extract refs from narration first (most reliable), fall back to party name.
    // Pass amount so the entry's own amount is filtered out as noise.
    const refs = [
      ...extractRefsForEntry(narration, amount),
      ...extractRefsForEntry(nameStr,   amount),
    ];
    // De-duplicate refs (same value could appear in both narration and party)
    const seen = new Set();
    const dedupRefs = [];
    for (const r of refs) {
      if (seen.has(r.value)) continue;
      seen.add(r.value);
      dedupRefs.push(r);
    }
    const utr = Refs.bestRef(dedupRefs);

    entries.push({
      date:       normalizeDate(rawDate),
      party:      nameStr,
      amount,
      debitCredit: debit > 0 ? "Dr" : "Cr",
      utr,
      refs:       dedupRefs,
      narration,
      ref:        "",
      sourceFile: fname,
    });

    // Advance by 2 if we consumed the narration row, else by 1
    i += narrRowIsNarration ? 2 : 1;
  }

  return entries;
}

// ── Entry point for Excel/CSV files ──────────────────────────────────────────
function parseExcelRows(filePath) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  // Use sheet_to_aoa (array of arrays) to preserve exact cell positions
  // including empty cells — critical for the paired-row format
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  return parseTallyExcel(raw, filePath);
}

// ── Tally XML parser (fallback for XML exports) ───────────────────────────────
function parseTallyXML(text, filePath) {
  const entries = [];
  const vRe = /<VOUCHER[^>]*>([\s\S]*?)<\/VOUCHER>/gi;
  let vm;
  while ((vm = vRe.exec(text)) !== null) {
    const v = vm[1];
    const gt = tag => { const m = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`, "i").exec(v); return m ? m[1].trim() : ""; };
    const date = gt("DATE"), narration = gt("NARRATION"), vchno = gt("VOUCHERNUMBER");
    const leRe = /<ALLLEDGERENTRIES\.LIST>([\s\S]*?)<\/ALLLEDGERENTRIES\.LIST>/gi;
    let lem;
    while ((lem = leRe.exec(v)) !== null) {
      const le = lem[1];
      const gl = tag => { const m = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`, "i").exec(le); return m ? m[1].trim() : ""; };
      const raw = parseFloat(gl("AMOUNT").replace(/,/g,"")) || 0;
      const amt = Math.abs(raw);
      if (!amt) continue;
      const xmlRefs = extractRefsForEntry(narration, amt);
      entries.push({
        date: normalizeDate(date),
        party: gl("LEDGERNAME"),
        amount: amt,
        debitCredit: raw < 0 ? "Dr" : "Cr",
        utr: Refs.bestRef(xmlRefs),
        refs: xmlRefs,
        narration,
        ref: vchno,
        sourceFile: path.basename(filePath),
      });
    }
  }
  return entries;
}
