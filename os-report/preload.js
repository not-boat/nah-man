"use strict";
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  pickDebtors:        ()       => ipcRenderer.invoke("pick-debtors"),
  pickLedgers:        ()       => ipcRenderer.invoke("pick-ledgers"),
  pickLedgerFolder:   ()       => ipcRenderer.invoke("pick-ledger-folder"),
  parseDebtorsFile:   (p)      => ipcRenderer.invoke("parse-debtors-file", p),
  parseLedgerFile:    (p)      => ipcRenderer.invoke("parse-ledger-file",  p),
  parseLedgerFolder:  (p)      => ipcRenderer.invoke("parse-ledger-folder",p),
  exportReport:       (data)   => ipcRenderer.invoke("export-report", data),
});
