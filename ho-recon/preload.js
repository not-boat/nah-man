const { contextBridge, ipcRenderer } = require("electron");
contextBridge.exposeInMainWorld("api", {
  minimize:             () => ipcRenderer.send("win-minimize"),
  maximize:             () => ipcRenderer.send("win-maximize"),
  close:                () => ipcRenderer.send("win-close"),
  openFiles:            (label) => ipcRenderer.invoke("open-files", label),
  parseFile:            (p)     => ipcRenderer.invoke("parse-file", p),
  exportXlsx:           (data)  => ipcRenderer.invoke("export-xlsx", data),
  exportTallyXml:       (list)  => ipcRenderer.invoke("export-tally-xml", list),
  exportDateCorrections:(list)  => ipcRenderer.invoke("export-date-corrections", list),
});
