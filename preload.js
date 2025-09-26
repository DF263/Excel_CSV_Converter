const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  pickExcelFiles: () => ipcRenderer.invoke("pick-excel-files"),
  pickOutputDir: () => ipcRenderer.invoke("pick-output-dir"),
  convertExcels: (payload) => ipcRenderer.invoke("convert-excels", payload),
});
