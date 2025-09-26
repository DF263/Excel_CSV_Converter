const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const fs = require("fs-extra");
const XLSX = require("xlsx");

let win;

function createWindow() {
  win = new BrowserWindow({
    width: 720,
    height: 820,
    resizable: false,
    title: "Excel → CSV Converter",
    icon: path.join(__dirname, "icon.ico"),
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.loadFile("index.html");
  // if (process.argv.includes("--dev")) win.webContents.openDevTools();
}

app.whenReady().then(createWindow);
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// ---------- IPC ----------

ipcMain.handle("pick-excel-files", async () => {
  const r = await dialog.showOpenDialog(win, {
    title: "엑셀 파일 선택",
    properties: ["openFile", "multiSelections"],
    filters: [
      { name: "Excel", extensions: ["xlsx", "xls"] },
      { name: "All", extensions: ["*"] },
    ],
  });
  if (r.canceled) return { success: false };
  return { success: true, files: r.filePaths };
});

ipcMain.handle("pick-output-dir", async () => {
  const r = await dialog.showOpenDialog(win, {
    title: "CSV 출력 폴더 선택",
    properties: ["openDirectory", "createDirectory"],
  });
  if (r.canceled) return { success: false };
  return { success: true, dir: r.filePaths[0] };
});

const ASCII_VISIBLE_RE = /^[\x20-\x7E]+$/;
const sanitize = (name, fallback = "sheet") => {
  const n = (name || "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^A-Za-z0-9._-]+/g, "");
  return n || fallback;
};
function ensureUnique(p) {
  if (!fs.existsSync(p)) return p;
  const { dir, name, ext } = path.parse(p);
  let i = 1;
  while (true) {
    const cand = path.join(dir, `${name}_${i}${ext}`);
    if (!fs.existsSync(cand)) return cand;
    i++;
  }
}

/**
 * files: string[]
 * outDir: string
 * onlyAsciiSheets: boolean
 * delimiter: string (default ',')
 * bom: boolean (default true)
 */
ipcMain.handle("convert-excels", async (_evt, payload) => {
  const {
    files,
    outDir,
    onlyAsciiSheets = true,
    delimiter = ",",
    bom = true,
    overwriteExisting = false,
  } = payload || {};
  if (!files?.length) return { success: false, error: "파일이 비어 있습니다." };
  if (!outDir)
    return { success: false, error: "출력 폴더가 지정되지 않았습니다." };

  await fs.ensureDir(outDir);

  const summary = { converted: 0, skippedSheets: 0, errors: [] };

  for (const file of files) {
    try {
      const wb = XLSX.readFile(file, { cellDates: true });
      const wbName = path.parse(file).name;

      for (const sheetName of wb.SheetNames) {
        if (onlyAsciiSheets && !ASCII_VISIBLE_RE.test(sheetName)) {
          summary.skippedSheets++;
          continue;
        }
        const ws = wb.Sheets[sheetName];
        // to CSV
        const csv = XLSX.utils.sheet_to_csv(ws, {
          FS: delimiter,
          blankrows: false,
        });

        // 파일명: 시트명.csv (중복 시 _1, _2 ... 부여)
        const safeSheet = sanitize(sheetName, "sheet");
        const baseFile = path.join(outDir, `${safeSheet}.csv`);
        const outPath = overwriteExisting ? baseFile : ensureUnique(baseFile);

        // BOM 여부
        const data = bom
          ? Buffer.concat([
              Buffer.from([0xef, 0xbb, 0xbf]),
              Buffer.from(csv, "utf8"),
            ])
          : Buffer.from(csv, "utf8");
        await fs.writeFile(outPath, data);
        summary.converted++;
      }
    } catch (e) {
      summary.errors.push({ file, message: e.message });
    }
  }

  return { success: true, summary };
});
