const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const fs = require("fs-extra");
const XLSX = require("xlsx");
const os = require("os");

let win;

// 설정 파일 경로
const configPath = path.join(os.homedir(), ".excel-csv-converter.json");

// 설정 불러오기
function loadSettings() {
  try {
    if (fs.existsSync(configPath)) {
      return JSON.parse(fs.readFileSync(configPath, "utf8"));
    }
  } catch (error) {
    console.log("설정 파일 불러오기 실패:", error);
  }
  return {
    lastExcelFiles: [],
    lastOutputDir: "",
  };
}

// 설정 저장
function saveSettings(settings) {
  try {
    fs.writeFileSync(configPath, JSON.stringify(settings, null, 2));
  } catch (error) {
    console.log("설정 파일 저장 실패:", error);
  }
}

function createWindow() {
  win = new BrowserWindow({
    width: 720,
    height: 850,
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
  if (process.argv.includes("--dev")) win.webContents.openDevTools();
}

app.whenReady().then(createWindow);
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// ---------- IPC ----------

// 설정 불러오기
ipcMain.handle("load-settings", async () => {
  return loadSettings();
});

ipcMain.handle("pick-excel-files", async () => {
  const settings = loadSettings();
  const r = await dialog.showOpenDialog(win, {
    title: "엑셀 파일 선택",
    properties: ["openFile", "multiSelections"],
    filters: [
      { name: "Excel", extensions: ["xlsx", "xls"] },
      { name: "All", extensions: ["*"] },
    ],
    defaultPath:
      settings.lastExcelFiles.length > 0
        ? path.dirname(settings.lastExcelFiles[0])
        : undefined,
  });
  if (r.canceled) return { success: false };

  // 선택된 파일 저장
  settings.lastExcelFiles = r.filePaths;
  saveSettings(settings);

  return { success: true, files: r.filePaths };
});

ipcMain.handle("pick-output-dir", async () => {
  const settings = loadSettings();
  const r = await dialog.showOpenDialog(win, {
    title: "CSV 출력 폴더 선택",
    properties: ["openDirectory", "createDirectory"],
    defaultPath: settings.lastOutputDir || undefined,
  });
  if (r.canceled) return { success: false };

  // 선택된 폴더 저장
  settings.lastOutputDir = r.filePaths[0];
  saveSettings(settings);

  return { success: true, dir: r.filePaths[0] };
});

const ASCII_VISIBLE_RE = /^[\x20-\x7E]+$/;
const KOREAN_RE = /[가-힣]/;

const sanitize = (name, fallback = "sheet") => {
  const n = (name || "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^A-Za-z0-9._-]+/g, "");
  return n || fallback;
};

// 워크시트에서 범위 가져오기
function getRange(ws) {
  if (!ws["!ref"]) return null;
  return XLSX.utils.decode_range(ws["!ref"]);
}

// 열이 비어있는지 확인 (헤더가 공백이거나 내용이 모두 비어있으면 true)
function isEmptyColumn(ws, col) {
  const range = getRange(ws);
  if (!range) return true;

  // 헤더(첫 번째 행) 확인
  const headerAddr = XLSX.utils.encode_cell({ r: 0, c: col });
  const headerCell = ws[headerAddr];

  // 헤더가 없거나 비어있거나 공백만 있으면 빈 열로 간주
  if (
    !headerCell ||
    headerCell.v === undefined ||
    headerCell.v === null ||
    String(headerCell.v).trim() === ""
  ) {
    return true;
  }

  // 헤더가 있으면 나머지 내용 확인 (헤더 제외)
  for (let row = 1; row <= range.e.r; row++) {
    const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
    const cell = ws[cellAddr];
    if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
      return false;
    }
  }
  return true;
}

// 헤더에 한글이 포함되어 있는지 확인
function hasKoreanHeader(ws, col) {
  const range = getRange(ws);
  if (!range) return false;

  const headerAddr = XLSX.utils.encode_cell({ r: 0, c: col });
  const headerCell = ws[headerAddr];
  if (!headerCell || !headerCell.v) return false;

  return KOREAN_RE.test(String(headerCell.v));
}

// 워크시트에서 특정 열들을 제거
function removeColumns(ws, columnsToRemove) {
  if (!ws["!ref"] || columnsToRemove.length === 0) return ws;

  const range = getRange(ws);
  if (!range) return ws;

  // 내림차순으로 정렬 (뒤에서부터 제거)
  const sortedCols = [...columnsToRemove].sort((a, b) => b - a);

  const newWs = {};
  const newCells = {};

  // 새로운 열 인덱스 매핑 생성
  const colMapping = {};
  let newColIndex = 0;

  for (let oldCol = 0; oldCol <= range.e.c; oldCol++) {
    if (!columnsToRemove.includes(oldCol)) {
      colMapping[oldCol] = newColIndex;
      newColIndex++;
    }
  }

  // 모든 셀을 새로운 위치로 복사
  for (let row = range.s.r; row <= range.e.r; row++) {
    for (let col = range.s.c; col <= range.e.c; col++) {
      if (columnsToRemove.includes(col)) continue;

      const oldAddr = XLSX.utils.encode_cell({ r: row, c: col });
      const newAddr = XLSX.utils.encode_cell({ r: row, c: colMapping[col] });

      if (ws[oldAddr]) {
        newCells[newAddr] = { ...ws[oldAddr] };
      }
    }
  }

  // 새로운 워크시트 생성
  Object.assign(newWs, newCells);

  // 새로운 범위 설정
  const newRange = {
    s: { r: range.s.r, c: 0 },
    e: { r: range.e.r, c: newColIndex - 1 },
  };

  if (newColIndex > 0) {
    newWs["!ref"] = XLSX.utils.encode_range(newRange);
  }

  // 기타 워크시트 속성 복사
  Object.keys(ws).forEach((key) => {
    if (key.startsWith("!") && key !== "!ref") {
      newWs[key] = ws[key];
    }
  });

  return newWs;
}
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

  const summary = {
    converted: 0,
    skippedSheets: 0,
    errors: [],
    removedEmptyColumns: 0,
    removedKoreanColumns: 0,
  };

  for (const file of files) {
    try {
      const wb = XLSX.readFile(file, { cellDates: true });
      const wbName = path.parse(file).name;

      for (const sheetName of wb.SheetNames) {
        if (onlyAsciiSheets && !ASCII_VISIBLE_RE.test(sheetName)) {
          summary.skippedSheets++;
          continue;
        }

        let ws = wb.Sheets[sheetName];
        const range = getRange(ws);

        if (range) {
          const columnsToRemove = [];

          // 제거할 열들 식별 (항상 적용)
          for (let col = range.s.c; col <= range.e.c; col++) {
            let shouldRemove = false;

            // 공백 열 체크 (항상 적용)
            if (isEmptyColumn(ws, col)) {
              shouldRemove = true;
              summary.removedEmptyColumns++;
            }

            // 한글 헤더 열 체크 (항상 적용)
            if (hasKoreanHeader(ws, col)) {
              shouldRemove = true;
              summary.removedKoreanColumns++;
            }

            if (shouldRemove) {
              columnsToRemove.push(col);
            }
          }

          // 열 제거 적용
          if (columnsToRemove.length > 0) {
            ws = removeColumns(ws, columnsToRemove);
          }
        }

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
