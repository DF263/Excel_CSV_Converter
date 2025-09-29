const el = (id) => document.getElementById(id);

let selectedFiles = [];
let outputDir = "";

function setStatus(msg, type = "info") {
  const s = el("status");
  s.textContent = msg || "";
  s.className = `status ${type}`;
}

function appendLog(line) {
  const log = el("log");
  log.textContent += line + "\n";
  log.scrollTop = log.scrollHeight;
}

// 파일 경로들을 표시하는 함수
function displayFilePaths(files) {
  if (!files.length) {
    el("files").textContent = "";
    return;
  }

  // 파일 경로들을 줄바꿈으로 구분하여 표시
  const pathText = files
    .map((file) => {
      // 파일명만 표시하고 전체 경로는 tooltip으로
      const fileName = file.split("\\").pop().split("/").pop();
      return fileName;
    })
    .join("\n");

  el("files").textContent = `${files.length}개 파일 선택됨:\n${pathText}`;
  el("files").title = files.join("\n"); // 전체 경로를 tooltip으로 표시
}

// 앱 시작 시 저장된 설정 불러오기
async function loadSavedSettings() {
  try {
    const settings = await window.api.loadSettings();
    if (settings.lastExcelFiles && settings.lastExcelFiles.length > 0) {
      selectedFiles = settings.lastExcelFiles;
      displayFilePaths(selectedFiles);
    }
    if (settings.lastOutputDir) {
      outputDir = settings.lastOutputDir;
      el("outDir").textContent = outputDir;
    }
  } catch (error) {
    console.log("설정 불러오기 실패:", error);
  }
}

// 페이지 로드 시 설정 불러오기
document.addEventListener("DOMContentLoaded", loadSavedSettings);

el("btnPickFiles").addEventListener("click", async () => {
  const r = await window.api.pickExcelFiles();
  if (!r.success) return;
  selectedFiles = r.files;
  displayFilePaths(selectedFiles);
});

el("btnPickOut").addEventListener("click", async () => {
  const r = await window.api.pickOutputDir();
  if (!r.success) return;
  outputDir = r.dir;
  el("outDir").textContent = outputDir;
});

el("btnConvert").addEventListener("click", async () => {
  if (!selectedFiles.length)
    return setStatus("엑셀 파일을 먼저 선택하세요.", "error");
  if (!outputDir) return setStatus("출력 폴더를 선택하세요.", "error");

  setStatus("변환 중...", "info");
  el("log").textContent = "";

  const payload = {
    files: selectedFiles,
    outDir: outputDir,
    onlyAsciiSheets: el("optAscii").checked,
    delimiter: ",",
    bom: true,
    overwriteExisting: el("optOverwrite").checked,
  };

  const res = await window.api.convertExcels(payload);
  if (!res.success) {
    setStatus(`오류: ${res.error}`, "error");
    return;
  }

  const {
    converted,
    skippedSheets,
    errors,
    removedEmptyColumns,
    removedKoreanColumns,
  } = res.summary;
  appendLog(`성공: ${converted}개 시트 변환`);
  appendLog(`스킵(비영문 시트): ${skippedSheets}개`);
  if (removedEmptyColumns > 0) {
    appendLog(`제거된 공백 열: ${removedEmptyColumns}개`);
  }
  if (removedKoreanColumns > 0) {
    appendLog(`제거된 한글 헤더 열: ${removedKoreanColumns}개`);
  }
  if (errors.length) {
    appendLog(`오류 파일: ${errors.length}개`);
    errors.forEach((e) => appendLog(`- ${e.file}: ${e.message}`));
  }

  setStatus("변환 완료!", errors.length ? "info" : "success");
});
