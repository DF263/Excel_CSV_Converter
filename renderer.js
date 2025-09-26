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

el("btnPickFiles").addEventListener("click", async () => {
  const r = await window.api.pickExcelFiles();
  if (!r.success) return;
  selectedFiles = r.files;
  el("files").textContent = selectedFiles.length
    ? `${selectedFiles.length}개 파일 선택됨`
    : "";
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

  const { converted, skippedSheets, errors } = res.summary;
  appendLog(`성공: ${converted}`);
  appendLog(`스킵(비영문 시트): ${skippedSheets}`);
  if (errors.length) {
    appendLog(`오류 파일: ${errors.length}`);
    errors.forEach((e) => appendLog(`- ${e.file}: ${e.message}`));
  }

  setStatus("변환 완료!", errors.length ? "info" : "success");
});
