const cashbookInput = document.getElementById("cashbookFile");
const runBtn = document.getElementById("runBtn");
const statusEl = document.getElementById("status");
const summaryCard = document.getElementById("summaryCard");
const tableCard = document.getElementById("tableCard");
const tbody = document.querySelector("#resultTable tbody");
const filterInput = document.getElementById("filterInput");
const downloadBtn = document.getElementById("downloadBtn");
const downloadCsvBtn = document.getElementById("downloadCsvBtn");

const state = {
  cashbookFile: null,
  parsedRows: [],
  resultRows: []
};

cashbookInput.addEventListener("change", () => {
  state.cashbookFile = cashbookInput.files?.[0] || null;
  updateReadyState();
});

runBtn.addEventListener("click", runAudit);
filterInput.addEventListener("input", () => renderTable(state.resultRows));
downloadBtn.addEventListener("click", downloadXlsx);
downloadCsvBtn.addEventListener("click", downloadCsv);

function updateReadyState() {
  runBtn.disabled = !state.cashbookFile;
  if (state.cashbookFile) {
    statusEl.textContent = "파일 선택 완료. 검토 실행 버튼을 누르세요.";
  } else {
    statusEl.textContent = "현금출납부 파일을 선택하면 검토를 실행할 수 있습니다.";
  }
}

async function runAudit() {
  try {
    statusEl.textContent = "현금출납부를 읽는 중입니다...";
    state.parsedRows = await parseCashbook(state.cashbookFile);
    statusEl.textContent = `${state.parsedRows.length.toLocaleString()}건을 읽었습니다. 제목 기준 검토 중입니다...`;
    state.resultRows = state.parsedRows
      .map(analyzeRow)
      .filter(Boolean);
    renderSummary();
    renderTable(state.resultRows);
    statusEl.textContent = `검토 완료: ${state.resultRows.length.toLocaleString()}건이 의심·확인 대상으로 선별되었습니다.`;
  } catch (err) {
    console.error(err);
    statusEl.textContent = `검토 중 오류가 발생했습니다: ${err.message || err}`;
  }
}

async function parseCashbook(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellDates: true, dense: false });
  let collected = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    const headerIndex = findHeaderRow(matrix);
    if (headerIndex < 0) continue;
    const headers = normalizeHeaders(matrix[headerIndex]);
    const rows = matrix.slice(headerIndex + 1).map((row, idx) => rowToObject(row, headers, sheetName, headerIndex + idx + 2));
    const mapped = rows.map(mapCashbookRow).filter(row => row.title || row.currentAccount || row.amount);
    collected = collected.concat(mapped);
  }

  if (collected.length === 0) {
    throw new Error("거래 데이터를 찾지 못했습니다. 제목, 금액, 목 관련 열이 포함되어 있는지 확인하세요.");
  }
  return collected;
}

function findHeaderRow(matrix) {
  let bestIndex = -1;
  let bestScore = 0;
  const keys = ["일자", "날짜", "제목", "적요", "내용", "세부", "항목", "목", "비목", "수입", "지출", "금액", "결의", "사업", "재원"];
  matrix.slice(0, 40).forEach((row, idx) => {
    const joined = row.map(v => normalize(v)).join(" ");
    let score = keys.reduce((acc, key) => acc + (joined.includes(key) ? 1 : 0), 0);
    if (row.filter(v => String(v).trim()).length >= 4) score += 1;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = idx;
    }
  });
  return bestScore >= 3 ? bestIndex : -1;
}

function normalizeHeaders(row) {
  const seen = new Map();
  return row.map((value, idx) => {
    let header = String(value || `열${idx + 1}`).replace(/\n/g, " ").replace(/\s+/g, " ").trim();
    if (!header) header = `열${idx + 1}`;
    const count = seen.get(header) || 0;
    seen.set(header, count + 1);
    return count ? `${header}_${count + 1}` : header;
  });
}

function rowToObject(row, headers, sheetName, rowNumber) {
  const obj = { __sheet: sheetName, __rowNumber: rowNumber };
  headers.forEach((header, idx) => { obj[header] = row[idx] ?? ""; });
  return obj;
}

function mapCashbookRow(raw) {
  const headers = Object.keys(raw);
  const get = (...patterns) => getField(raw, headers, patterns);
  const income = parseAmount(get("수입금액", "수입", "입금", "차변"));
  const expense = parseAmount(get("지출금액", "지출", "출금", "대변"));
  const genericAmount = parseAmount(get("금액", "거래금액", "결의금액"));

  const typeText = String(get("수입지출", "구분", "결의구분", "수지구분") || "");
  let type = "";
  if (expense > 0 || /지출|출금|지급|대변/.test(typeText)) type = "지출";
  else if (income > 0 || /수입|입금|수납|차변/.test(typeText)) type = "수입";
  else type = genericAmount < 0 ? "지출" : "수입·지출 미상";

  const titleParts = [
    get("제목", "품의제목", "적요", "내용", "거래내용", "결의제목"),
    get("세부사업", "사업명", "세부항목명", "세부항목"),
    get("거래처", "지급처", "수납자", "비고")
  ].filter(Boolean);

  const currentAccount = get("원가비목", "원가비목명", "비목", "목", "세목", "세부항목명", "세부항목", "과목명", "예산과목");

  return {
    sheet: raw.__sheet,
    rowNumber: raw.__rowNumber,
    date: get("일자", "날짜", "결의일", "거래일", "수납일", "지급일"),
    type,
    currentAccount: String(currentAccount || "").trim(),
    title: cleanText(titleParts.join(" / ")),
    amount: expense || income || Math.abs(genericAmount) || 0,
    raw
  };
}

function getField(raw, headers, patterns) {
  const normalizedPatterns = patterns.map(normalize);
  const scored = headers.map(header => {
    const nh = normalize(header);
    let score = 0;
    normalizedPatterns.forEach(pattern => {
      if (nh === pattern) score += 100;
      else if (nh.includes(pattern)) score += 30 + pattern.length;
      else if (pattern.includes(nh) && nh.length > 1) score += 10;
    });
    return { header, score };
  }).filter(x => x.score > 0).sort((a, b) => b.score - a.score);
  if (!scored.length) return "";
  return raw[scored[0].header];
}

function analyzeRow(row) {
  const targetType = row.type.includes("지출") ? "expense" : (row.type.includes("수입") ? "income" : "both");
  const text = normalize(`${row.title} ${row.currentAccount}`);
  const currentNorm = normalize(row.currentAccount);
  let best = null;

  for (const rule of AUDIT_RULES) {
    if (!(rule.type === "both" || rule.type === targetType)) continue;
    if (containsAny(text, rule.avoid || [])) continue;
    let score = 0;
    if (rule.anyAll) {
      const hits = (rule.keywords || []).filter(k => text.includes(normalize(k)));
      score = hits.length >= 2 ? hits.length * 12 : 0;
    } else {
      for (const keyword of rule.keywords || []) {
        const nk = normalize(keyword);
        if (text.includes(nk)) score += Math.min(40, 8 + nk.length);
      }
    }
    if (score <= 0) continue;

    const expectedMatchesCurrent = accountMatches(rule.expected, currentNorm);
    const negativeMatch = containsAny(currentNorm, rule.currentNegative || []);

    if (rule.severity !== "adjust" && expectedMatchesCurrent && !negativeMatch) {
      continue;
    }

    if (negativeMatch) score += 30;
    if (!expectedMatchesCurrent) score += 12;
    if (!row.currentAccount) score += 8;

    if (!best || score > best.score) best = { rule, score };
  }

  if (!best) return null;

  const category = best.rule.severity === "warn"
    ? "목 부적정 의심"
    : best.rule.severity === "adjust"
      ? "과목경정·반납 확인"
      : "추가 확인 필요";

  return {
    원본시트: row.sheet,
    원본행: row.rowNumber,
    일자: formatCell(row.date),
    구분: row.type,
    현재목: row.currentAccount || "미확인",
    제목적요: row.title,
    금액: row.amount,
    판단구분: category,
    권장확인목: best.rule.expected,
    검토의견: best.rule.opinion,
    적용규칙: best.rule.id
  };
}


function renderSummary() {
  summaryCard.hidden = false;
  tableCard.hidden = false;
  const total = state.parsedRows.length;
  const flagged = state.resultRows.length;
  const amount = state.resultRows.reduce((sum, row) => sum + Number(row.금액 || 0), 0);
  document.getElementById("totalRows").textContent = total.toLocaleString();
  document.getElementById("flagRows").textContent = flagged.toLocaleString();
  document.getElementById("amountTotal").textContent = `${amount.toLocaleString()}원`;
}

function renderTable(rows) {
  const keyword = normalize(filterInput.value || "");
  const visible = keyword
    ? rows.filter(row => normalize(Object.values(row).join(" ")).includes(keyword))
    : rows;
  tbody.innerHTML = visible.map(row => `
    <tr>
      <td>${escapeHtml(row.원본행)}</td>
      <td>${escapeHtml(row.일자)}</td>
      <td>${escapeHtml(row.구분)}</td>
      <td>${escapeHtml(row.현재목)}</td>
      <td>${escapeHtml(row.제목적요)}</td>
      <td class="amount">${Number(row.금액 || 0).toLocaleString()}</td>
      <td>${badge(row.판단구분)}</td>
      <td>${escapeHtml(row.권장확인목)}</td>
      <td>${escapeHtml(row.검토의견)}</td>
    </tr>
  `).join("");
}

function badge(value) {
  const cls = value.includes("부적정") ? "warn" : value.includes("반납") ? "adjust" : "check";
  return `<span class="badge ${cls}">${escapeHtml(value)}</span>`;
}

function downloadXlsx() {
  if (!state.resultRows.length) return;
  const summary = summarizeResults(state.resultRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "요약");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.resultRows), "검토결과");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ 안내: "본 결과는 제목 기준 의심 항목 선별 자료이며, 최종 판단은 증빙 확인 후 결정해야 합니다." }]), "안내");
  XLSX.writeFile(wb, `현금출납부_목별_검토결과_${dateStamp()}.xlsx`);
}

function downloadCsv() {
  if (!state.resultRows.length) return;
  const ws = XLSX.utils.json_to_sheet(state.resultRows);
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `현금출납부_목별_검토결과_${dateStamp()}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function summarizeResults(rows) {
  const groups = new Map();
  for (const row of rows) {
    const key = row.판단구분;
    const current = groups.get(key) || { 구분: key, 건수: 0, 금액: 0 };
    current.건수 += 1;
    current.금액 += Number(row.금액 || 0);
    groups.set(key, current);
  }
  return Array.from(groups.values());
}

function accountMatches(expected, currentNorm) {
  if (!currentNorm) return false;
  const synonyms = ACCOUNT_SYNONYMS[expected] || [expected];
  return synonyms.some(word => currentNorm.includes(normalize(word)));
}

function containsAny(text, words) {
  return words.some(word => text.includes(normalize(word)));
}

function parseAmount(value) {
  if (typeof value === "number") return Math.abs(value);
  const text = String(value || "").replace(/,/g, "").replace(/원/g, "").trim();
  if (!text) return 0;
  const n = Number(text.replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? Math.abs(n) : 0;
}

function normalize(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()\[\]{}·ㆍ,._\-:;\/\\]/g, "");
}

function cleanText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function formatCell(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  return cleanText(value);
}

function dateStamp() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,"0")}${String(d.getDate()).padStart(2,"0")}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
