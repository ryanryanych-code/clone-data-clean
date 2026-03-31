const MERGED_COLUMN = "合并项";
const EMPTY_VALUE = "未填写";
const TOP_SUPPLIER_COUNT = 5;
const TOP_REASON_COUNT = 10;
const MAX_PIE_SEGMENTS = 8;

const fileInput = document.getElementById("fileInput");
const rowCount = document.getElementById("rowCount");
const groupCount = document.getElementById("groupCount");
const statusMessage = document.getElementById("statusMessage");
const statsTable = document.getElementById("statsTable");
const supplierTable = document.getElementById("supplierTable");
const defectTopTable = document.getElementById("defectTopTable");
const issueDetailTable = document.getElementById("issueDetailTable");
const issueDetailTitle = document.getElementById("issueDetailTitle");
const exportFilterBtn = document.getElementById("exportFilterBtn");
const toggleUploadPanelBtn = document.getElementById("toggleUploadPanelBtn");
const uploadPanel = document.getElementById("uploadPanel");
const modelProjectPieCanvas = document.getElementById("modelProjectPie");
const reasonPieCanvas = document.getElementById("reasonPie");
const progressPieCanvas = document.getElementById("progressPie");
const progressStackChartCanvas = document.getElementById("progressStackChart");
const defectTrendChartCanvas = document.getElementById("defectTrendChart");

const timeRangeFilter = document.getElementById("timeRangeFilter");
const modelFilter = document.getElementById("modelFilter");
const supplierFilter = document.getElementById("supplierFilter");
const statusFilter = document.getElementById("statusFilter");

let statsRows = [];
let allRows = [];
let filteredRows = [];
let hasDateData = false;
let filterBaseDate = null;
let modelProjectPieChart = null;
let reasonPieChart = null;
let progressPieChart = null;
let progressStackChart = null;
let defectTrendChart = null;

if (window.Chart) {
  Chart.defaults.color = "#dfeeff";
  Chart.defaults.font.family = '"Noto Sans SC", sans-serif';
  Chart.defaults.font.size = 12;
}

fileInput.addEventListener("change", async (event) => {
  const [file] = event.target.files || [];
  if (!file) {
    return;
  }

  resetBeforeProcess(file.name);

  try {
    const rawRows = await readFile(file);
    processRows(rawRows, file.name);
  } catch (error) {
    clearAll(error.message || "文件处理失败，请检查文件格式。");
  }
});

[timeRangeFilter, modelFilter, supplierFilter, statusFilter].forEach((select) => {
  select.addEventListener("change", () => {
    if (!allRows.length) {
      return;
    }
    applyFilters();
  });
});

exportFilterBtn.addEventListener("click", () => {
  if (statsRows.length === 0) {
    return;
  }
  exportWorkbook(statsRows, "筛选后合并项统计.xlsx", "统计结果");
});

if (toggleUploadPanelBtn && uploadPanel) {
  toggleUploadPanelBtn.addEventListener("click", () => {
    uploadPanel.classList.toggle("is-collapsed");
    const expanded = !uploadPanel.classList.contains("is-collapsed");
    toggleUploadPanelBtn.setAttribute("aria-expanded", expanded ? "true" : "false");
  });
}

tryAutoLoadData();

function processRows(rawRows, sourceName = "数据文件") {
  if (!Array.isArray(rawRows) || rawRows.length === 0) {
    throw new Error("文件中没有可处理的数据。");
  }

  const columns = resolveColumns(rawRows);
  allRows = buildNormalizedRows(rawRows, columns);

  initializeFilters(allRows);
  applyFilters();

  statusMessage.textContent = `处理完成：${sourceName}，共 ${allRows.length} 行数据，已生成筛选结果。`;
}

function tryAutoLoadData() {
  if (!window.AUTOLOAD_CSV_BASE64) {
    return;
  }

  const sourceName = window.AUTOLOAD_SOURCE_NAME || "自动数据";
  resetBeforeProcess(sourceName);

  try {
    const csvText = decodeBase64Utf8(window.AUTOLOAD_CSV_BASE64);
    const rawRows = readRowsFromCsvText(csvText);
    processRows(rawRows, sourceName);
  } catch (error) {
    clearAll(error.message || "自动加载数据失败，请检查自动数据文件。");
  }
}

function decodeBase64Utf8(base64Text) {
  const binaryText = atob(String(base64Text || ""));
  const bytes = Uint8Array.from(binaryText, (char) => char.charCodeAt(0));
  return new TextDecoder("utf-8").decode(bytes);
}

function readRowsFromCsvText(csvText) {
  if (!String(csvText || "").trim()) {
    throw new Error("自动数据为空。");
  }

  try {
    const workbook = XLSX.read(csvText, { type: "string" });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      throw new Error("自动数据中没有可读取的工作表。");
    }
    const sheet = workbook.Sheets[firstSheetName];
    return XLSX.utils.sheet_to_json(sheet, { defval: "" });
  } catch {
    throw new Error("自动数据解析失败，请检查 CSV 格式。");
  }
}

function initializeFilters(rows) {
  const dateList = rows
    .map((row) => row.dateObj)
    .filter((date) => date instanceof Date && !Number.isNaN(date.getTime()));
  hasDateData = dateList.length > 0;
  filterBaseDate = hasDateData ? new Date(Math.max(...dateList.map((date) => date.getTime()))) : null;

  populateSelect(modelFilter, uniqueValues(rows.map((row) => row.model)), "全部");
  populateSelect(supplierFilter, uniqueValues(rows.map((row) => row.supplier)), "全部");

  timeRangeFilter.value = hasDateData ? "this_month" : "all";
  timeRangeFilter.disabled = !hasDateData;
  statusFilter.value = "all";

  setFilterControlsEnabled(true);
}

function applyFilters() {
  filteredRows = allRows.filter((row) => {
    return (
      matchesTimeRange(row, timeRangeFilter.value) &&
      matchesSelect(row.model, modelFilter.value) &&
      matchesSelect(row.supplier, supplierFilter.value) &&
      matchesStatus(row.progress, statusFilter.value)
    );
  });

  statsRows = buildStats(filteredRows);

  const modelProjectData = buildModelProjectData(filteredRows);
  const reasonPieData = buildReasonPieData(filteredRows);
  const progressPieData = buildProgressPieData(filteredRows);
  const trendData = buildTrendData(filteredRows, timeRangeFilter.value);
  const supplierDetailRows = buildTopSupplierDetails(filteredRows);
  const defectTopRows = buildTopDefectDetails(filteredRows);
  const reasonParetoData = buildReasonParetoData(filteredRows);

  rowCount.textContent = `${filteredRows.length} / ${allRows.length}`;
  groupCount.textContent = String(statsRows.length);

  renderTable(statsTable, withTotalRow(statsRows, "合并项", "数量"));
  renderTable(supplierTable, withTotalRow(supplierDetailRows, "供应商", "项目数"));
  renderTable(defectTopTable, withTotalRow(defectTopRows, "不良名称", "项目数"));
  renderTable(issueDetailTable, []);
  issueDetailTitle.textContent = "问题点明细";

  renderModelProjectPie(modelProjectData);
  renderReasonPie(reasonPieData);
  renderProgressPie(progressPieData);
  renderDefectTrend(trendData, filteredRows);
  renderReasonPareto(reasonParetoData, filteredRows);

  exportFilterBtn.disabled = statsRows.length === 0;

  const filterDesc = [
    `时间=${displayTimeLabel(timeRangeFilter.value)}`,
    `车型=${displaySelectLabel(modelFilter)}`,
    `供应商=${displaySelectLabel(supplierFilter)}`,
    `状态=${displayStatusLabel(statusFilter.value)}`,
  ].join("，");

  statusMessage.textContent = `当前筛选：${filterDesc}。命中 ${filteredRows.length} 行。`;
}

function resetBeforeProcess(fileName) {
  statusMessage.textContent = `正在处理 ${fileName}...`;
  rowCount.textContent = "0";
  groupCount.textContent = "0";
  renderTable(statsTable, []);
  renderTable(supplierTable, []);
  renderTable(defectTopTable, []);
  renderTable(issueDetailTable, []);
  issueDetailTitle.textContent = "问题点明细";
  clearCharts();
  setFilterControlsEnabled(false);
  exportFilterBtn.disabled = true;
}

function clearAll(message) {
  statsRows = [];
  allRows = [];
  filteredRows = [];
  rowCount.textContent = "0";
  groupCount.textContent = "0";
  renderTable(statsTable, []);
  renderTable(supplierTable, []);
  renderTable(defectTopTable, []);
  renderTable(issueDetailTable, []);
  issueDetailTitle.textContent = "问题点明细";
  clearCharts();
  hasDateData = false;
  filterBaseDate = null;
  setFilterControlsEnabled(false);
  exportFilterBtn.disabled = true;
  statusMessage.textContent = message;
}

function setFilterControlsEnabled(enabled) {
  modelFilter.disabled = !enabled;
  supplierFilter.disabled = !enabled;
  statusFilter.disabled = !enabled;
  exportFilterBtn.disabled = !enabled;
  timeRangeFilter.disabled = !enabled || !hasDateData;
}

function populateSelect(selectElement, values, allLabel) {
  const currentValue = selectElement.value;
  selectElement.innerHTML = "";

  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = allLabel;
  selectElement.appendChild(allOption);

  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    selectElement.appendChild(option);
  });

  if ([...selectElement.options].some((opt) => opt.value === currentValue)) {
    selectElement.value = currentValue;
  } else {
    selectElement.value = "all";
  }
}

function uniqueValues(values) {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b, "zh-CN"));
}

function matchesSelect(value, selected) {
  if (selected === "all") {
    return true;
  }
  return value === selected;
}

function matchesStatus(progress, selected) {
  if (selected === "all") {
    return true;
  }
  return progress === selected;
}

function matchesTimeRange(row, selectedRange) {
  if (selectedRange === "all") {
    return true;
  }
  if (!hasDateData) {
    return true;
  }

  if (!row.dateObj) {
    return false;
  }

  const now = filterBaseDate ? new Date(filterBaseDate) : new Date();
  const target = row.dateObj;

  if (selectedRange === "this_year") {
    return target.getFullYear() === now.getFullYear();
  }

  if (selectedRange === "this_quarter") {
    const quarterStartMonth = Math.floor(now.getMonth() / 3) * 3;
    const quarterStart = new Date(now.getFullYear(), quarterStartMonth, 1);
    const quarterEnd = new Date(now.getFullYear(), quarterStartMonth + 3, 1);
    return target >= quarterStart && target < quarterEnd;
  }

  if (selectedRange === "this_month") {
    return target.getFullYear() === now.getFullYear() && target.getMonth() === now.getMonth();
  }

  if (selectedRange === "this_week") {
    const mondayOffset = (now.getDay() + 6) % 7;
    const weekStart = new Date(now);
    weekStart.setDate(now.getDate() - mondayOffset);
    weekStart.setHours(0, 0, 0, 0);

    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 7);

    return target >= weekStart && target < weekEnd;
  }

  return true;
}

function displayTimeLabel(value) {
  if (value === "this_week") return "本周";
  if (value === "this_month") return "本月";
  if (value === "this_quarter") return "本季度";
  if (value === "this_year") return "本年";
  return "全部";
}

function displayStatusLabel(value) {
  if (value === "已关闭") return "已关闭";
  if (value === "解析中") return "解析中";
  if (value === "对策中") return "对策中";
  return "全部";
}

function displaySelectLabel(selectElement) {
  const selectedOption = selectElement.options[selectElement.selectedIndex];
  return selectedOption ? selectedOption.textContent : "全部";
}

function clearCharts() {
  if (modelProjectPieChart) {
    modelProjectPieChart.destroy();
    modelProjectPieChart = null;
  }
  if (reasonPieChart) {
    reasonPieChart.destroy();
    reasonPieChart = null;
  }
  if (progressPieChart) {
    progressPieChart.destroy();
    progressPieChart = null;
  }
  if (progressStackChart) {
    progressStackChart.destroy();
    progressStackChart = null;
  }
  if (defectTrendChart) {
    defectTrendChart.destroy();
    defectTrendChart = null;
  }
}

function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = event.target.result;
        const workbook =
          file.name.toLowerCase().endsWith(".csv")
            ? XLSX.read(data, { type: "string" })
            : XLSX.read(data, { type: "array" });

        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) {
          reject(new Error("文件中没有可读取的工作表。"));
          return;
        }

        const sheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        resolve(rows);
      } catch {
        reject(new Error("读取文件失败，请确认上传的是标准 Excel 或 CSV 文件。"));
      }
    };

    reader.onerror = () => reject(new Error("文件读取失败，请稍后重试。"));

    if (file.name.toLowerCase().endsWith(".csv")) {
      reader.readAsText(file, "utf-8");
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
}

function resolveColumns(rows) {
  const allColumns = collectColumns(rows);
  const required = {
    inspection: pickColumn(allColumns, ["检测项名称"], /检测项|项目名称|inspection/i, "检测项名称"),
    defect: pickColumn(allColumns, ["不良名称"], /不良名称|缺陷名称|defect/i, "不良名称"),
    model: pickColumn(allColumns, ["车型"], /车型|model/i, "车型"),
    supplier: pickColumn(allColumns, ["供应商"], /供应商|supplier/i, "供应商"),
    level: pickColumn(allColumns, ["不良等级", "等级"], /等级|level/i, "不良等级"),
    progress: pickColumn(allColumns, ["解决进度"], /解决进度|进度|progress/i, "解决进度"),
  };

  const reason = pickColumn(
    allColumns,
    ["发生原因", "不良原因", "原因"],
    /发生原因|不良原因|原因|cause|reason/i,
    "发生原因（可选）",
    true
  );

  const status = pickColumn(
    allColumns,
    ["不良状态", "状态"],
    /不良状态|状态|status/i,
    "状态（可选）",
    true
  );

  const date = pickColumn(
    allColumns,
    ["录入时间", "日期", "检测日期", "时间"],
    /录入时间|日期|检测日期|时间|date/i,
    "日期（可选）",
    true
  );

  return {
    ...required,
    reason: reason || required.defect,
    status,
    date,
  };
}

function pickColumn(allColumns, exactCandidates, fuzzyPattern, fieldName, optional = false) {
  const exactMatch = exactCandidates.find((name) => allColumns.includes(name));
  if (exactMatch) {
    return exactMatch;
  }

  const fuzzyMatch = allColumns.find((name) => fuzzyPattern.test(String(name)));
  if (fuzzyMatch) {
    return fuzzyMatch;
  }

  if (optional) {
    return "";
  }

  throw new Error(`缺少必要列：${fieldName}。`);
}

function collectColumns(rows) {
  const set = new Set();
  rows.forEach((row) => {
    Object.keys(row || {}).forEach((key) => set.add(key));
  });
  return Array.from(set);
}

function buildNormalizedRows(rows, columns) {
  return rows.map((row) => {
    const inspectionName = normalizeCell(row[columns.inspection]);
    const defectName = normalizeCell(row[columns.defect]);
    const mergedName = [inspectionName, defectName].filter(Boolean).join(" - ");

    const progress = normalizeProgress(normalizeCell(row[columns.progress]) || EMPTY_VALUE);
    const rawStatus = columns.status ? normalizeCell(row[columns.status]) : "";
    const normalizedStatus = normalizeStatus(rawStatus, progress);

    return {
      merged: mergedName || EMPTY_VALUE,
      model: normalizeCell(row[columns.model]) || EMPTY_VALUE,
      supplier: normalizeCell(row[columns.supplier]) || EMPTY_VALUE,
      level: normalizeLevel(normalizeCell(row[columns.level])),
      progress,
      reason: normalizeCell(row[columns.reason]) || EMPTY_VALUE,
      statusText: normalizedStatus.text,
      statusClass: normalizedStatus.className,
      dateObj: columns.date ? parseDate(row[columns.date]) : null,
    };
  });
}

function normalizeProgress(progressText) {
  const source = String(progressText || "").trim();
  if (!source || source === EMPTY_VALUE) {
    return "解析中";
  }
  if (/已关闭|关闭|完结|完成|closed|done/i.test(source) && !/未关闭/.test(source)) {
    return "已关闭";
  }
  if (/对策中|整改中|改善中|措施中|countermeasure|action/i.test(source)) {
    return "对策中";
  }
  return "解析中";
}

function normalizeLevel(levelText) {
  const source = String(levelText || "").trim().toUpperCase();
  if (source.includes("A")) return "A";
  if (source.includes("B")) return "B";
  if (source.includes("C")) return "C";
  if (source.includes("D")) return "D";
  return "D";
}

function normalizeStatus(statusText, progressText) {
  const source = `${statusText} ${progressText}`;

  if (/已关闭|关闭|完成/.test(source) && !/未关闭/.test(source)) {
    return { text: "已关闭", className: "closed" };
  }

  if (/未关闭|处理中|进行中|待处理|打开/.test(source)) {
    return { text: "未关闭", className: "open" };
  }

  return { text: "未关闭", className: "open" };
}

function parseDate(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number") {
    if (value > 1e11) {
      const msDate = new Date(value);
      return Number.isNaN(msDate.getTime()) ? null : msDate;
    }
    if (value > 1e9) {
      const secDate = new Date(value * 1000);
      return Number.isNaN(secDate.getTime()) ? null : secDate;
    }
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) {
      return null;
    }
    return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, parsed.S || 0);
  }

  const text = String(value).trim();
  if (!text) {
    return null;
  }

  if (/^\d+$/.test(text)) {
    const numeric = Number(text);
    if (!Number.isNaN(numeric)) {
      return parseDate(numeric);
    }
  }

  const normalized = text
    .replace(/[年./]/g, "-")
    .replace(/月/g, "-")
    .replace(/[日号]/g, " ")
    .replace(/T/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const matched = normalized.match(
    /(\d{4})\D+(\d{1,2})\D+(\d{1,2})(?:\D+(\d{1,2})(?:\D+(\d{1,2}))?(?:\D+(\d{1,2}))?)?/
  );
  if (matched) {
    const [, year, month, day, hour = "0", minute = "0", second = "0"] = matched;
    const manualDate = new Date(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour),
      Number(minute),
      Number(second)
    );
    if (!Number.isNaN(manualDate.getTime())) {
      return manualDate;
    }
  }

  const date = new Date(normalized);
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  return date;
}

function buildStats(rows) {
  const grouped = new Map();

  rows.forEach((row) => {
    if (!grouped.has(row.merged)) {
      grouped.set(row.merged, {
        count: 0,
        levelCounter: new Map(),
        supplierCounter: new Map(),
      });
    }

    const item = grouped.get(row.merged);
    item.count += 1;
    item.levelCounter.set(row.level, (item.levelCounter.get(row.level) || 0) + 1);
    item.supplierCounter.set(row.supplier, (item.supplierCounter.get(row.supplier) || 0) + 1);
  });

  return Array.from(grouped.entries())
    .map(([mergedName, item]) => ({
      [MERGED_COLUMN]: mergedName,
      数量: item.count,
      等级: pickTopLabel(item.levelCounter),
      供应商: pickTopLabel(item.supplierCounter),
    }))
    .sort((a, b) => {
      if (b.数量 !== a.数量) {
        return b.数量 - a.数量;
      }
      return a[MERGED_COLUMN].localeCompare(b[MERGED_COLUMN], "zh-CN");
    })
    .map((row, index) => ({
      序号: index + 1,
      [MERGED_COLUMN]: row[MERGED_COLUMN],
      数量: row.数量,
      等级: row.等级,
      供应商: row.供应商,
    }));
}

function buildModelProjectData(rows) {
  const projectSetByModel = new Map();

  rows.forEach((row) => {
    if (!projectSetByModel.has(row.model)) {
      projectSetByModel.set(row.model, new Set());
    }
    projectSetByModel.get(row.model).add(row.merged);
  });

  const entries = Array.from(projectSetByModel.entries())
    .map(([model, set]) => ({ model, count: set.size }))
    .sort((a, b) => b.count - a.count);

  if (entries.length > MAX_PIE_SEGMENTS) {
    const kept = entries.slice(0, MAX_PIE_SEGMENTS - 1);
    const others = entries.slice(MAX_PIE_SEGMENTS - 1);
    const otherCount = others.reduce((sum, item) => sum + item.count, 0);
    entries.length = 0;
    entries.push(...kept, { model: `其他(${others.length})`, count: otherCount });
  }

  return {
    labels: entries.map((item) => item.model),
    values: entries.map((item) => item.count),
  };
}

function buildReasonPieData(rows) {
  return buildCountPieData(rows.map((row) => row.reason));
}

function buildProgressPieData(rows) {
  return buildCountPieData(rows.map((row) => row.progress));
}

function buildCountPieData(values) {
  const counter = new Map();
  values.forEach((value) => {
    const key = value || EMPTY_VALUE;
    counter.set(key, (counter.get(key) || 0) + 1);
  });

  const entries = Array.from(counter.entries())
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count);

  if (entries.length > MAX_PIE_SEGMENTS) {
    const kept = entries.slice(0, MAX_PIE_SEGMENTS - 1);
    const others = entries.slice(MAX_PIE_SEGMENTS - 1);
    const otherCount = others.reduce((sum, item) => sum + item.count, 0);
    entries.length = 0;
    entries.push(...kept, { label: `其他(${others.length})`, count: otherCount });
  }

  return {
    labels: entries.map((item) => item.label),
    values: entries.map((item) => item.count),
  };
}

function buildTrendData(rows, selectedRange) {
  const dateRows = rows.filter((row) => row.dateObj instanceof Date && !Number.isNaN(row.dateObj.getTime()));
  if (!dateRows.length) {
    return { labels: [], values: [], granularity: "day", buckets: [] };
  }

  const bounds = getTrendBounds(dateRows, selectedRange);
  const granularity = detectTrendGranularity(selectedRange, bounds.start, bounds.end);
  const buckets = buildTrendBuckets(bounds.start, bounds.end, granularity);
  const counter = new Map(buckets.map((bucket) => [bucket.key, 0]));

  dateRows.forEach((row) => {
    const key = resolveTrendBucketKey(row.dateObj, bounds.start, bounds.end, granularity);
    if (!counter.has(key)) {
      return;
    }
    counter.set(key, counter.get(key) + 1);
  });

  return {
    labels: buckets.map((bucket) => bucket.label),
    values: buckets.map((bucket) => counter.get(bucket.key) || 0),
    granularity,
    buckets,
  };
}

function getTrendBounds(dateRows, selectedRange) {
  const base = filterBaseDate ? new Date(filterBaseDate) : new Date(Math.max(...dateRows.map((row) => row.dateObj.getTime())));

  if (selectedRange === "this_week") {
    const start = startOfWeek(base);
    return { start, end: addDays(start, 7) };
  }

  if (selectedRange === "this_month") {
    const start = new Date(base.getFullYear(), base.getMonth(), 1);
    return { start, end: new Date(base.getFullYear(), base.getMonth() + 1, 1) };
  }

  if (selectedRange === "this_quarter") {
    const quarterStartMonth = Math.floor(base.getMonth() / 3) * 3;
    const start = new Date(base.getFullYear(), quarterStartMonth, 1);
    return { start, end: new Date(base.getFullYear(), quarterStartMonth + 3, 1) };
  }

  if (selectedRange === "this_year") {
    const start = new Date(base.getFullYear(), 0, 1);
    return { start, end: new Date(base.getFullYear() + 1, 0, 1) };
  }

  const minDate = toDayStart(new Date(Math.min(...dateRows.map((row) => row.dateObj.getTime()))));
  const maxDate = toDayStart(new Date(Math.max(...dateRows.map((row) => row.dateObj.getTime()))));
  return { start: minDate, end: addDays(maxDate, 1) };
}

function detectTrendGranularity(selectedRange, start, end) {
  if (selectedRange === "this_week") return "day";
  if (selectedRange === "this_month") return "week";
  if (selectedRange === "this_quarter") return "month";
  if (selectedRange === "this_year") return "month";

  const daySpan = Math.max(1, Math.ceil((end.getTime() - start.getTime()) / 86400000));
  if (daySpan <= 14) return "day";
  if (daySpan <= 120) return "week";
  return "month";
}

function buildTrendBuckets(start, end, granularity) {
  if (granularity === "day") {
    const buckets = [];
    for (let cursor = new Date(start); cursor < end; cursor = addDays(cursor, 1)) {
      const day = toDayStart(cursor);
      buckets.push({
        key: `d-${day.getTime()}`,
        label: formatDayLabel(day),
        start: day,
        end: addDays(day, 1),
      });
    }
    return buckets;
  }

  if (granularity === "week") {
    const buckets = [];
    for (let weekStart = startOfWeek(start); weekStart < end; weekStart = addDays(weekStart, 7)) {
      const bucketStart = maxDate(weekStart, start);
      const bucketEnd = minDate(addDays(weekStart, 7), end);
      buckets.push({
        key: `w-${bucketStart.getTime()}-${bucketEnd.getTime()}`,
        label: `${formatMonthDayLabel(bucketStart)}~${formatMonthDayLabel(addDays(bucketEnd, -1))}`,
        start: bucketStart,
        end: bucketEnd,
      });
    }
    return buckets;
  }

  const buckets = [];
  for (
    let monthStart = new Date(start.getFullYear(), start.getMonth(), 1);
    monthStart < end;
    monthStart = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 1)
  ) {
    const bucketStart = maxDate(monthStart, start);
    const bucketEnd = minDate(new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 1), end);
    buckets.push({
      key: `m-${bucketStart.getFullYear()}-${bucketStart.getMonth() + 1}`,
      label: `${bucketStart.getFullYear()}-${String(bucketStart.getMonth() + 1).padStart(2, "0")}`,
      start: bucketStart,
      end: bucketEnd,
    });
  }
  return buckets;
}

function resolveTrendBucketKey(dateObj, start, end, granularity) {
  const date = toDayStart(dateObj);
  if (granularity === "day") {
    return `d-${date.getTime()}`;
  }

  if (granularity === "week") {
    const weekStart = startOfWeek(date);
    const bucketStart = maxDate(weekStart, start);
    const bucketEnd = minDate(addDays(weekStart, 7), end);
    return `w-${bucketStart.getTime()}-${bucketEnd.getTime()}`;
  }

  return `m-${date.getFullYear()}-${date.getMonth() + 1}`;
}

function toDayStart(date) {
  const cloned = new Date(date);
  cloned.setHours(0, 0, 0, 0);
  return cloned;
}

function addDays(date, days) {
  const cloned = new Date(date);
  cloned.setDate(cloned.getDate() + days);
  return cloned;
}

function startOfWeek(date) {
  const dayStart = toDayStart(date);
  const mondayOffset = (dayStart.getDay() + 6) % 7;
  return addDays(dayStart, -mondayOffset);
}

function minDate(a, b) {
  return a.getTime() <= b.getTime() ? new Date(a) : new Date(b);
}

function maxDate(a, b) {
  return a.getTime() >= b.getTime() ? new Date(a) : new Date(b);
}

function formatDayLabel(date) {
  return `${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatMonthDayLabel(date) {
  return `${String(date.getMonth() + 1).padStart(2, "0")}/${String(date.getDate()).padStart(2, "0")}`;
}

function buildTopSupplierDetails(rows) {
  const supplierStats = new Map();

  rows.forEach((row) => {
    if (!supplierStats.has(row.supplier)) {
      supplierStats.set(row.supplier, {
        projectSet: new Set(),
        levelCounter: new Map(),
        mergedCounter: new Map(),
      });
    }

    const stat = supplierStats.get(row.supplier);
    stat.projectSet.add(row.merged);
    stat.levelCounter.set(row.level, (stat.levelCounter.get(row.level) || 0) + 1);
    stat.mergedCounter.set(row.merged, (stat.mergedCounter.get(row.merged) || 0) + 1);
  });

  return Array.from(supplierStats.entries())
    .map(([supplier, stat]) => {
      const topDefect = pickTopLabelsAndCount(stat.mergedCounter);
      return {
        供应商: supplier,
        等级: pickTopLabel(stat.levelCounter),
        项目数: stat.projectSet.size,
        最多不良名称: topDefect.labels,
        不良数量: topDefect.count,
      };
    })
    .sort((a, b) => {
      if (b.项目数 !== a.项目数) {
        return b.项目数 - a.项目数;
      }
      if (b.不良数量 !== a.不良数量) {
        return b.不良数量 - a.不良数量;
      }
      return a.供应商.localeCompare(b.供应商, "zh-CN");
    })
    .slice(0, TOP_SUPPLIER_COUNT);
}

function buildTopDefectDetails(rows) {
  const defectStats = new Map();

  rows.forEach((row) => {
    if (!defectStats.has(row.merged)) {
      defectStats.set(row.merged, {
        supplierSet: new Set(),
        levelCounter: new Map(),
        supplierCounter: new Map(),
      });
    }

    const stat = defectStats.get(row.merged);
    stat.supplierSet.add(row.supplier);
    stat.levelCounter.set(row.level, (stat.levelCounter.get(row.level) || 0) + 1);
    stat.supplierCounter.set(row.supplier, (stat.supplierCounter.get(row.supplier) || 0) + 1);
  });

  return Array.from(defectStats.entries())
    .map(([defectName, stat]) => {
      const topSupplier = pickTopLabelsAndCount(stat.supplierCounter);
      return {
        不良名称: defectName,
        等级: pickTopLabel(stat.levelCounter),
        项目数: stat.supplierSet.size,
        最多供应商: topSupplier.labels,
        不良数量: topSupplier.count,
      };
    })
    .sort((a, b) => {
      if (b.项目数 !== a.项目数) {
        return b.项目数 - a.项目数;
      }
      if (b.不良数量 !== a.不良数量) {
        return b.不良数量 - a.不良数量;
      }
      return a.不良名称.localeCompare(b.不良名称, "zh-CN");
    })
    .slice(0, TOP_SUPPLIER_COUNT);
}

function pickTopLabel(counterMap) {
  if (!counterMap || counterMap.size === 0) {
    return EMPTY_VALUE;
  }

  let maxValue = 0;
  counterMap.forEach((value) => {
    if (value > maxValue) {
      maxValue = value;
    }
  });

  const candidates = Array.from(counterMap.entries())
    .filter(([, value]) => value === maxValue)
    .map(([label]) => label)
    .sort((a, b) => String(a).localeCompare(String(b), "zh-CN"));

  return candidates[0] || EMPTY_VALUE;
}

function pickTopLabels(counterMap, limit = 3) {
  if (!counterMap || counterMap.size === 0) {
    return EMPTY_VALUE;
  }

  let maxValue = 0;
  counterMap.forEach((value) => {
    if (value > maxValue) {
      maxValue = value;
    }
  });

  const candidates = Array.from(counterMap.entries())
    .filter(([, value]) => value === maxValue)
    .map(([label]) => label)
    .sort((a, b) => String(a).localeCompare(String(b), "zh-CN"));

  const shown = candidates.slice(0, limit).join(" / ");
  if (candidates.length <= limit) {
    return shown || EMPTY_VALUE;
  }
  return `${shown} 等${candidates.length}项`;
}

function pickTopLabelsAndCount(counterMap, limit = 3) {
  if (!counterMap || counterMap.size === 0) {
    return { labels: EMPTY_VALUE, count: 0 };
  }

  let maxValue = 0;
  counterMap.forEach((value) => {
    if (value > maxValue) {
      maxValue = value;
    }
  });

  return {
    labels: pickTopLabels(counterMap, limit),
    count: maxValue,
  };
}

function buildReasonParetoData(rows) {
  if (!rows.length) {
    return { labels: [], stackDatasets: [], cumulativePercent: [] };
  }

  const reasonTotals = new Map();
  const progressTotals = new Map();

  rows.forEach((row) => {
    reasonTotals.set(row.reason, (reasonTotals.get(row.reason) || 0) + 1);
    progressTotals.set(row.progress, (progressTotals.get(row.progress) || 0) + 1);
  });

  const topReasons = Array.from(reasonTotals.entries())
    .map(([reason, count]) => ({ reason, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, TOP_REASON_COUNT);

  const reasonLabels = topReasons.map((item) => item.reason);
  const progressLabels = Array.from(progressTotals.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([progress]) => progress);

  const reasonLabelSet = new Set(reasonLabels);
  const stackCounter = new Map();

  progressLabels.forEach((progress) => {
    stackCounter.set(progress, new Map(reasonLabels.map((reason) => [reason, 0])));
  });

  rows.forEach((row) => {
    if (!reasonLabelSet.has(row.reason)) {
      return;
    }
    const reasonCounter = stackCounter.get(row.progress);
    reasonCounter.set(row.reason, (reasonCounter.get(row.reason) || 0) + 1);
  });

  const reasonCounts = topReasons.map((item) => item.count);
  const total = reasonCounts.reduce((sum, count) => sum + count, 0) || 1;

  let running = 0;
  const cumulativePercent = reasonCounts.map((count, index) => {
    running += count;
    if (index === reasonCounts.length - 1) {
      return 100;
    }
    return Number(((running / total) * 100).toFixed(2));
  });

  const stackDatasets = progressLabels.map((progress, index) => {
    const reasonCounter = stackCounter.get(progress);
    return {
      type: "bar",
      label: progress,
      data: reasonLabels.map((reason) => reasonCounter.get(reason) || 0),
      backgroundColor: colorForIndex(index, 0.82),
      borderColor: colorForIndex(index, 1),
      borderWidth: 1,
      yAxisID: "yCount",
      stack: "progress",
    };
  });

  return {
    labels: reasonLabels,
    stackDatasets,
    cumulativePercent,
  };
}

function renderModelProjectPie(data) {
  if (!modelProjectPieCanvas) {
    return;
  }
  modelProjectPieChart = renderDoughnutPie(modelProjectPieCanvas, modelProjectPieChart, data);
}

function renderReasonPie(data) {
  if (!reasonPieCanvas) {
    return;
  }
  reasonPieChart = renderDoughnutPie(reasonPieCanvas, reasonPieChart, data);
}

function renderProgressPie(data) {
  if (!progressPieCanvas) {
    return;
  }
  progressPieChart = renderDoughnutPie(progressPieCanvas, progressPieChart, data);
}

function renderDoughnutPie(canvas, chartRef, data) {
  if (chartRef) {
    chartRef.destroy();
  }

  const total = data.values.reduce((sum, value) => sum + value, 0) || 1;

  return new Chart(canvas, {
    type: "doughnut",
    data: {
      labels: data.labels,
      datasets: [
        {
          data: data.values,
          backgroundColor: data.values.map((_, index) => colorForIndex(index, 0.82)),
          borderColor: data.values.map((_, index) => colorForIndex(index, 1)),
          borderWidth: 1,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "40%",
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            color: "#eef7ff",
            boxWidth: 14,
            boxHeight: 14,
            padding: 10,
          },
        },
        tooltip: {
          backgroundColor: "rgba(7, 28, 66, 0.95)",
          titleColor: "#ffffff",
          bodyColor: "#dff2ff",
          borderColor: "rgba(95, 187, 255, 0.65)",
          borderWidth: 1,
        },
        datalabels: {
          display: true,
          color: "#f7fbff",
          anchor: "center",
          align: "center",
          offset: 0,
          clamp: true,
          clip: false,
          font: {
            size: 9,
            weight: "700",
          },
          formatter: (value) => {
            const percent = ((value / total) * 100).toFixed(1);
            return `${percent}%`;
          },
        },
      },
    },
    plugins: window.ChartDataLabels ? [ChartDataLabels] : [],
  });
}

function renderDefectTrend(data, rows) {
  if (!defectTrendChartCanvas) {
    return;
  }
  if (defectTrendChart) {
    defectTrendChart.destroy();
  }

  const xTitleMap = {
    day: "日期（日）",
    week: "日期（周）",
    month: "日期（月）",
  };

  defectTrendChart = new Chart(defectTrendChartCanvas, {
    type: "line",
    data: {
      labels: data.labels,
      datasets: [
        {
          label: "不良数量",
          data: data.values,
          fill: true,
          tension: 0.28,
          borderColor: "rgba(76, 194, 255, 1)",
          backgroundColor: "rgba(76, 194, 255, 0.18)",
          pointRadius: 3,
          pointHoverRadius: 5,
          pointBackgroundColor: "rgba(123, 224, 200, 1)",
          pointBorderColor: "rgba(7, 20, 38, 1)",
          pointBorderWidth: 1.5,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false,
        },
        tooltip: {
          backgroundColor: "rgba(7, 28, 66, 0.95)",
          titleColor: "#ffffff",
          bodyColor: "#dff2ff",
          borderColor: "rgba(95, 187, 255, 0.65)",
          borderWidth: 1,
        },
        datalabels: {
          display: true,
          align: "top",
          anchor: "end",
          color: "#9fe7d2",
          offset: 4,
          font: {
            size: 10,
            weight: "700",
          },
          formatter: (value) => String(Number(value || 0)),
        },
      },
      onClick: (_, elements) => {
        if (!elements.length) {
          return;
        }
        const { index } = elements[0];
        const bucket = data.buckets[index];
        if (!bucket) {
          return;
        }

        const detailRows = buildIssueDetailRowsByDateRange(rows, bucket.start, bucket.end);
        issueDetailTitle.textContent = `问题点明细：${bucket.label}`;
        renderTable(issueDetailTable, withTotalRow(detailRows, "不良原因", "数量"));
        issueDetailTable.closest(".panel")?.scrollIntoView({ behavior: "smooth", block: "start" });
      },
      onHover: (event, elements) => {
        event.native.target.style.cursor = elements.length ? "pointer" : "default";
      },
      scales: {
        x: {
          title: {
            display: true,
            text: xTitleMap[data.granularity] || "日期",
            color: "#d8ecff",
          },
          ticks: {
            color: "#e6f4ff",
            maxRotation: 0,
            autoSkip: true,
            maxTicksLimit: 12,
          },
          grid: {
            color: "rgba(98, 160, 235, 0.12)",
          },
        },
        y: {
          beginAtZero: true,
          title: {
            display: true,
            text: "不良数量",
            color: "#d8ecff",
          },
          ticks: {
            precision: 0,
            color: "#e6f4ff",
          },
          grid: {
            color: "rgba(98, 160, 235, 0.16)",
          },
        },
      },
    },
    plugins: window.ChartDataLabels ? [ChartDataLabels] : [],
  });
}

function renderReasonPareto(data, rows) {
  if (!progressStackChartCanvas) {
    return;
  }
  if (progressStackChart) {
    progressStackChart.destroy();
  }

  progressStackChart = new Chart(progressStackChartCanvas, {
    type: "bar",
    data: {
      labels: data.labels,
      datasets: [
        ...data.stackDatasets,
        {
          type: "line",
          label: "累计占比(%)",
          data: data.cumulativePercent,
          borderColor: "#ff9f4a",
          backgroundColor: "#ff9f4a",
          borderWidth: 2.5,
          tension: 0.2,
          pointRadius: 3.5,
          pointHoverRadius: 5,
          yAxisID: "yPercent",
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: {
          right: 20,
          top: 12,
        },
      },
      onClick: (_, elements) => {
        if (!elements.length) {
          return;
        }

        const { index, datasetIndex } = elements[0];
        const reason = data.labels[index];
        const selectedDataset = progressStackChart.data.datasets[datasetIndex];
        const selectedProgress = selectedDataset.type === "line" ? "" : selectedDataset.label;
        const detailRows = buildIssueDetailRowsByReason(rows, reason, selectedProgress);

        issueDetailTitle.textContent = selectedProgress
          ? `问题点明细：${reason} / ${selectedProgress}`
          : `问题点明细：${reason}`;
        renderTable(issueDetailTable, withTotalRow(detailRows, "不良原因", "数量"));
        issueDetailTable.closest(".panel")?.scrollIntoView({ behavior: "smooth", block: "start" });
      },
      onHover: (event, elements) => {
        event.native.target.style.cursor = elements.length ? "pointer" : "default";
      },
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            color: "#ecf6ff",
            boxWidth: 14,
            boxHeight: 14,
            padding: 12,
          },
        },
        tooltip: {
          backgroundColor: "rgba(7, 28, 66, 0.95)",
          titleColor: "#ffffff",
          bodyColor: "#dff2ff",
          borderColor: "rgba(95, 187, 255, 0.65)",
          borderWidth: 1,
        },
        datalabels: {
          display: (context) => context.dataset.type === "line",
          color: "#ffb779",
          align: (context) => {
            const idx = context.dataIndex;
            const last = context.dataset.data.length - 1;
            if (idx === 0) return "right";
            if (idx === last) return "left";
            return idx % 2 === 0 ? "top" : "bottom";
          },
          anchor: "end",
          offset: (context) => {
            const idx = context.dataIndex;
            if (idx === 0 || idx === context.dataset.data.length - 1) {
              return 10;
            }
            return idx % 2 === 0 ? 10 : 8;
          },
          clip: false,
          backgroundColor: "rgba(8, 26, 58, 0.78)",
          borderColor: "rgba(255, 183, 121, 0.38)",
          borderWidth: 1,
          borderRadius: 4,
          padding: {
            top: 2,
            right: 4,
            bottom: 2,
            left: 4,
          },
          formatter: (value) => `${Number(value).toFixed(1)}%`,
          font: {
            size: 10,
            weight: "700",
          },
        },
      },
      scales: {
        x: {
          stacked: true,
          title: {
            display: true,
            text: "不良原因",
            color: "#d8ecff",
          },
          ticks: {
            maxRotation: 25,
            minRotation: 0,
            color: "#e6f4ff",
          },
          grid: {
            color: "rgba(98, 160, 235, 0.12)",
          },
        },
        yCount: {
          type: "linear",
          position: "left",
          beginAtZero: true,
          stacked: true,
          ticks: {
            precision: 0,
            color: "#e6f4ff",
          },
          title: {
            display: true,
            text: "数量",
            color: "#d8ecff",
          },
          grid: {
            color: "rgba(98, 160, 235, 0.16)",
          },
        },
        yPercent: {
          type: "linear",
          position: "right",
          beginAtZero: true,
          min: 0,
          max: 110,
          grid: {
            drawOnChartArea: false,
          },
          ticks: {
            callback: (value) => `${value}%`,
            color: "#ffd6a7",
          },
          title: {
            display: true,
            text: "累计占比",
            color: "#ffd6a7",
          },
        },
      },
    },
    plugins: window.ChartDataLabels ? [ChartDataLabels] : [],
  });
}

function buildIssueDetailRowsByReason(rows, reason, progress = "") {
  const counter = new Map();

  rows.forEach((row) => {
    if (row.reason !== reason) {
      return;
    }
    if (progress && row.progress !== progress) {
      return;
    }

    const key = [row.reason, row.progress, row.merged, row.supplier, row.level].join("|");
    counter.set(key, (counter.get(key) || 0) + 1);
  });

  return Array.from(counter.entries())
    .map(([key, count]) => {
      const [reasonName, progressName, merged, supplier, level] = key.split("|");
      return {
        不良原因: reasonName,
        解决进度: progressName,
        合并项: merged,
        供应商: supplier,
        等级: level,
        数量: count,
      };
    })
    .sort((a, b) => b.数量 - a.数量);
}

function buildIssueDetailRowsByDateRange(rows, start, end) {
  const counter = new Map();

  rows.forEach((row) => {
    if (!(row.dateObj instanceof Date) || Number.isNaN(row.dateObj.getTime())) {
      return;
    }
    if (row.dateObj < start || row.dateObj >= end) {
      return;
    }

    const dayText = formatDateText(row.dateObj);
    const key = [dayText, row.reason, row.progress, row.merged, row.supplier, row.level].join("|");
    counter.set(key, (counter.get(key) || 0) + 1);
  });

  return Array.from(counter.entries())
    .map(([key, count]) => {
      const [dayText, reasonName, progressName, merged, supplier, level] = key.split("|");
      return {
        日期: dayText,
        不良原因: reasonName,
        解决进度: progressName,
        合并项: merged,
        供应商: supplier,
        等级: level,
        数量: count,
      };
    })
    .sort((a, b) => {
      if (b.数量 !== a.数量) {
        return b.数量 - a.数量;
      }
      return b.日期.localeCompare(a.日期, "zh-CN");
    });
}

function formatDateText(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function withTotalRow(rows, labelColumn, metricColumn) {
  if (!rows.length || !(metricColumn in rows[0])) {
    return rows;
  }

  const total = rows.reduce((sum, row) => sum + Number(row[metricColumn] || 0), 0);
  const totalRow = {};

  Object.keys(rows[0]).forEach((column) => {
    if (column === labelColumn) {
      totalRow[column] = "合计";
      return;
    }
    if (column === metricColumn) {
      totalRow[column] = total;
      return;
    }
    totalRow[column] = "-";
  });

  return [...rows, totalRow];
}

function renderTable(table, rows) {
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!rows.length) {
    const emptyHeadRow = document.createElement("tr");
    const emptyHeadCell = document.createElement("th");
    emptyHeadCell.textContent = "暂无数据";
    emptyHeadRow.appendChild(emptyHeadCell);
    thead.appendChild(emptyHeadRow);
    return;
  }

  const columns = Object.keys(rows[0]);
  const headRow = document.createElement("tr");

  columns.forEach((column) => {
    const th = document.createElement("th");
    th.textContent = column;
    headRow.appendChild(th);
  });

  thead.appendChild(headRow);

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    columns.forEach((column) => {
      const td = document.createElement("td");
      td.textContent = normalizeCell(row[column]);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

function exportWorkbook(rows, fileName, sheetName) {
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  XLSX.writeFile(workbook, fileName);
}

function colorForIndex(index, alpha = 1) {
  const palette = [
    [184, 92, 56],
    [63, 125, 88],
    [15, 118, 110],
    [217, 119, 6],
    [59, 130, 246],
    [220, 38, 38],
    [99, 102, 241],
    [8, 145, 178],
    [20, 184, 166],
    [139, 92, 246],
  ];
  const [r, g, b] = palette[index % palette.length];
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function normalizeCell(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}
