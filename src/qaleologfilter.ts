import * as fs from "fs";
import * as path from "path";
import * as XLSX from "xlsx";
import {
  cleanLogText,
  extractDate,
  extractErrorMessage,
  extractMethodAndEndpoint,
  resolveDetectLevel,
} from "./logParser";
import {
  loadWorkbookFromOneDrive,
  syncWorkbookToOneDrive,
} from "./onedriveSync";

const ROOT_DIR = process.cwd();
const INPUT_DIR = path.join(ROOT_DIR, "src", "qa", "leo");
const OUTPUT_DIR = path.join(ROOT_DIR, "src", "qa", "leo", "output");

const SKIP_LEVELS = new Set(["info"]);

function parseFileDate(filename: string): string | null {
  const iso = filename.match(/^(\d{4})-(\d{2})-(\d{2})\.json$/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;

  const dmy = filename.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\.json$/);
  if (dmy) {
    const day = dmy[1].padStart(2, "0");
    const month = dmy[2].padStart(2, "0");
    return `${dmy[3]}-${month}-${day}`;
  }

  return null;
}

const files = fs
  .readdirSync(INPUT_DIR)
  .map((f) => ({ name: f, date: parseFileDate(f) }))
  .filter((f) => f.date !== null)
  .sort((a, b) => a.date!.localeCompare(b.date!));

if (files.length === 0) {
  console.error("No input files found in qa/leo");
  process.exit(1);
}

const latest = files[files.length - 1];
const latestFile = latest.name;
const dateStr = latest.date!;
console.log(`Latest file selected: ${latestFile}`);

const filePath = path.join(INPUT_DIR, latestFile);
let data: {
  series?: { fields?: { name?: string; values?: Record<string, string>[] }[] }[];
};

const raw = fs.readFileSync(filePath, "utf-8");
if (!raw.trim()) {
  console.error(`Input file is empty: ${latestFile}`);
  process.exit(1);
}

try {
  data = JSON.parse(raw);
} catch {
  console.error(`Failed to parse JSON: ${latestFile}`);
  process.exit(1);
}

const HEADER_ROW: (string | undefined)[] = [
  "Service Name",
  "Date",
  "Method",
  "Endpoint",
  "Count",
  "Error Message",
  "Detect-Level",
];

const WORKBOOK_PATH = path.join(ROOT_DIR, "output", "GrafanaErrorLog.xlsx");
const SHEET_NAME = "QA";
const ALL_SHEETS = ["QA", "UAT", "CLIENT_UAT", "PROD"];
const ENDPOINT_COL = 3;
const COUNT_COL = 4;

const outputDir = path.join(ROOT_DIR, "output");

if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

loadWorkbookFromOneDrive(WORKBOOK_PATH);

let workbook: XLSX.WorkBook;
if (fs.existsSync(WORKBOOK_PATH)) {
  workbook = XLSX.readFile(WORKBOOK_PATH);
} else {
  workbook = XLSX.utils.book_new();
  for (const name of ALL_SHEETS) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([HEADER_ROW]), name);
  }
}

let rows: (string | undefined)[][];
if (workbook.Sheets[SHEET_NAME]) {
  const existingRows = XLSX.utils.sheet_to_json(workbook.Sheets[SHEET_NAME], {
    header: 1,
    defval: "",
  }) as (string | undefined)[][];
  if (existingRows.length > 0) {
    if (!existingRows[0].includes("Count")) {
      rows = existingRows.map((row, i) => {
        if (i === 0) {
          const h = [...row];
          h.splice(COUNT_COL, 0, "Count");
          return h;
        }
        const r = [...row];
        r.splice(COUNT_COL, 0, "1");
        return r;
      });
    } else {
      rows = existingRows;
    }
  } else {
    rows = [HEADER_ROW];
  }
} else {
  rows = [HEADER_ROW];
}

const seenEndpoints = new Map<string, number>();
for (let i = 1; i < rows.length; i++) {
  const ep = rows[i][ENDPOINT_COL] ?? "";
  if (ep) seenEndpoints.set(ep, i);
}

let totalLogs = 0;
let filteredLogs = 0;

for (const series of data.series ?? []) {
  for (const field of series.fields ?? []) {
    if (field.name !== "labels") continue;

    for (const value of field.values ?? []) {
      totalLogs++;

      const detectLevel = value.detected_level ?? "";
      if (SKIP_LEVELS.has(detectLevel)) continue;

      const log = value.log ?? "";
      const serviceName = value.service_name ?? "";
      const cleanLog = cleanLogText(log);
      const date = extractDate(cleanLog);
      const { method, endpoint } = extractMethodAndEndpoint(cleanLog);
      const errorMessage = extractErrorMessage(cleanLog);
      const level = resolveDetectLevel(detectLevel, cleanLog);
      if (level !== "error") continue;

      if (!endpoint) continue;

      if (endpoint && seenEndpoints.has(endpoint)) {
        const idx = seenEndpoints.get(endpoint)!;
        const current = Number(rows[idx][COUNT_COL] ?? 1);
        rows[idx][COUNT_COL] = String(current + 1);
        filteredLogs++;
        continue;
      }
      if (endpoint) seenEndpoints.set(endpoint, rows.length);

      filteredLogs++;
      rows.push([
        serviceName,
        date,
        method,
        endpoint,
        "1",
        errorMessage,
        level,
      ]);
    }
  }
}

console.log(`Total logs: ${totalLogs}`);
console.log(`Filtered logs: ${filteredLogs}`);

const combinedRows = rows;

const worksheet = XLSX.utils.aoa_to_sheet(combinedRows);

const padding = 3;
const minWidth = 15;
const maxWidth = 60;

const autoFitCols = (sheetRows: (string | undefined)[][]) => {
  const colCount = sheetRows[0]?.length ?? 0;
  return Array.from({ length: colCount }, (_, col) => {
    let maxLen = 0;
    for (const row of sheetRows) {
      const len = String(row[col] ?? "").length;
      if (len > maxLen) maxLen = len;
    }
    const width = Math.min(Math.max(maxLen + padding, minWidth), maxWidth);
    return { wch: width };
  });
};

worksheet["!cols"] = autoFitCols(combinedRows);

if (workbook.SheetNames.includes(SHEET_NAME)) {
  workbook.Sheets[SHEET_NAME] = worksheet;
} else {
  XLSX.utils.book_append_sheet(workbook, worksheet, SHEET_NAME);
}

for (const name of workbook.SheetNames) {
  if (name === SHEET_NAME) continue;
  const sheetRows = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
    header: 1,
    defval: "",
  }) as (string | undefined)[][];
  if (sheetRows.length > 0) {
    workbook.Sheets[name]["!cols"] = autoFitCols(sheetRows);
  }
}

XLSX.writeFile(workbook, WORKBOOK_PATH);

syncWorkbookToOneDrive(WORKBOOK_PATH);

console.log(`Excel generated successfully: ${WORKBOOK_PATH} (sheet: ${SHEET_NAME})`);
