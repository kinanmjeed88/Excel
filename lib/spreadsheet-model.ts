import { displayCellValue, type CellValues } from "./formula-engine";

export type CellNumberFormat = "general" | "currency" | "percent" | "decimal";

export type CellFormat = {
  numberFormat?: CellNumberFormat;
  backgroundColor?: string;
  textColor?: string;
};

export type CellFormats = Record<string, CellFormat>;
export type CellRange = { start: string; end: string };
export type ColumnWidths = Record<string, number>;

export type SpreadsheetSheet = {
  id: string;
  name: string;
  cells: CellValues;
  cellFormats: CellFormats;
  mergedCells: CellRange[];
  columnWidths: ColumnWidths;
  rowCount: number;
  columnCount: number;
};

export type SpreadsheetWorkbook = {
  sheets: SpreadsheetSheet[];
  activeSheetId: string;
};

export type TemplateKind = "grades" | "expenses" | "sales";

export const INITIAL_ROWS = 14;
export const INITIAL_COLUMNS = 6;
export const DEFAULT_COLUMN_WIDTH = 78;
export const MIN_COLUMN_WIDTH = 58;
export const MAX_COLUMN_WIDTH = 180;

const HEADER_FORMAT: CellFormat = { backgroundColor: "#E9EEFF", textColor: "#2457E5" };
const ACCENT_FORMAT: CellFormat = { backgroundColor: "#E8F7F0", textColor: "#16865B" };

export const TEMPLATE_DETAILS: Record<TemplateKind, { title: string; description: string; icon: "school" | "account-balance-wallet" | "storefront" }> = {
  grades: { title: "كشف درجات", description: "طالب، اختباران، مجموع ومتوسط", icon: "school" },
  expenses: { title: "مصروفات شهرية", description: "سجّل البنود والمبالغ والفئات", icon: "account-balance-wallet" },
  sales: { title: "مبيعات", description: "منتج وكمية وسعر وإجمالي", icon: "storefront" },
};

function styleHeaders(addresses: string[]) {
  return addresses.reduce<CellFormats>((formats, address) => ({ ...formats, [address]: HEADER_FORMAT }), {});
}

function parseAddress(address: string) {
  const match = address.toUpperCase().match(/^([A-Z]+)([1-9]\d*)$/);
  if (!match) return null;
  const column = match[1].split("").reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1;
  return { column, row: Number(match[2]) - 1 };
}

function columnLabel(index: number) {
  let label = "";
  let current = index;
  do {
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26) - 1;
  } while (current >= 0);
  return label;
}

function normalizeArabicDigits(value: string) {
  const arabic = "٠١٢٣٤٥٦٧٨٩";
  const persian = "۰۱۲۳۴۵۶۷۸۹";
  return value
    .replace(/[٠-٩]/g, (digit) => String(arabic.indexOf(digit)))
    .replace(/[۰-۹]/g, (digit) => String(persian.indexOf(digit)));
}

function isRange(value: unknown): value is CellRange {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<CellRange>;
  return typeof candidate.start === "string" && typeof candidate.end === "string" && Boolean(parseAddress(candidate.start) && parseAddress(candidate.end));
}

export function rangeAddresses(start: string, end: string) {
  const first = parseAddress(start);
  const last = parseAddress(end);
  if (!first || !last) return [];
  const addresses: string[] = [];
  for (let row = Math.min(first.row, last.row); row <= Math.max(first.row, last.row); row += 1) {
    for (let column = Math.min(first.column, last.column); column <= Math.max(first.column, last.column); column += 1) {
      addresses.push(`${columnLabel(column)}${row + 1}`);
    }
  }
  return addresses;
}

export function getRangeBounds(range: CellRange) {
  const first = parseAddress(range.start);
  const last = parseAddress(range.end);
  if (!first || !last) return null;
  return {
    startColumn: Math.min(first.column, last.column),
    endColumn: Math.max(first.column, last.column),
    startRow: Math.min(first.row, last.row),
    endRow: Math.max(first.row, last.row),
  };
}

export function isAddressInRange(address: string, range: CellRange) {
  const current = parseAddress(address);
  const bounds = getRangeBounds(range);
  if (!current || !bounds) return false;
  return current.column >= bounds.startColumn && current.column <= bounds.endColumn && current.row >= bounds.startRow && current.row <= bounds.endRow;
}

export function getMergedRangeForAddress(sheet: SpreadsheetSheet, address: string) {
  return sheet.mergedCells.find((range) => isAddressInRange(address, range));
}

export function isMergedChild(sheet: SpreadsheetSheet, address: string) {
  const range = getMergedRangeForAddress(sheet, address);
  return Boolean(range && range.start !== address);
}

export function getColumnWidth(sheet: SpreadsheetSheet, column: string) {
  return Math.min(MAX_COLUMN_WIDTH, Math.max(MIN_COLUMN_WIDTH, sheet.columnWidths[column] ?? DEFAULT_COLUMN_WIDTH));
}

export function rangeNumericSummary(sheet: SpreadsheetSheet, range: CellRange) {
  const values = rangeAddresses(range.start, range.end)
    .map((address) => {
      const normalized = normalizeArabicDigits(displayCellValue(address, sheet.cells)).replace(/[^0-9.-]/g, "");
      return normalized ? Number(normalized) : Number.NaN;
    })
    .filter((value) => Number.isFinite(value));
  const total = values.reduce((sum, value) => sum + value, 0);
  return { total, average: values.length ? total / values.length : 0, count: values.length };
}

export function createEmptySheet(id = "sheet-1", name = "ورقة 1"): SpreadsheetSheet {
  return { id, name, cells: {}, cellFormats: {}, mergedCells: [], columnWidths: {}, rowCount: INITIAL_ROWS, columnCount: INITIAL_COLUMNS };
}

export function createInitialWorkbook(): SpreadsheetWorkbook {
  const sheet = createEmptySheet();
  return { sheets: [sheet], activeSheetId: sheet.id };
}

export function normalizeWorkbook(value: unknown): SpreadsheetWorkbook | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Partial<SpreadsheetWorkbook>;
  if (!Array.isArray(candidate.sheets) || !candidate.sheets.length || typeof candidate.activeSheetId !== "string") return null;

  const sheets = candidate.sheets
    .filter((sheet): sheet is SpreadsheetSheet => Boolean(sheet && typeof sheet.id === "string" && typeof sheet.name === "string"))
    .map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      cells: sheet.cells && typeof sheet.cells === "object" ? sheet.cells : {},
      cellFormats: sheet.cellFormats && typeof sheet.cellFormats === "object" ? sheet.cellFormats : {},
      mergedCells: Array.isArray(sheet.mergedCells) ? sheet.mergedCells.filter(isRange) : [],
      columnWidths: sheet.columnWidths && typeof sheet.columnWidths === "object" ? sheet.columnWidths : {},
      rowCount: Math.max(1, Number(sheet.rowCount) || INITIAL_ROWS),
      columnCount: Math.max(1, Number(sheet.columnCount) || INITIAL_COLUMNS),
    }));
  if (!sheets.length) return null;
  const activeSheetId = sheets.some((sheet) => sheet.id === candidate.activeSheetId) ? candidate.activeSheetId : sheets[0].id;
  return { sheets, activeSheetId };
}

export function formatCellDisplay(address: string, sheet: SpreadsheetSheet) {
  const value = displayCellValue(address, sheet.cells);
  if (!value || value === "#خطأ") return value;
  const format = sheet.cellFormats[address]?.numberFormat ?? "general";
  const numeric = Number(normalizeArabicDigits(value).replace(/,/g, ""));
  if (!Number.isFinite(numeric) || format === "general") return value;
  if (format === "currency") return `${numeric.toFixed(2)} ر.س`;
  if (format === "percent") return `${(numeric * 100).toFixed(0)}%`;
  return numeric.toFixed(2);
}

export function formatLabel(format: CellNumberFormat) {
  const labels: Record<CellNumberFormat, string> = { general: "عادي", currency: "عملة", percent: "نسبة", decimal: "منزلتان" };
  return labels[format];
}

export function createTemplateSheet(kind: TemplateKind, id: string): SpreadsheetSheet {
  const details = TEMPLATE_DETAILS[kind];
  const base = createEmptySheet(id, details.title);

  if (kind === "grades") {
    return {
      ...base,
      columnCount: 5,
      cells: { A1: "الطالب", B1: "اختبار ١", C1: "اختبار ٢", D1: "المجموع", E1: "المتوسط", A2: "أحمد محمد", B2: "٥", C2: "٥", D2: "=B2+C2", E2: "=AVERAGE(B2:C2)" },
      cellFormats: { ...styleHeaders(["A1", "B1", "C1", "D1", "E1"]), D2: ACCENT_FORMAT, E2: ACCENT_FORMAT },
    };
  }

  if (kind === "expenses") {
    return {
      ...base,
      columnCount: 4,
      cells: { A1: "البند", B1: "المبلغ", C1: "الفئة", D1: "الإجمالي", A2: "مواصلات", B2: "25", C2: "تنقل", A3: "وجبة", B3: "40", C3: "طعام", D2: "=SUM(B2:B3)" },
      cellFormats: { ...styleHeaders(["A1", "B1", "C1", "D1"]), B2: { numberFormat: "currency" }, B3: { numberFormat: "currency" }, D2: { ...ACCENT_FORMAT, numberFormat: "currency" } },
    };
  }

  return {
    ...base,
    columnCount: 4,
    cells: { A1: "المنتج", B1: "الكمية", C1: "سعر الوحدة", D1: "الإجمالي", A2: "دفتر", B2: "3", C2: "12", D2: "=B2*C2" },
    cellFormats: { ...styleHeaders(["A1", "B1", "C1", "D1"]), C2: { numberFormat: "currency" }, D2: { ...ACCENT_FORMAT, numberFormat: "currency" } },
  };
}
