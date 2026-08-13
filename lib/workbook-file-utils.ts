import * as XLSX from "xlsx";

import { displayCellValue, type CellValues } from "./formula-engine";

export type PortableSheet = {
  id: string;
  name: string;
  cells: CellValues;
  rowCount: number;
  columnCount: number;
};

export type ImportSummary = {
  sheets: PortableSheet[];
  truncated: boolean;
};

const MAX_IMPORT_ROWS = 200;
const MAX_IMPORT_COLUMNS = 40;
const DEFAULT_ROWS = 14;
const DEFAULT_COLUMNS = 6;

function safeSheetName(name: string, fallback: string) {
  const cleaned = name.replace(/[\\/*?:\[\]]/g, " ").trim().slice(0, 31);
  return cleaned || fallback;
}

function cellAddress(column: number, row: number) {
  return XLSX.utils.encode_cell({ c: column, r: row });
}

function readCellValue(cell: XLSX.CellObject | undefined) {
  if (!cell) return "";
  if (cell.f) return `=${cell.f}`;
  if (cell.v === undefined || cell.v === null) return "";
  return String(cell.w ?? cell.v);
}

function worksheetToPortableSheet(
  worksheet: XLSX.WorkSheet,
  sheetName: string,
  sheetIndex: number,
): { sheet: PortableSheet; truncated: boolean } {
  const range = worksheet["!ref"] ? XLSX.utils.decode_range(worksheet["!ref"]) : null;
  const sourceRows = range ? range.e.r + 1 : 0;
  const sourceColumns = range ? range.e.c + 1 : 0;
  const rowCount = Math.max(DEFAULT_ROWS, Math.min(sourceRows, MAX_IMPORT_ROWS));
  const columnCount = Math.max(DEFAULT_COLUMNS, Math.min(sourceColumns, MAX_IMPORT_COLUMNS));
  const cells: CellValues = {};

  if (range) {
    for (let row = 0; row < Math.min(sourceRows, MAX_IMPORT_ROWS); row += 1) {
      for (let column = 0; column < Math.min(sourceColumns, MAX_IMPORT_COLUMNS); column += 1) {
        const address = cellAddress(column, row);
        const value = readCellValue(worksheet[address]);
        if (value) cells[address] = value;
      }
    }
  }

  return {
    sheet: {
      id: `imported-${Date.now()}-${sheetIndex}`,
      name: safeSheetName(sheetName, `ورقة ${sheetIndex + 1}`),
      cells,
      rowCount,
      columnCount,
    },
    truncated: sourceRows > MAX_IMPORT_ROWS || sourceColumns > MAX_IMPORT_COLUMNS,
  };
}

export function importSpreadsheetFile(fileName: string, contents: ArrayBuffer | string): ImportSummary {
  const isCsv = fileName.toLowerCase().endsWith(".csv");
  const workbook = isCsv
    ? XLSX.read(typeof contents === "string" ? contents : new TextDecoder().decode(contents), { type: "string", cellFormula: true })
    : XLSX.read(contents, { type: "array", cellFormula: true });

  const output = workbook.SheetNames.slice(0, 12).map((name, index) =>
    worksheetToPortableSheet(workbook.Sheets[name], name, index),
  );

  return {
    sheets: output.map((item) => item.sheet),
    truncated: workbook.SheetNames.length > 12 || output.some((item) => item.truncated),
  };
}

function toWorksheet(sheet: PortableSheet) {
  const worksheet: XLSX.WorkSheet = {};
  const maxRow = Math.max(0, sheet.rowCount - 1);
  const maxColumn = Math.max(0, sheet.columnCount - 1);
  worksheet["!ref"] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: maxColumn, r: maxRow } });

  Object.entries(sheet.cells).forEach(([address, raw]) => {
    if (raw.startsWith("=")) {
      const display = displayCellValue(address, sheet.cells);
      const numeric = Number(display);
      worksheet[address] = Number.isFinite(numeric)
        ? { t: "n", f: raw.slice(1), v: numeric }
        : { t: "s", f: raw.slice(1), v: display };
      return;
    }
    const numeric = Number(raw.replace(/,/g, ""));
    worksheet[address] = Number.isFinite(numeric) && raw.trim() !== "" ? { t: "n", v: numeric } : { t: "s", v: raw };
  });

  return worksheet;
}

export function exportWorkbookToXlsx(sheets: PortableSheet[]) {
  const workbook = XLSX.utils.book_new();
  sheets.forEach((sheet, index) => {
    XLSX.utils.book_append_sheet(workbook, toWorksheet(sheet), safeSheetName(sheet.name, `ورقة ${index + 1}`));
  });
  return XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true }) as ArrayBuffer;
}

function escapeCsv(value: string) {
  return /[",\n\r]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

export function exportSheetToCsv(sheet: PortableSheet) {
  const rows: string[] = [];
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    const values: string[] = [];
    for (let column = 0; column < sheet.columnCount; column += 1) {
      const address = cellAddress(column, row - 1);
      values.push(escapeCsv(displayCellValue(address, sheet.cells)));
    }
    rows.push(values.join(","));
  }
  return `\uFEFF${rows.join("\r\n")}`;
}
