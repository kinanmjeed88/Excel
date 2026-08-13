import { describe, expect, it } from "vitest";

import {
  createEmptySheet,
  createTemplateSheet,
  formatCellDisplay,
  getColumnWidth,
  getMergedRangeForAddress,
  isMergedChild,
  normalizeWorkbook,
  rangeAddresses,
  rangeNumericSummary,
} from "../lib/spreadsheet-model";

describe("نموذج المصنف المحسن", () => {
  it("ينشئ قالب كشف درجات يحسب ٥ + ٥ = 10", () => {
    const sheet = createTemplateSheet("grades", "grades-1");
    expect(sheet.cells.D2).toBe("=B2+C2");
    expect(formatCellDisplay("D2", sheet)).toBe("10");
    expect(sheet.cellFormats.A1?.backgroundColor).toBe("#E9EEFF");
  });

  it("ينسق العملة والنسبة المئوية والمنازل العشرية", () => {
    const sheet = createTemplateSheet("sales", "sales-1");
    sheet.cells.A3 = "0.25";
    sheet.cellFormats.A3 = { numberFormat: "percent" };
    sheet.cells.A4 = "5";
    sheet.cellFormats.A4 = { numberFormat: "decimal" };
    expect(formatCellDisplay("C2", sheet)).toBe("12.00 ر.س");
    expect(formatCellDisplay("A3", sheet)).toBe("25%");
    expect(formatCellDisplay("A4", sheet)).toBe("5.00");
  });

  it("يتوافق مع المصنفات المحفوظة قبل إضافة تنسيق الخلايا", () => {
    const workbook = normalizeWorkbook({ sheets: [{ id: "one", name: "ورقة 1", cells: { A1: "5" }, rowCount: 14, columnCount: 6 }], activeSheetId: "one" });
    expect(workbook?.sheets[0].cellFormats).toEqual({});
  });

  it("يحلل نطاقاً يتضمن أرقاماً عربية ويستبعد القيم النصية", () => {
    const sheet = createEmptySheet();
    sheet.cells = { A1: "٥", B1: "7", C1: "اسم", A2: "=A1+B1" };
    expect(rangeAddresses("A1", "B2")).toEqual(["A1", "B1", "A2", "B2"]);
    expect(rangeNumericSummary(sheet, { start: "A1", end: "C1" })).toEqual({ total: 12, average: 6, count: 2 });
    expect(rangeNumericSummary(sheet, { start: "A1", end: "A2" })).toEqual({ total: 17, average: 8.5, count: 2 });
  });

  it("يتعرف على الخلايا المدمجة ويقيد عرض الأعمدة داخل الحدود المسموحة", () => {
    const sheet = createEmptySheet();
    sheet.mergedCells = [{ start: "A1", end: "C1" }];
    sheet.columnWidths = { A: 20, B: 250, C: 120 };
    expect(getMergedRangeForAddress(sheet, "B1")).toEqual({ start: "A1", end: "C1" });
    expect(isMergedChild(sheet, "B1")).toBe(true);
    expect(isMergedChild(sheet, "A1")).toBe(false);
    expect(getColumnWidth(sheet, "A")).toBe(58);
    expect(getColumnWidth(sheet, "B")).toBe(180);
    expect(getColumnWidth(sheet, "C")).toBe(120);
  });
});
