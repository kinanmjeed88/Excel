import { describe, expect, it } from "vitest";

import {
  exportSheetToCsv,
  exportWorkbookToXlsx,
  importSpreadsheetFile,
  type PortableSheet,
} from "../lib/workbook-file-utils";

const sheet: PortableSheet = {
  id: "sheet-1",
  name: "المبيعات",
  rowCount: 3,
  columnCount: 2,
  cells: { A1: "اسم", B1: "القيمة", A2: "قلم, أزرق", B2: "8", B3: "=B2*2" },
};

describe("تحويل ملفات المصنف", () => {
  it("يصدر CSV بالقيم المحسوبة ويهّرب الفواصل", () => {
    const csv = exportSheetToCsv(sheet);
    expect(csv).toContain('"قلم, أزرق",8');
    expect(csv).toContain(",16");
  });

  it("يصدر Excel ويعيد استيراد القيم والصيغ والأوراق", () => {
    const file = exportWorkbookToXlsx([sheet]);
    const imported = importSpreadsheetFile("المبيعات.xlsx", file);
    expect(imported.sheets).toHaveLength(1);
    expect(imported.sheets[0]).toMatchObject({ name: "المبيعات" });
    expect(imported.sheets[0].cells).toMatchObject({ A1: "اسم", A2: "قلم, أزرق", B3: "=B2*2" });
  });

  it("يستورد ملف CSV إلى ورقة عمل", () => {
    const buffer = new TextEncoder().encode("البند,العدد\nدفتر,3").buffer;
    const imported = importSpreadsheetFile("قائمة.csv", buffer);
    expect(imported.sheets[0].cells).toMatchObject({ A1: "البند", B1: "العدد", A2: "دفتر", B2: "3" });
  });
});
