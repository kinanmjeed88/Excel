import { describe, expect, it } from "vitest";

import { appendFormulaDraftToken, buildRelativeFormula, getRelativeFormulaReferences } from "../lib/formula-editor";
import { displayCellValue } from "../lib/formula-engine";

describe("appendFormulaDraftToken", () => {
  it("ينشئ الصيغة التعليمية =B2+C2 من رموز الشريط السريع", () => {
    const reference = appendFormulaDraftToken("", "B2");
    const plus = appendFormulaDraftToken(reference, "+");
    expect(appendFormulaDraftToken(plus, "C2")).toBe("=B2+C2");
  });

  it("يحافظ على صيغة بادئة جاهزة ولا يكرر علامة = الأولى", () => {
    expect(appendFormulaDraftToken("=", "=")).toBe("=");
    expect(appendFormulaDraftToken("=SUM(", "B2:B6)")).toBe("=SUM(B2:B6)");
  });

  it("يحوّل قيمة قائمة إلى صيغة عند إدراج رمز حسابي", () => {
    expect(appendFormulaDraftToken("٥", "+C2")).toBe("=٥+C2");
  });

  it("ينشئ مراجع نسبية للصف المحدد بدلاً من تثبيت مراجع الصف 2", () => {
    expect(getRelativeFormulaReferences("D3")).toEqual({ row: 3, firstReference: "B3", secondReference: "C3" });
    expect(buildRelativeFormula("D3", "add")).toBe("=B3+C3");
    expect(buildRelativeFormula("D3", "sum")).toBe("=SUM(B3:C3)");
    expect(buildRelativeFormula("D3", "average")).toBe("=AVERAGE(B3:C3)");
  });

  it("يحسِب D3 من قيم الصف الثالث لا من مثال الصف الثاني", () => {
    const cells = { B2: "5", C2: "5", B3: "9", C3: "7", D3: buildRelativeFormula("D3", "add") };
    expect(displayCellValue("D3", cells)).toBe("16");
  });

  it("يبني مراجع الصف الأول عند تحرير خلية في هذا الصف", () => {
    expect(getRelativeFormulaReferences("D1")).toEqual({ row: 1, firstReference: "B1", secondReference: "C1" });
    expect(buildRelativeFormula("D1", "add")).toBe("=B1+C1");
  });

  it("لا يلحق مرجعاً جديداً بصيغة مكتملة بشكل غير صالح", () => {
    expect(appendFormulaDraftToken("=B2+C2", "B3")).toBe("=B3");
    expect(appendFormulaDraftToken("=B3", "+")).toBe("=B3+");
    expect(appendFormulaDraftToken("=", "+")).toBe("=");
  });
});
