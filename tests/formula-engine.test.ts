import { describe, expect, it } from "vitest";

import { evaluateCell } from "../lib/formula-engine";

describe("محرك الصيغ", () => {
  it("ينفذ عمليات الجمع والطرح والضرب والقسمة بحسب أولوية العمليات", () => {
    const cells = { A1: "5", B1: "7", C1: "=A1+B1*2", D1: "=C1/3-1" };
    expect(evaluateCell("C1", cells)).toMatchObject({ kind: "number", value: 19, display: "19" });
    expect(evaluateCell("D1", cells)).toMatchObject({ kind: "number", value: 16 / 3, display: "5.33333333" });
  });

  it("يدعم دالتي SUM وAVERAGE على نطاق من الخلايا", () => {
    const cells = { A1: "3", A2: "7", A3: "10", B1: "=SUM(A1:A3)", B2: "=AVERAGE(A1:A3)" };
    expect(evaluateCell("B1", cells)).toMatchObject({ kind: "number", value: 20, display: "20" });
    expect(evaluateCell("B2", cells)).toMatchObject({ kind: "number", value: 20 / 3, display: "6.66666667" });
  });

  it("يعيد حالة خطأ عند القسمة على صفر أو المراجع الدائرية", () => {
    expect(evaluateCell("A1", { A1: "=4/0" })).toMatchObject({ kind: "error", display: "#خطأ" });
    expect(evaluateCell("A1", { A1: "=B1+1", B1: "=A1+1" })).toMatchObject({ kind: "error", display: "#خطأ" });
  });
});
