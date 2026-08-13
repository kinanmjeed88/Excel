import { describe, expect, it } from "vitest";

import { appendFormulaDraftToken } from "../lib/formula-editor";

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
});
