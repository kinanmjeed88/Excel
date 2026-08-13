export type RelativeFormulaPreset = "add" | "subtract" | "multiply" | "divide" | "sum" | "average";

type FormulaReferences = {
  row: number;
  firstReference: string;
  secondReference: string;
};

function formulaRowForAddress(address: string) {
  const match = address.trim().toUpperCase().match(/^[A-Z]+([1-9]\d*)$/);
  return Number(match?.[1] ?? 1);
}

function isCellReference(token: string) {
  return /^[A-Z]+[1-9]\d*$/i.test(token);
}

function isFunctionStart(token: string) {
  return /^[A-Z]+\($/i.test(token);
}

function expectsOperand(expression: string) {
  return expression === "=" || /[+\-*/(:,]$/.test(expression);
}

function endsWithOperand(expression: string) {
  return /(?:[A-Z]+[1-9]\d*|\d|\))$/i.test(expression);
}

/** يعيد مراجع الدرجات الموازية للصف الجاري تحريره. */
export function getRelativeFormulaReferences(address: string): FormulaReferences {
  const row = formulaRowForAddress(address);
  return { row, firstReference: `B${row}`, secondReference: `C${row}` };
}

/** ينشئ صيغة مكتملة بمراجع نسبية، لذا لا تُعاد استخدام مراجع صف ثابت. */
export function buildRelativeFormula(address: string, preset: RelativeFormulaPreset) {
  const { firstReference, secondReference } = getRelativeFormulaReferences(address);
  const operands = `${firstReference}${preset === "add" ? "+" : preset === "subtract" ? "-" : preset === "multiply" ? "*" : "/"}${secondReference}`;

  if (preset === "sum") return `=SUM(${firstReference}:${secondReference})`;
  if (preset === "average") return `=AVERAGE(${firstReference}:${secondReference})`;
  return `=${operands}`;
}

/**
 * يضيف رمزاً صالحاً إلى مسودة الصيغة. عند اختيار مرجع جديد بعد صيغة مكتملة
 * يبدأ المرجع صيغة جديدة بدلاً من إنتاج تعبير خاطئ مثل =B2+C2B3.
 */
export function appendFormulaDraftToken(currentDraft: string, token: string) {
  const nextToken = token.trim();
  let expression = currentDraft.trim();
  if (!nextToken) return expression;
  if (nextToken === "=") return expression.startsWith("=") ? expression : "=";

  if (!expression) {
    return isCellReference(nextToken) || isFunctionStart(nextToken) ? `=${nextToken}` : "=";
  }
  if (!expression.startsWith("=")) expression = `=${expression}`;

  if (isCellReference(nextToken)) {
    return expectsOperand(expression) ? `${expression}${nextToken}` : `=${nextToken}`;
  }
  if (/^[+\-*/]$/.test(nextToken)) {
    return endsWithOperand(expression) ? `${expression}${nextToken}` : expression;
  }
  if (isFunctionStart(nextToken)) {
    return expectsOperand(expression) ? `${expression}${nextToken}` : `=${nextToken}`;
  }
  return `${expression}${nextToken}`;
}
