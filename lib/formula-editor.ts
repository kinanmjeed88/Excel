export type RelativeFormulaPreset = "add" | "subtract" | "multiply" | "divide" | "sum" | "average";

type ParsedAddress = {
  column: number;
  row: number;
};

type FormulaReferences = {
  row: number;
  firstReference: string;
  secondReference: string;
};

function formulaRowForAddress(address: string) {
  const match = address.trim().toUpperCase().match(/^[A-Z]+([1-9]\d*)$/);
  return Number(match?.[1] ?? 1);
}

function columnIndex(label: string) {
  return label.split("").reduce((sum, character) => sum * 26 + character.charCodeAt(0) - 64, 0) - 1;
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

function parseCellAddress(address: string): ParsedAddress | null {
  const match = address.trim().toUpperCase().match(/^([A-Z]+)([1-9]\d*)$/);
  if (!match) return null;
  return { column: columnIndex(match[1]), row: Number(match[2]) };
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

/** يعيد الخلايا الهدف الواقعة بين خلية الصيغة والخلية التي انتهى عندها السحب. */
export function getFormulaFillTargets(sourceAddress: string, targetAddress: string) {
  const source = parseCellAddress(sourceAddress);
  const target = parseCellAddress(targetAddress);
  if (!source || !target || (source.column !== target.column && source.row !== target.row)) return [];

  const targets: string[] = [];
  if (source.column === target.column) {
    const direction = target.row >= source.row ? 1 : -1;
    for (let row = source.row + direction; direction > 0 ? row <= target.row : row >= target.row; row += direction) {
      targets.push(`${columnLabel(source.column)}${row}`);
    }
    return targets;
  }

  const direction = target.column >= source.column ? 1 : -1;
  for (let column = source.column + direction; direction > 0 ? column <= target.column : column >= target.column; column += direction) {
    targets.push(`${columnLabel(column)}${source.row}`);
  }
  return targets;
}

/**
 * يحرك كل مراجع الخلايا داخل صيغة بمقدار المسافة بين المصدر والهدف،
 * ليحاكي التعبئة النسبية في جداول البيانات من دون تثبيت مراجع صف سابق.
 */
export function translateFormulaForFill(formula: string, sourceAddress: string, targetAddress: string) {
  const source = parseCellAddress(sourceAddress);
  const target = parseCellAddress(targetAddress);
  if (!source || !target || !formula.trim().startsWith("=")) return formula;

  const rowDelta = target.row - source.row;
  const columnDelta = target.column - source.column;
  return formula.replace(/\b([A-Z]+)([1-9]\d*)\b/gi, (reference, columnLabelValue: string, rowValue: string) => {
    const nextColumn = columnIndex(columnLabelValue.toUpperCase()) + columnDelta;
    const nextRow = Number(rowValue) + rowDelta;
    return nextColumn < 0 || nextRow < 1 ? reference : `${columnLabel(nextColumn)}${nextRow}`;
  });
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
