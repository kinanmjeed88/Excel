export type CellValues = Record<string, string>;

export type FormulaResult =
  | { kind: "number"; value: number; display: string }
  | { kind: "text"; display: string }
  | { kind: "error"; display: "#خطأ"; message: string };

const CELL_REFERENCE = /^[A-Z]+[1-9]\d*$/;

function columnNumber(label: string) {
  return label.split("").reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
}

function columnLabel(index: number) {
  let label = "";
  let current = index;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    current = Math.floor((current - 1) / 26);
  }
  return label;
}

function rangeReferences(start: string, end: string) {
  const startMatch = start.match(/^([A-Z]+)(\d+)$/);
  const endMatch = end.match(/^([A-Z]+)(\d+)$/);
  if (!startMatch || !endMatch) return [];

  const startColumn = columnNumber(startMatch[1]);
  const endColumn = columnNumber(endMatch[1]);
  const startRow = Number(startMatch[2]);
  const endRow = Number(endMatch[2]);
  const references: string[] = [];

  for (let column = Math.min(startColumn, endColumn); column <= Math.max(startColumn, endColumn); column += 1) {
    for (let row = Math.min(startRow, endRow); row <= Math.max(startRow, endRow); row += 1) {
      references.push(`${columnLabel(column)}${row}`);
    }
  }
  return references;
}

function formatNumber(value: number) {
  if (Number.isInteger(value)) return String(value);
  return String(Number(value.toFixed(8)));
}

function calculateExpression(input: string) {
  const normalized = input.replace(/\s+/g, "");
  const tokens = normalized.match(/\d+(?:\.\d+)?|[()+\-*/]/g) ?? [];
  if (!tokens.length || tokens.join("") !== normalized) throw new Error("صيغة غير صالحة");

  let position = 0;
  const peek = () => tokens[position];
  const consume = () => tokens[position++];

  const parseFactor = (): number => {
    const token = consume();
    if (token === "-") return -parseFactor();
    if (token === "+") return parseFactor();
    if (token === "(") {
      const value = parseExpression();
      if (consume() !== ")") throw new Error("قوس غير مكتمل");
      return value;
    }
    if (!token || !/^\d+(?:\.\d+)?$/.test(token)) throw new Error("قيمة غير صالحة");
    return Number(token);
  };

  const parseTerm = (): number => {
    let value = parseFactor();
    while (peek() === "*" || peek() === "/") {
      const operator = consume();
      const next = parseFactor();
      if (operator === "/" && next === 0) throw new Error("لا يمكن القسمة على صفر");
      value = operator === "*" ? value * next : value / next;
    }
    return value;
  };

  const parseExpression = (): number => {
    let value = parseTerm();
    while (peek() === "+" || peek() === "-") {
      const operator = consume();
      const next = parseTerm();
      value = operator === "+" ? value + next : value - next;
    }
    return value;
  };

  const result = parseExpression();
  if (position !== tokens.length || !Number.isFinite(result)) throw new Error("ناتج غير صالح");
  return result;
}

export function evaluateCell(
  address: string,
  cells: CellValues,
  visiting = new Set<string>(),
): FormulaResult {
  const normalizedAddress = address.toUpperCase();
  const raw = (cells[normalizedAddress] ?? "").trim();
  if (!raw) return { kind: "number", value: 0, display: "0" };
  if (!raw.startsWith("=")) {
    const numeric = Number(raw.replace(/,/g, ""));
    return Number.isFinite(numeric) && raw !== ""
      ? { kind: "number", value: numeric, display: formatNumber(numeric) }
      : { kind: "text", display: raw };
  }
  if (visiting.has(normalizedAddress)) {
    return { kind: "error", display: "#خطأ", message: "مرجع دائري" };
  }

  const nextVisiting = new Set(visiting);
  nextVisiting.add(normalizedAddress);

  try {
    let expression = raw.slice(1).toUpperCase();
    expression = expression.replace(/\b(SUM|AVG|AVERAGE)\(([A-Z]+\d+):([A-Z]+\d+)\)/g, (_, fn, start, end) => {
      const values = rangeReferences(start, end).map((reference) => {
        const result = evaluateCell(reference, cells, nextVisiting);
        if (result.kind !== "number") throw new Error("النطاق يحتوي قيمة نصية أو غير صالحة");
        return result.value;
      });
      if (!values.length) throw new Error("نطاق غير صالح");
      const sum = values.reduce((total, value) => total + value, 0);
      return String(fn === "SUM" ? sum : sum / values.length);
    });

    expression = expression.replace(/\b([A-Z]+[1-9]\d*)\b/g, (reference) => {
      if (!CELL_REFERENCE.test(reference)) throw new Error("مرجع غير صالح");
      const result = evaluateCell(reference, cells, nextVisiting);
      if (result.kind !== "number") throw new Error("لا يمكن إجراء حساب على نص");
      return String(result.value);
    });

    const value = calculateExpression(expression);
    return { kind: "number", value, display: formatNumber(value) };
  } catch (error) {
    return {
      kind: "error",
      display: "#خطأ",
      message: error instanceof Error ? error.message : "تعذر حساب الصيغة",
    };
  }
}

export function displayCellValue(address: string, cells: CellValues) {
  const result = evaluateCell(address, cells);
  return result.display === "0" && !cells[address] ? "" : result.display;
}
