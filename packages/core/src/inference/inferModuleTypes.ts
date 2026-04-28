import { getAccessibleSymbolsAtLine, resolveSymbolAtPosition } from "../symbol/buildModuleSymbols";
import { normalizeIdentifier } from "../types/helpers";
import type {
  AnalysisResult,
  Diagnostic,
  InferredSymbolType,
  ParseResult,
  ProcedureDeclarationNode,
  SymbolInfo,
  SymbolTable,
  TypeInferenceResult
} from "../types/model";

const CAST_FUNCTION_TYPES = new Map<string, string>([
  ["cbool", "Boolean"],
  ["cbyte", "Byte"],
  ["ccur", "Currency"],
  ["cdate", "Date"],
  ["cdbl", "Double"],
  ["cint", "Integer"],
  ["clng", "Long"],
  ["csng", "Single"],
  ["cstr", "String"],
  ["cvar", "Variant"]
]);

const NUMERIC_TYPES = new Set(["byte", "currency", "double", "integer", "long", "longlong", "longptr", "single"]);
const SCALAR_TYPES = new Set(["boolean", "date", "nothing", "string", "variant", ...NUMERIC_TYPES]);
const CREATE_OBJECT_PROGID_TYPES = new Map<string, string>([
  ["wscript.shell", "WshShell"]
]);

export function inferModuleTypes(parseResult: ParseResult, symbolTable: SymbolTable): TypeInferenceResult {
  const diagnostics: Diagnostic[] = [];
  const symbolTypes = new Map<string, InferredSymbolType>();

  seedExplicitTypes(symbolTable, symbolTypes);

  for (const member of parseResult.module.members) {
    if (member.kind !== "procedureDeclaration") {
      continue;
    }

    for (const statement of member.body) {
      const assignment =
        statement.kind === "assignmentStatement"
          ? {
              expressionRange: statement.expressionRange,
              expressionText: statement.expressionText,
              isSet: statement.assignmentKind === "set",
              targetName: statement.targetName,
              targetRange: statement.targetRange
            }
          : undefined;

      if (!assignment) {
        continue;
      }

      if (!assignment.targetName) {
        continue;
      }

      const targetSymbol = resolveSymbolAtPosition(symbolTable, assignment.targetName, assignment.targetRange.start);
      const inferredExpressionType = inferExpressionType(symbolTable, symbolTypes, statement.range.start.line, assignment.expressionText);

      if (!targetSymbol || !inferredExpressionType) {
        continue;
      }

      const targetTypeName = getSymbolTypeNameFromMap(symbolTypes, targetSymbol) ?? targetSymbol.typeName;

      if (targetTypeName) {
        const missingSetAssignment = shouldWarnMissingSetAssignment(
          symbolTable,
          targetSymbol,
          targetTypeName,
          inferredExpressionType,
          assignment.isSet
        );
        const compatibleWithSet = areTypesCompatible(targetTypeName, inferredExpressionType, { isSetAssignment: true });

        if (missingSetAssignment) {
          diagnostics.push({
            code: "set-required",
            message: `Set is required to assign ${inferredExpressionType} to ${targetTypeName}.`,
            range: assignment.targetRange,
            severity: "warning"
          });
        }

        if (!areTypesCompatible(targetTypeName, inferredExpressionType, { isSetAssignment: assignment.isSet }) && (!missingSetAssignment || !compatibleWithSet)) {
          diagnostics.push({
            code: "type-mismatch",
            message: `Type mismatch: cannot assign ${inferredExpressionType} to ${targetTypeName}.`,
            range: assignment.expressionRange,
            severity: "warning"
          });
        }
      } else {
        setSymbolType(symbolTypes, targetSymbol, inferredExpressionType, "assignment");
      }

      const procedureSymbol = findProcedureSymbolForReturn(symbolTable, member, targetSymbol);

      if (procedureSymbol) {
        const procedureTypeName = getSymbolTypeNameFromMap(symbolTypes, procedureSymbol) ?? procedureSymbol.typeName;

        if (!procedureTypeName) {
          setSymbolType(symbolTypes, procedureSymbol, inferredExpressionType, "return");
        }
      }
    }
  }

  return {
    diagnostics,
    symbolTypes: [...symbolTypes.values()]
  };
}

export function getSymbolTypeName(
  result: Pick<AnalysisResult, "typeInference"> | TypeInferenceResult,
  symbol: SymbolInfo | undefined
): string | undefined {
  if (!symbol) {
    return undefined;
  }

  const typeInference = "typeInference" in result ? result.typeInference : result;
  return symbol.typeName ?? getSymbolTypeNameFromMap(typeInference.symbolTypes, symbol);
}

export function inferExpressionTypeAtLine(
  result: Pick<AnalysisResult, "symbols" | "typeInference">,
  line: number,
  expressionText: string
): string | undefined {
  return inferExpressionType(result.symbols, createSymbolTypeMap(result.typeInference.symbolTypes), line, expressionText);
}

function seedExplicitTypes(symbolTable: SymbolTable, sink: Map<string, InferredSymbolType>): void {
  for (const symbol of symbolTable.allSymbols) {
    if (symbol.typeName) {
      setSymbolType(sink, symbol, symbol.typeName, "explicit");
    }
  }
}

function inferExpressionType(
  symbolTable: SymbolTable,
  symbolTypes: Map<string, InferredSymbolType>,
  line: number,
  expressionText: string
): string | undefined {
  const normalizedExpression = unwrapParentheses(expressionText.trim());

  if (/^"(?:[^"]|"")*"$/u.test(normalizedExpression)) {
    return "String";
  }

  if (/^#.+#$/u.test(normalizedExpression)) {
    return "Date";
  }

  if (/^Nothing$/iu.test(normalizedExpression)) {
    return "Nothing";
  }

  if (/^New\s+([A-Za-z_][A-Za-z0-9_\.]*)$/iu.test(normalizedExpression)) {
    return /^New\s+([A-Za-z_][A-Za-z0-9_\.]*)$/iu.exec(normalizedExpression)?.[1];
  }

  if (/^(?:True|False)$/iu.test(normalizedExpression)) {
    return "Boolean";
  }

  if (/^[+-]?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?$/u.test(normalizedExpression)) {
    return /[.Ee]/u.test(normalizedExpression) ? "Double" : "Long";
  }

  if (/^(?:CreateObject|GetObject)\s*\(/iu.test(normalizedExpression)) {
    return inferCreateObjectType(normalizedExpression) ?? "Object";
  }

  if (/^Array\s*\(/iu.test(normalizedExpression)) {
    return "Variant";
  }

  const comparisonType = inferComparisonExpressionType(symbolTable, symbolTypes, line, normalizedExpression);

  if (comparisonType) {
    return comparisonType;
  }

  const concatenationType = inferConcatenationExpressionType(symbolTable, symbolTypes, line, normalizedExpression);

  if (concatenationType) {
    return concatenationType;
  }

  const arithmeticType = inferArithmeticExpressionType(symbolTable, symbolTypes, line, normalizedExpression);

  if (arithmeticType) {
    return arithmeticType;
  }

  const callMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)\s*\((.*)\)$/iu.exec(normalizedExpression);

  if (callMatch) {
    const castType = CAST_FUNCTION_TYPES.get(normalizeIdentifier(callMatch[1]));

    if (castType) {
      return castType;
    }

    const symbol = resolveSymbolAtPosition(symbolTable, callMatch[1], { character: 0, line });
    return symbol ? getResolvedTypeName(symbolTypes, symbol) : undefined;
  }

  const identifierMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)$/iu.exec(normalizedExpression);

  if (!identifierMatch) {
    return undefined;
  }

  const symbol = resolveSymbolAtPosition(symbolTable, identifierMatch[1], { character: 0, line });
  return symbol ? getResolvedTypeName(symbolTypes, symbol) : undefined;
}

function findProcedureSymbolForReturn(
  symbolTable: SymbolTable,
  procedure: ProcedureDeclarationNode,
  targetSymbol: SymbolInfo
): SymbolInfo | undefined {
  if (
    procedure.procedureKind === "Sub" ||
    targetSymbol.scope !== "procedure" ||
    targetSymbol.normalizedName !== normalizeIdentifier(procedure.name)
  ) {
    return undefined;
  }

  return symbolTable.moduleSymbols.find(
    (symbol) =>
      symbol.kind === "procedure" &&
      symbol.normalizedName === targetSymbol.normalizedName &&
      symbol.selectionRange.start.line === procedure.headerRange.start.line &&
      symbol.selectionRange.start.character === procedure.headerRange.start.character
  );
}

function setSymbolType(
  sink: Map<string, InferredSymbolType>,
  symbol: SymbolInfo,
  typeName: string,
  source: InferredSymbolType["source"]
): void {
  const key = getSymbolKey(symbol);
  const current = sink.get(key);

  if (!current || compareSourcePrecedence(source, current.source) >= 0) {
    sink.set(key, {
      source,
      symbol,
      typeName
    });
  }
}

function getResolvedTypeName(symbolTypes: Map<string, InferredSymbolType>, symbol: SymbolInfo): string | undefined {
  return symbol.typeName ?? symbolTypes.get(getSymbolKey(symbol))?.typeName;
}

function getSymbolTypeNameFromMap(symbolTypes: Map<string, InferredSymbolType> | InferredSymbolType[], symbol: SymbolInfo): string | undefined {
  if (Array.isArray(symbolTypes)) {
    return symbolTypes.find((entry) => getSymbolKey(entry.symbol) === getSymbolKey(symbol))?.typeName;
  }

  return symbolTypes.get(getSymbolKey(symbol))?.typeName;
}

function createSymbolTypeMap(symbolTypes: InferredSymbolType[]): Map<string, InferredSymbolType> {
  return new Map(symbolTypes.map((entry) => [getSymbolKey(entry.symbol), entry]));
}

function getSymbolKey(symbol: SymbolInfo): string {
  return `${symbol.scope}:${symbol.kind}:${symbol.normalizedName}:${symbol.selectionRange.start.line}:${symbol.selectionRange.start.character}:${symbol.selectionRange.end.line}:${symbol.selectionRange.end.character}`;
}

function compareSourcePrecedence(left: InferredSymbolType["source"], right: InferredSymbolType["source"]): number {
  return getSourcePrecedence(left) - getSourcePrecedence(right);
}

function getSourcePrecedence(source: InferredSymbolType["source"]): number {
  switch (source) {
    case "explicit":
      return 3;
    case "return":
      return 2;
    default:
      return 1;
  }
}

export function areTypesCompatible(
  targetTypeName: string,
  valueTypeName: string,
  options: { isSetAssignment?: boolean } = {}
): boolean {
  const normalizedTargetType = normalizeTypeName(targetTypeName);
  const normalizedValueType = normalizeTypeName(valueTypeName);

  if (normalizedTargetType === normalizedValueType || normalizedTargetType === "variant" || normalizedValueType === "variant") {
    return true;
  }

  if (NUMERIC_TYPES.has(normalizedTargetType) && NUMERIC_TYPES.has(normalizedValueType)) {
    return true;
  }

  if (normalizedTargetType === "object" && normalizedValueType === "object") {
    return true;
  }

  if (options.isSetAssignment) {
    if (normalizedTargetType === "object" && isReferenceTypeName(normalizedValueType)) {
      return true;
    }

    if (normalizedValueType === "nothing" && isReferenceTypeName(normalizedTargetType)) {
      return true;
    }
  }

  return false;
}

function normalizeTypeName(typeName: string): string {
  return typeName.replace(/\s+/gu, "").toLowerCase();
}

function inferArithmeticExpressionType(
  symbolTable: SymbolTable,
  symbolTypes: Map<string, InferredSymbolType>,
  line: number,
  expressionText: string
): string | undefined {
  const parts = splitTopLevelExpression(expressionText, (text, index) => {
    const currentCharacter = text[index];

    if (["*", "/", "\\", "^"].includes(currentCharacter)) {
      return 1;
    }

    if (["+", "-"].includes(currentCharacter)) {
      const previousCharacter = getAdjacentNonWhitespaceCharacter(text, index, -1) ?? "";
      return /[A-Za-z0-9_\]")#]/u.test(previousCharacter) ? 1 : 0;
    }

    if (text.slice(index, index + 3).toLowerCase() === "mod" && isTokenBoundary(text, index - 1) && isTokenBoundary(text, index + 3)) {
      return 3;
    }

    return 0;
  });

  if (parts.length <= 1) {
    return undefined;
  }

  const operandTypes = parts.map((part) => inferExpressionType(symbolTable, symbolTypes, line, part.trim()));

  if (operandTypes.some((typeName) => typeName === undefined)) {
    return undefined;
  }

  if (operandTypes.some((typeName) => normalizeTypeName(typeName!) === "variant")) {
    return "Variant";
  }

  if (!operandTypes.every((typeName) => NUMERIC_TYPES.has(normalizeTypeName(typeName!)))) {
    return undefined;
  }

  return expressionText.includes("/") ? "Double" : "Long";
}

function inferComparisonExpressionType(
  symbolTable: SymbolTable,
  symbolTypes: Map<string, InferredSymbolType>,
  line: number,
  expressionText: string
): string | undefined {
  const parts = splitTopLevelExpression(expressionText, (text, index) => {
    const currentSlice = text.slice(index);

    if (currentSlice.startsWith("<=") || currentSlice.startsWith(">=") || currentSlice.startsWith("<>")) {
      return 2;
    }

    if (["<", ">", "="].includes(text[index])) {
      return 1;
    }

    if (currentSlice.toLowerCase().startsWith("is") && isTokenBoundary(text, index - 1) && isTokenBoundary(text, index + 2)) {
      return 2;
    }

    if (currentSlice.toLowerCase().startsWith("like") && isTokenBoundary(text, index - 1) && isTokenBoundary(text, index + 4)) {
      return 4;
    }

    return 0;
  });

  if (parts.length <= 1) {
    return undefined;
  }

  const operandTypes = parts.map((part) => inferExpressionType(symbolTable, symbolTypes, line, part.trim()));

  if (operandTypes.some((typeName) => typeName === undefined)) {
    return undefined;
  }

  return "Boolean";
}

function inferConcatenationExpressionType(
  symbolTable: SymbolTable,
  symbolTypes: Map<string, InferredSymbolType>,
  line: number,
  expressionText: string
): string | undefined {
  const parts = splitTopLevelExpression(expressionText, (text, index) => (text[index] === "&" ? 1 : 0));

  if (parts.length <= 1) {
    return undefined;
  }

  const operandTypes = parts.map((part) => inferExpressionType(symbolTable, symbolTypes, line, part.trim()));

  if (operandTypes.some((typeName) => typeName === undefined)) {
    return undefined;
  }

  return "String";
}

function getAdjacentNonWhitespaceCharacter(text: string, startIndex: number, direction: -1 | 1): string | undefined {
  let index = startIndex + direction;

  while (index >= 0 && index < text.length) {
    if (!/\s/u.test(text[index] ?? "")) {
      return text[index];
    }

    index += direction;
  }

  return undefined;
}

function isReferenceTypeName(typeName: string): boolean {
  return !SCALAR_TYPES.has(typeName);
}

function shouldWarnMissingSetAssignment(
  symbolTable: SymbolTable,
  targetSymbol: SymbolInfo,
  targetTypeName: string,
  valueTypeName: string,
  isSetAssignment: boolean
): boolean {
  if (isSetAssignment) {
    return false;
  }

  const normalizedTargetType = normalizeTypeName(targetTypeName);
  const normalizedValueType = normalizeTypeName(valueTypeName);

  if (!isObjectReferenceType(symbolTable, targetSymbol, normalizedTargetType)) {
    return false;
  }

  if (normalizedValueType === "nothing") {
    return true;
  }

  return isObjectReferenceType(symbolTable, undefined, normalizedValueType);
}

function isObjectReferenceType(symbolTable: SymbolTable, symbol: SymbolInfo | undefined, normalizedTypeName: string): boolean {
  if (!isReferenceTypeName(normalizedTypeName) || normalizedTypeName === "variant") {
    return false;
  }

  if (symbol?.isArray) {
    return false;
  }

  const localUserDefinedType = symbolTable.moduleSymbols.find(
    (entry) => (entry.kind === "enum" || entry.kind === "type") && normalizeTypeName(entry.name) === normalizedTypeName
  );

  return !localUserDefinedType;
}

function splitTopLevelExpression(
  text: string,
  getOperatorLength: (text: string, index: number) => number
): string[] {
  const parts: string[] = [];
  let buffer = "";
  let depth = 0;
  let index = 0;

  while (index < text.length) {
    const currentCharacter = text[index];

    if (currentCharacter === "\"") {
      const nextIndex = skipStringLiteral(text, index);
      buffer += text.slice(index, nextIndex);
      index = nextIndex;
      continue;
    }

    if (currentCharacter === "#") {
      const nextIndex = skipDateLiteral(text, index);
      buffer += text.slice(index, nextIndex);
      index = nextIndex;
      continue;
    }

    if (currentCharacter === "(") {
      depth += 1;
      buffer += currentCharacter;
      index += 1;
      continue;
    }

    if (currentCharacter === ")") {
      depth = Math.max(0, depth - 1);
      buffer += currentCharacter;
      index += 1;
      continue;
    }

    const operatorLength = depth === 0 ? getOperatorLength(text, index) : 0;

    if (operatorLength > 0) {
      parts.push(buffer.trim());
      buffer = "";
      index += operatorLength;
      continue;
    }

    buffer += currentCharacter;
    index += 1;
  }

  parts.push(buffer.trim());
  return parts.filter((part) => part.length > 0);
}

function isTokenBoundary(text: string, index: number): boolean {
  if (index < 0 || index >= text.length) {
    return true;
  }

  return !/[A-Za-z0-9_]/u.test(text[index] ?? "");
}

function skipDateLiteral(text: string, startIndex: number): number {
  let index = startIndex + 1;

  while (index < text.length) {
    if (text[index] === "#") {
      return index + 1;
    }

    index += 1;
  }

  return index;
}

function skipStringLiteral(text: string, startIndex: number): number {
  let index = startIndex + 1;

  while (index < text.length) {
    if (text[index] === "\"" && text[index + 1] === "\"") {
      index += 2;
      continue;
    }

    if (text[index] === "\"") {
      return index + 1;
    }

    index += 1;
  }

  return index;
}

function unwrapParentheses(expressionText: string): string {
  let currentExpression = expressionText;

  while (currentExpression.startsWith("(") && currentExpression.endsWith(")") && parenthesesAreBalanced(currentExpression)) {
    currentExpression = currentExpression.slice(1, -1).trim();
  }

  return currentExpression;
}

function parenthesesAreBalanced(expressionText: string): boolean {
  let depth = 0;

  for (const character of expressionText) {
    if (character === "(") {
      depth += 1;
    } else if (character === ")") {
      depth -= 1;

      if (depth < 0) {
        return false;
      }
    }
  }

  return depth === 0;
}

function inferCreateObjectType(expressionText: string): string | undefined {
  const match = /^(?:CreateObject|GetObject)\s*\(\s*("(?:[^"]|"")*")/iu.exec(expressionText);
  const progId = match?.[1] ? readVbaStringLiteral(match[1]) : undefined;

  return progId ? CREATE_OBJECT_PROGID_TYPES.get(progId.toLowerCase()) : undefined;
}

function readVbaStringLiteral(text: string): string | undefined {
  if (!/^"(?:[^"]|"")*"$/u.test(text)) {
    return undefined;
  }

  return text.slice(1, -1).replace(/""/gu, "\"");
}

