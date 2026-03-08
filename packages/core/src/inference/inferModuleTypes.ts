import { removeStringAndDateLiterals } from "../parser/text";
import { getAccessibleSymbolsAtLine } from "../symbol/buildModuleSymbols";
import { normalizeIdentifier } from "../types/helpers";
import type {
  AnalysisResult,
  Diagnostic,
  InferredSymbolType,
  LinePosition,
  ParseResult,
  ProcedureDeclarationNode,
  SourceRange,
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
  ["cstr", "String"]
]);

const NUMERIC_TYPES = new Set(["byte", "currency", "double", "integer", "long", "longlong", "longptr", "single"]);

export function inferModuleTypes(parseResult: ParseResult, symbolTable: SymbolTable): TypeInferenceResult {
  const diagnostics: Diagnostic[] = [];
  const symbolTypes = new Map<string, InferredSymbolType>();

  seedExplicitTypes(symbolTable, symbolTypes);

  for (const member of parseResult.module.members) {
    if (member.kind !== "procedureDeclaration") {
      continue;
    }

    for (const statement of member.body) {
      if (statement.kind !== "executableStatement" || statement.range.start.line !== statement.range.end.line) {
        continue;
      }

      const assignment = parseSimpleAssignment(statement.text, statement.range);

      if (!assignment) {
        continue;
      }

      const targetSymbol = resolveSymbolAtPosition(symbolTable, statement.range.start.line, assignment.targetName, assignment.targetRange.start);
      const inferredExpressionType = inferExpressionType(symbolTable, symbolTypes, statement.range.start.line, assignment.expressionText);

      if (!targetSymbol || !inferredExpressionType) {
        continue;
      }

      const targetTypeName = getSymbolTypeNameFromMap(symbolTypes, targetSymbol) ?? targetSymbol.typeName;

      if (targetTypeName) {
        if (!areTypesCompatible(targetTypeName, inferredExpressionType)) {
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

function seedExplicitTypes(symbolTable: SymbolTable, sink: Map<string, InferredSymbolType>): void {
  for (const symbol of symbolTable.allSymbols) {
    if (symbol.typeName) {
      setSymbolType(sink, symbol, symbol.typeName, "explicit");
    }
  }
}

function parseSimpleAssignment(
  text: string,
  statementRange: SourceRange
): { expressionRange: SourceRange; expressionText: string; targetName: string; targetRange: SourceRange } | undefined {
  const match = /^\s*([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)\s*=\s*(.+?)\s*$/i.exec(text);

  if (!match || /=/.test(removeStringAndDateLiterals(match[2]))) {
    return undefined;
  }

  const targetName = match[1];
  const targetStartCharacter = match.index ?? 0;
  const equalsIndex = text.indexOf("=", targetStartCharacter + targetName.length);

  if (equalsIndex < 0) {
    return undefined;
  }

  const expressionStart = equalsIndex + 1 + (match[2].length - match[2].trimStart().length);
  const expressionText = match[2].trim();

  return {
    expressionRange: createInlineRange(statementRange.start.line, expressionStart, expressionStart + expressionText.length),
    expressionText,
    targetName: targetName.replace(/[$%&!#@]$/, ""),
    targetRange: createInlineRange(statementRange.start.line, targetStartCharacter, targetStartCharacter + targetName.length)
  };
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

  if (/^(?:True|False)$/iu.test(normalizedExpression)) {
    return "Boolean";
  }

  if (/^[+-]?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?$/u.test(normalizedExpression)) {
    return /[.Ee]/u.test(normalizedExpression) ? "Double" : "Long";
  }

  const callMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)\s*\((.*)\)$/iu.exec(normalizedExpression);

  if (callMatch) {
    const castType = CAST_FUNCTION_TYPES.get(normalizeIdentifier(callMatch[1]));

    if (castType) {
      return castType;
    }

    const symbol = resolveSymbolAtPosition(symbolTable, line, callMatch[1], { character: 0, line });
    return symbol ? getResolvedTypeName(symbolTypes, symbol) : undefined;
  }

  const identifierMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)$/iu.exec(normalizedExpression);

  if (!identifierMatch) {
    return undefined;
  }

  const symbol = resolveSymbolAtPosition(symbolTable, line, identifierMatch[1], { character: 0, line });
  return symbol ? getResolvedTypeName(symbolTypes, symbol) : undefined;
}

function resolveSymbolAtPosition(symbolTable: SymbolTable, line: number, identifier: string, position: LinePosition): SymbolInfo | undefined {
  const matchingSymbols = getAccessibleSymbolsAtLine(symbolTable, line).filter(
    (symbol) => symbol.normalizedName === normalizeIdentifier(identifier)
  );

  if (matchingSymbols.length === 0) {
    return undefined;
  }

  const declarationMatches = matchingSymbols.filter((symbol) => symbol.kind !== "module" && positionWithinRange(position, symbol.selectionRange));

  if (declarationMatches.length > 0) {
    return declarationMatches.find((symbol) => symbol.scope === "module") ?? declarationMatches[0];
  }

  return matchingSymbols.find((symbol) => symbol.scope === "procedure") ?? matchingSymbols.find((symbol) => symbol.kind !== "module") ?? matchingSymbols[0];
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

function areTypesCompatible(targetTypeName: string, valueTypeName: string): boolean {
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

  return false;
}

function normalizeTypeName(typeName: string): string {
  return typeName.replace(/\s+/gu, "").toLowerCase();
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

function positionWithinRange(position: LinePosition, range: SourceRange): boolean {
  if (position.line < range.start.line || position.line > range.end.line) {
    return false;
  }

  if (position.line === range.start.line && position.character < range.start.character) {
    return false;
  }

  if (position.line === range.end.line && position.character > range.end.character) {
    return false;
  }

  return true;
}

function createInlineRange(line: number, startCharacter: number, endCharacter: number): SourceRange {
  return {
    start: {
      character: startCharacter,
      line
    },
    end: {
      character: endCharacter,
      line
    }
  };
}
