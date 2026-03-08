import { BUILTIN_IDENTIFIERS, VBA_KEYWORDS } from "../lexer/keywords";
import { parseModule } from "../parser/parseModule";
import { extractIdentifierAtPosition, removeStringAndDateLiterals, splitCodeAndComment } from "../parser/text";
import { normalizeIdentifier } from "../types/helpers";
import { AnalysisResult, AnalyzeModuleOptions, Diagnostic, LinePosition, OutlineSymbol, ParseResult, SymbolInfo, SymbolTable } from "../types/model";
import { buildModuleSymbols, getAccessibleSymbolsAtLine } from "../symbol/buildModuleSymbols";

export function analyzeModule(text: string, options: AnalyzeModuleOptions = {}): AnalysisResult {
  const parseResult = parseModule(text, options);
  const symbols = buildModuleSymbols(parseResult);
  const diagnostics = [...parseResult.diagnostics, ...collectUndeclaredVariableDiagnostics(parseResult, symbols)];

  return {
    ...parseResult,
    diagnostics,
    symbols
  };
}

export function findDefinition(result: AnalysisResult | ParseResult & { symbols: AnalysisResult["symbols"] }, position: LinePosition): SymbolInfo | undefined {
  const identifier = extractIdentifierAtPosition(result.source.originalText, position);

  if (!identifier) {
    return undefined;
  }

  const normalizedIdentifier = normalizeIdentifier(identifier);
  const accessibleSymbols = getAccessibleSymbolsAtLine(result.symbols, position.line);
  const matchingSymbols = accessibleSymbols.filter((symbol) => symbol.normalizedName === normalizedIdentifier);

  if (matchingSymbols.length === 0) {
    return undefined;
  }

  const declarationMatches = matchingSymbols.filter(
    (symbol) => symbol.kind !== "module" && isPositionWithinRange(position, symbol.selectionRange)
  );

  if (declarationMatches.length > 0) {
    return declarationMatches.find((symbol) => symbol.scope === "module") ?? declarationMatches[0];
  }

  return matchingSymbols.find((symbol) => symbol.scope === "procedure") ?? matchingSymbols.find((symbol) => symbol.kind !== "module") ?? matchingSymbols[0];
}

export function getCompletionSymbols(
  result: AnalysisResult | ParseResult & { symbols: AnalysisResult["symbols"] },
  position: LinePosition
): SymbolInfo[] {
  return getAccessibleSymbolsAtLine(result.symbols, position.line);
}

export function getDocumentOutline(result: AnalysisResult | ParseResult): OutlineSymbol[] {
  const children: OutlineSymbol[] = [];

  for (const member of result.module.members) {
    switch (member.kind) {
      case "constDeclaration":
        children.push({
          kind: "constant",
          name: member.name,
          range: member.range,
          selectionRange: member.range
        });
        break;
      case "declareStatement":
        children.push({
          kind: "declare",
          name: member.name,
          range: member.range,
          selectionRange: member.range
        });
        break;
      case "enumDeclaration":
        children.push({
          children: member.members.map((enumMember) => ({
            kind: "enumMember",
            name: enumMember.name,
            range: enumMember.range,
            selectionRange: enumMember.range
          })),
          kind: "enum",
          name: member.name,
          range: member.range,
          selectionRange: member.range
        });
        break;
      case "procedureDeclaration":
        children.push({
          kind: "procedure",
          name: member.name,
          range: member.range,
          selectionRange: member.headerRange
        });
        break;
      case "typeDeclaration":
        children.push({
          children: member.members.map((typeMember) => ({
            kind: "typeMember",
            name: typeMember.name,
            range: typeMember.range,
            selectionRange: typeMember.range
          })),
          kind: "type",
          name: member.name,
          range: member.range,
          selectionRange: member.range
        });
        break;
      case "variableDeclaration":
        for (const declarator of member.declarators) {
          children.push({
            kind: "variable",
            name: declarator.name,
            range: declarator.range,
            selectionRange: declarator.range
          });
        }
        break;
      default:
        break;
    }
  }

  return [
    {
      children,
      kind: "module",
      name: result.module.name,
      range: result.module.range,
      selectionRange: result.module.range
    }
  ];
}

function collectUndeclaredVariableDiagnostics(parseResult: ParseResult, symbolTable: SymbolTable): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];

  for (const member of parseResult.module.members) {
    if (member.kind !== "procedureDeclaration") {
      continue;
    }

    for (const statement of member.body) {
      if (statement.kind !== "executableStatement") {
        continue;
      }

      const accessibleNames = new Set(
        getAccessibleSymbolsAtLine(symbolTable, statement.range.start.line).map((symbol) => symbol.normalizedName)
      );
      const { code } = splitCodeAndComment(statement.text);
      const scrubbedText = removeStringAndDateLiterals(code);

      for (const match of scrubbedText.matchAll(/[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?/g)) {
        const rawIdentifier = match[0];
        const normalizedIdentifier = normalizeIdentifier(rawIdentifier);
        const startIndex = match.index ?? 0;
        const previousCharacter = scrubbedText[startIndex - 1] ?? "";
        const nextCharacter = scrubbedText[startIndex + rawIdentifier.length] ?? "";

        if (previousCharacter === "." || previousCharacter === ":") {
          continue;
        }

        if (VBA_KEYWORDS.has(normalizedIdentifier) || BUILTIN_IDENTIFIERS.has(normalizedIdentifier)) {
          continue;
        }

        if (accessibleNames.has(normalizedIdentifier)) {
          continue;
        }

        if (nextCharacter === ":" && startIndex === 0) {
          continue;
        }

        diagnostics.push({
          code: "undeclared-variable",
          message: `Undeclared identifier '${rawIdentifier.replace(/[$%&!#@]$/, "")}'.`,
          range: {
            start: {
              character: statement.range.start.character + startIndex,
              line: statement.range.start.line
            },
            end: {
              character: statement.range.start.character + startIndex + rawIdentifier.length,
              line: statement.range.start.line
            }
          },
          severity: "error"
        });
      }
    }
  }

  return diagnostics;
}

function isPositionWithinRange(position: LinePosition, range: SymbolInfo["selectionRange"]): boolean {
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
