import { VBA_KEYWORDS } from "../lexer/keywords";
import { getSymbolTypeName, areTypesCompatible } from "../inference/inferModuleTypes";
import { splitCodeAndComment } from "../parser/text";
import { getAccessibleSymbolsAtLine } from "../symbol/buildModuleSymbols";
import { normalizeIdentifier } from "../types/helpers";
import type {
  AnalysisResult,
  DeclareStatementNode,
  Diagnostic,
  LinePosition,
  ProcedureDeclarationNode,
  SourceRange,
  SymbolInfo
} from "../types/model";

type CallableMember = DeclareStatementNode | ProcedureDeclarationNode;

interface InvocationArgument {
  range: SourceRange;
  text: string;
}

interface Invocation {
  arguments: InvocationArgument[];
  name: string;
  nameRange: SourceRange;
}

export interface ResolvedCallable {
  callable: CallableMember;
  symbol: SymbolInfo;
}

export function collectByRefArgumentDiagnostics(
  result: AnalysisResult,
  resolveCallableAtPosition: (position: LinePosition) => ResolvedCallable | undefined = (position) => findLocalCallable(result, position)
): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];

  for (const member of result.module.members) {
    if (member.kind !== "procedureDeclaration") {
      continue;
    }

    for (const statement of member.body) {
      if (statement.kind !== "executableStatement" || statement.range.start.line !== statement.range.end.line) {
        continue;
      }

      const originalLine = result.source.normalizedLines[statement.range.start.line] ?? statement.text;
      const { code } = splitCodeAndComment(originalLine);

      for (const invocation of collectInvocations(code, statement.range.start.line)) {
        const resolvedCallable = resolveCallableAtPosition(invocation.nameRange.start);

        if (!resolvedCallable || invocation.arguments.some((argument) => usesNamedArgument(argument.text))) {
          continue;
        }

        for (let argumentIndex = 0; argumentIndex < invocation.arguments.length; argumentIndex += 1) {
          const parameter = resolvedCallable.callable.parameters[argumentIndex];

          if (!parameter || parameter.direction !== "byRef") {
            continue;
          }

          const argument = invocation.arguments[argumentIndex];
          const assignableArgument = resolveAssignableArgument(result, argument);

          if (!assignableArgument) {
            diagnostics.push({
              code: "byref-argument-risk",
              message: `ByRef parameter '${parameter.name}' in ${resolvedCallable.symbol.name} receives an expression. Introduce a temporary variable before the call.`,
              range: argument.range,
              severity: "warning"
            });
            continue;
          }

          const parameterTypeName = parameter.typeName;
          const argumentTypeName = getSymbolTypeName(result, assignableArgument.symbol);

          if (!parameterTypeName || !argumentTypeName) {
            continue;
          }

          if (!areTypesCompatible(parameterTypeName, argumentTypeName, { isSetAssignment: true })) {
            diagnostics.push({
              code: "byref-argument-type-mismatch",
              message: `ByRef parameter '${parameter.name}' in ${resolvedCallable.symbol.name} expects ${parameterTypeName} but receives ${argumentTypeName}. VBA may raise a ByRef argument type mismatch.`,
              range: argument.range,
              severity: "warning"
            });
          }
        }
      }
    }
  }

  return diagnostics;
}

function findLocalCallable(result: AnalysisResult, position: LinePosition): ResolvedCallable | undefined {
  const identifier = getIdentifierAtPosition(result, position);

  if (!identifier) {
    return undefined;
  }

  const accessibleSymbols = getAccessibleSymbolsAtLine(result.symbols, position.line);
  const matchingSymbols = accessibleSymbols.filter((symbol) => symbol.normalizedName === normalizeIdentifier(identifier));

  if (matchingSymbols.length === 0) {
    return undefined;
  }

  const declarationMatches = matchingSymbols.filter(
    (symbol) => symbol.kind !== "module" && isPositionWithinRange(position, symbol.selectionRange)
  );
  const targetSymbol =
    declarationMatches.find((symbol) => symbol.scope === "module") ??
    declarationMatches[0] ??
    matchingSymbols.find((symbol) => symbol.scope === "procedure") ??
    matchingSymbols.find((symbol) => symbol.kind !== "module") ??
    matchingSymbols[0];

  if (!targetSymbol) {
    return undefined;
  }

  const callable = findCallableMember(result, targetSymbol);
  return callable ? { callable, symbol: targetSymbol } : undefined;
}

function findCallableMember(result: AnalysisResult, symbol: SymbolInfo): CallableMember | undefined {
  return result.module.members.find((member): member is CallableMember => {
    if ((member.kind !== "declareStatement" && member.kind !== "procedureDeclaration") || member.name !== symbol.name) {
      return false;
    }

    const selectionRange = member.kind === "procedureDeclaration" ? member.headerRange : member.range;
    return rangesEqual(selectionRange, symbol.selectionRange);
  });
}

function resolveAssignableArgument(
  result: AnalysisResult,
  argument: InvocationArgument
): { symbol: SymbolInfo } | undefined {
  const trimmedText = argument.text.trim();
  const identifierMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)$/u.exec(trimmedText);

  if (identifierMatch) {
    const symbol = resolveArgumentSymbol(result, argument.range.start);

    if (symbol && isAssignableSymbol(symbol)) {
      return { symbol };
    }

    return undefined;
  }

  const arrayReferenceMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)\s*\((.+)\)$/u.exec(trimmedText);

  if (!arrayReferenceMatch) {
    return undefined;
  }

  const symbol = resolveArgumentSymbol(result, argument.range.start);

  if (!symbol || !isAssignableSymbol(symbol) || !symbol.isArray) {
    return undefined;
  }

  return { symbol };
}

function resolveArgumentSymbol(result: AnalysisResult, position: LinePosition): SymbolInfo | undefined {
  const identifier = getIdentifierAtPosition(result, position);

  if (!identifier) {
    return undefined;
  }

  const accessibleSymbols = getAccessibleSymbolsAtLine(result.symbols, position.line);
  const matchingSymbols = accessibleSymbols.filter((symbol) => symbol.normalizedName === normalizeIdentifier(identifier));

  if (matchingSymbols.length === 0) {
    return undefined;
  }

  const declarationMatches = matchingSymbols.filter(
    (symbol) => symbol.kind !== "module" && isPositionWithinRange(position, symbol.selectionRange)
  );

  return (
    declarationMatches.find((symbol) => symbol.scope === "module") ??
    declarationMatches[0] ??
    matchingSymbols.find((symbol) => symbol.scope === "procedure") ??
    matchingSymbols.find((symbol) => symbol.kind !== "module") ??
    matchingSymbols[0]
  );
}

function collectInvocations(text: string, line: number): Invocation[] {
  const invocations = [
    ...collectCallKeywordInvocations(text, line),
    ...collectBareInvocations(text, line),
    ...collectParenthesizedInvocations(text, line)
  ];
  const uniqueInvocations = new Map<string, Invocation>();

  for (const invocation of invocations) {
    const key = `${invocation.nameRange.start.character}:${invocation.nameRange.end.character}:${invocation.arguments.length}`;

    if (!uniqueInvocations.has(key)) {
      uniqueInvocations.set(key, invocation);
    }
  }

  return [...uniqueInvocations.values()];
}

function collectCallKeywordInvocations(text: string, line: number): Invocation[] {
  const match = /^\s*Call\s+([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)\s*\((.*)\)\s*$/iu.exec(text);

  if (!match?.[1]) {
    return [];
  }

  const callPrefixLength = (/^\s*Call\s+/iu.exec(text)?.[0].length ?? 0);
  const nameStartCharacter = callPrefixLength;
  const openParenIndex = text.indexOf("(", nameStartCharacter);
  const closeParenIndex = findMatchingCloseParen(text, openParenIndex);

  if (openParenIndex < 0 || closeParenIndex < 0) {
    return [];
  }

  return [
    {
      arguments: splitArgumentsWithRanges(text.slice(openParenIndex + 1, closeParenIndex), line, openParenIndex + 1),
      name: match[1],
      nameRange: createInlineRange(line, nameStartCharacter, nameStartCharacter + match[1].length)
    }
  ];
}

function collectBareInvocations(text: string, line: number): Invocation[] {
  if (findAssignmentOperatorIndex(text) >= 0) {
    return [];
  }

  const match = /^\s*([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)(?:\s+(.*\S))?\s*$/u.exec(text);

  if (!match?.[1] || !match[2]) {
    return [];
  }

  if (VBA_KEYWORDS.has(normalizeIdentifier(match[1]))) {
    return [];
  }

  const leadingWhitespace = (/^\s*/u.exec(text)?.[0].length ?? 0);
  const nameStartCharacter = leadingWhitespace;
  const argumentsStartCharacter = nameStartCharacter + match[1].length + (/\s+/u.exec(text.slice(nameStartCharacter + match[1].length))?.[0].length ?? 0);

  return [
    {
      arguments: splitArgumentsWithRanges(match[2], line, argumentsStartCharacter),
      name: match[1],
      nameRange: createInlineRange(line, nameStartCharacter, nameStartCharacter + match[1].length)
    }
  ];
}

function collectParenthesizedInvocations(text: string, line: number): Invocation[] {
  const invocations: Invocation[] = [];
  let index = 0;

  while (index < text.length) {
    const currentCharacter = text[index];

    if (currentCharacter === "\"") {
      index = skipStringLiteral(text, index);
      continue;
    }

    if (currentCharacter === "#") {
      index = skipDateLiteral(text, index);
      continue;
    }

    if (currentCharacter !== "(") {
      index += 1;
      continue;
    }

    const identifier = getIdentifierBeforeOpenParen(text, index);
    const closeParenIndex = findMatchingCloseParen(text, index);

    if (!identifier || closeParenIndex < 0) {
      index += 1;
      continue;
    }

    invocations.push({
      arguments: splitArgumentsWithRanges(text.slice(index + 1, closeParenIndex), line, index + 1),
      name: identifier.text,
      nameRange: createInlineRange(line, identifier.startCharacter, identifier.startCharacter + identifier.text.length)
    });
    index += 1;
  }

  return invocations;
}

function splitArgumentsWithRanges(text: string, line: number, baseCharacter: number): InvocationArgument[] {
  const parsedArguments: InvocationArgument[] = [];
  let startIndex = 0;
  let depth = 0;
  let index = 0;

  while (index < text.length) {
    const currentCharacter = text[index];

    if (currentCharacter === "\"") {
      index = skipStringLiteral(text, index);
      continue;
    }

    if (currentCharacter === "#") {
      index = skipDateLiteral(text, index);
      continue;
    }

    if (currentCharacter === "(") {
      depth += 1;
      index += 1;
      continue;
    }

    if (currentCharacter === ")") {
      depth = Math.max(0, depth - 1);
      index += 1;
      continue;
    }

    if (currentCharacter === "," && depth === 0) {
      pushArgument(startIndex, index);
      startIndex = index + 1;
    }

    index += 1;
  }

  pushArgument(startIndex, text.length);
  return parsedArguments;

  function pushArgument(start: number, end: number): void {
    const rawText = text.slice(start, end);
    const leadingWhitespace = rawText.length - rawText.trimStart().length;
    const trailingWhitespace = rawText.length - rawText.trimEnd().length;
    const trimmedText = rawText.trim();

    if (trimmedText.length === 0) {
      return;
    }

    parsedArguments.push({
      range: createInlineRange(line, baseCharacter + start + leadingWhitespace, baseCharacter + end - trailingWhitespace),
      text: trimmedText
    });
  }
}

function getIdentifierAtPosition(result: AnalysisResult, position: LinePosition): string | undefined {
  const line = result.source.originalText.replace(/\r\n?/gu, "\n").split("\n")[position.line];

  if (line === undefined) {
    return undefined;
  }

  for (const match of line.matchAll(/[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?/gu)) {
    const start = match.index ?? 0;
    const end = start + match[0].length;

    if (position.character >= start && position.character <= end) {
      return match[0];
    }
  }

  return undefined;
}

function getIdentifierBeforeOpenParen(
  text: string,
  openParenIndex: number
): { startCharacter: number; text: string } | undefined {
  let endIndex = openParenIndex - 1;

  while (endIndex >= 0 && /\s/u.test(text[endIndex] ?? "")) {
    endIndex -= 1;
  }

  if (endIndex < 0) {
    return undefined;
  }

  let startIndex = endIndex;

  while (startIndex >= 0 && /[A-Za-z0-9_$%&!#@]/u.test(text[startIndex] ?? "")) {
    startIndex -= 1;
  }

  startIndex += 1;

  const identifier = text.slice(startIndex, endIndex + 1);

  if (!/^[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?$/u.test(identifier)) {
    return undefined;
  }

  return {
    startCharacter: startIndex,
    text: identifier
  };
}

function findMatchingCloseParen(text: string, openParenIndex: number): number {
  let depth = 0;
  let index = openParenIndex;

  while (index < text.length) {
    const currentCharacter = text[index];

    if (currentCharacter === "\"") {
      index = skipStringLiteral(text, index);
      continue;
    }

    if (currentCharacter === "#") {
      index = skipDateLiteral(text, index);
      continue;
    }

    if (currentCharacter === "(") {
      depth += 1;
    } else if (currentCharacter === ")") {
      depth -= 1;

      if (depth === 0) {
        return index;
      }
    }

    index += 1;
  }

  return -1;
}

function findAssignmentOperatorIndex(text: string): number {
  let depth = 0;
  let index = 0;

  while (index < text.length) {
    const currentCharacter = text[index];

    if (currentCharacter === "\"") {
      index = skipStringLiteral(text, index);
      continue;
    }

    if (currentCharacter === "#") {
      index = skipDateLiteral(text, index);
      continue;
    }

    if (currentCharacter === "(") {
      depth += 1;
      index += 1;
      continue;
    }

    if (currentCharacter === ")") {
      depth = Math.max(0, depth - 1);
      index += 1;
      continue;
    }

    if (depth === 0 && currentCharacter === "=") {
      const previousNonWhitespaceCharacter = getAdjacentNonWhitespaceCharacter(text, index, -1);

      if (previousNonWhitespaceCharacter !== "<" && previousNonWhitespaceCharacter !== ">") {
        return index;
      }
    }

    index += 1;
  }

  return -1;
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

function usesNamedArgument(argumentText: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_]*\s*:=/u.test(argumentText.trim());
}

function isAssignableSymbol(symbol: SymbolInfo): boolean {
  return symbol.kind === "parameter" || symbol.kind === "variable";
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

function rangesEqual(left: SourceRange, right: SourceRange): boolean {
  return (
    left.start.line === right.start.line &&
    left.start.character === right.start.character &&
    left.end.line === right.end.line &&
    left.end.character === right.end.character
  );
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
