import { VBA_KEYWORDS } from "../lexer/keywords";
import { getSymbolTypeName, areTypesCompatible } from "../inference/inferModuleTypes";
import { splitCodeAndComment } from "../parser/text";
import { resolveSymbolAtPosition } from "../symbol/buildModuleSymbols";
import { normalizeIdentifier } from "../types/helpers";
import { getProcedureStatementReferenceSegments } from "./procedureStatementReferences";
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
      for (const invocation of collectStatementInvocations(result, statement)) {
        const resolvedCallable = resolveCallableAtPosition(invocation.nameRange.start);

        if (!resolvedCallable || invocation.arguments.some((argument) => usesNamedArgument(argument.text))) {
          continue;
        }

        for (let argumentIndex = 0; argumentIndex < invocation.arguments.length; argumentIndex += 1) {
          const parameter = resolvedCallable.callable.parameters[argumentIndex];
          const argument = invocation.arguments[argumentIndex];

          if (!parameter || parameter.direction !== "byRef" || !argument || argument.text.trim().length === 0) {
            continue;
          }

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

function collectStatementInvocations(
  result: AnalysisResult,
  statement: ProcedureDeclarationNode["body"][number]
): Invocation[] {
  if (statement.kind === "callStatement") {
    return [
      {
        arguments: statement.arguments.map((argument) => ({
          range: argument.range,
          text: argument.text
        })),
        name: statement.name,
        nameRange: statement.nameRange
      }
    ];
  }

  if (statement.kind === "constStatement" || statement.kind === "declarationStatement") {
    return [];
  }

  const structuredReferenceSegments = getProcedureStatementReferenceSegments(statement);

  if (structuredReferenceSegments !== undefined) {
    return structuredReferenceSegments
      .filter((segment) => segment.role === "read")
      .flatMap((segment) => collectInvocationsInSourceRange(result, segment.range));
  }

  return collectInvocationsInSourceRange(result, statement.range);
}

function collectInvocationsInSourceRange(result: AnalysisResult, range: SourceRange): Invocation[] {
  const flattenedRange = flattenCodeRange(result, range);

  return collectInvocations(flattenedRange.text, 0, 0).map((invocation) => ({
    arguments: invocation.arguments.map((argument) => ({
      range: mapFlattenedRange(flattenedRange.positions, argument.range),
      text: argument.text
    })),
    name: invocation.name,
    nameRange: mapFlattenedRange(flattenedRange.positions, invocation.nameRange)
  }));
}

function flattenCodeRange(
  result: AnalysisResult,
  range: SourceRange
): { positions: Array<LinePosition | undefined>; text: string } {
  let text = "";
  const positions: Array<LinePosition | undefined> = [];
  let hasOutput = false;

  for (let lineIndex = range.start.line; lineIndex <= range.end.line; lineIndex += 1) {
    const originalLine = result.source.normalizedLines[lineIndex];

    if (originalLine === undefined) {
      continue;
    }

    const sliceStartCharacter = lineIndex === range.start.line ? range.start.character : 0;
    const sliceEndCharacter = lineIndex === range.end.line ? range.end.character : originalLine.length;
    const lineSlice = originalLine.slice(sliceStartCharacter, sliceEndCharacter);
    const { code } = splitCodeAndComment(lineSlice);
    const continued = isLineContinuation(code);
    const codeWithoutContinuation = continued ? code.replace(/\s+_\s*$/, "") : code;
    const trimmedCode = codeWithoutContinuation.trimEnd();
    const leadingTrimLength = hasOutput ? trimmedCode.length - trimmedCode.trimStart().length : 0;
    const emittedText = hasOutput ? trimmedCode.trimStart() : trimmedCode;

    if (hasOutput) {
      text += " ";
      positions.push(undefined);
    }

    for (let index = 0; index < emittedText.length; index += 1) {
      text += emittedText[index];
      positions.push({
        character: sliceStartCharacter + leadingTrimLength + index,
        line: lineIndex
      });
    }

    hasOutput = true;
  }

  return { positions, text };
}

function mapFlattenedRange(
  positions: Array<LinePosition | undefined>,
  range: SourceRange
): SourceRange {
  const startPosition = positions[range.start.character];
  const endPosition = positions[Math.max(range.end.character - 1, range.start.character)];

  if (!startPosition || !endPosition) {
    return range;
  }

  return {
    start: startPosition,
    end: {
      character: endPosition.character + 1,
      line: endPosition.line
    }
  };
}

function findLocalCallable(result: AnalysisResult, position: LinePosition): ResolvedCallable | undefined {
  const identifier = getIdentifierAtPosition(result, position);

  if (!identifier) {
    return undefined;
  }

  const targetSymbol = resolveSymbolAtPosition(result.symbols, identifier, position);

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

  return resolveSymbolAtPosition(result.symbols, identifier, position);
}

function collectInvocations(text: string, line: number, baseCharacter = 0): Invocation[] {
  const invocations = [
    ...collectCallKeywordInvocations(text, line, baseCharacter),
    ...collectBareInvocations(text, line, baseCharacter),
    ...collectParenthesizedInvocations(text, line, baseCharacter)
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

function collectCallKeywordInvocations(text: string, line: number, baseCharacter = 0): Invocation[] {
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
      arguments: splitArgumentsWithRanges(text.slice(openParenIndex + 1, closeParenIndex), line, baseCharacter + openParenIndex + 1),
      name: match[1],
      nameRange: createInlineRange(line, baseCharacter + nameStartCharacter, baseCharacter + nameStartCharacter + match[1].length)
    }
  ];
}

function collectBareInvocations(text: string, line: number, baseCharacter = 0): Invocation[] {
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
      arguments: splitArgumentsWithRanges(match[2], line, baseCharacter + argumentsStartCharacter),
      name: match[1],
      nameRange: createInlineRange(line, baseCharacter + nameStartCharacter, baseCharacter + nameStartCharacter + match[1].length)
    }
  ];
}

function collectParenthesizedInvocations(text: string, line: number, baseCharacter = 0): Invocation[] {
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
      arguments: splitArgumentsWithRanges(text.slice(index + 1, closeParenIndex), line, baseCharacter + index + 1),
      name: identifier.text,
      nameRange: createInlineRange(
        line,
        baseCharacter + identifier.startCharacter,
        baseCharacter + identifier.startCharacter + identifier.text.length
      )
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

function isLineContinuation(code: string): boolean {
  return /\s+_\s*$/.test(code);
}

function usesNamedArgument(argumentText: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_]*\s*:=/u.test(argumentText.trim());
}

function isAssignableSymbol(symbol: SymbolInfo): boolean {
  return symbol.kind === "parameter" || symbol.kind === "variable";
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
