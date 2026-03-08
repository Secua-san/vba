import { buildLogicalLines, removeStringAndDateLiterals, splitCodeAndComment } from "../parser/text";
import { createSourceDocument } from "../types/helpers";

type BlockKind = "case" | "directiveIf" | "do" | "enum" | "for" | "if" | "procedure" | "select" | "type" | "while" | "with";
type LineKind =
  | "blankOrComment"
  | "caseBranch"
  | "directiveElseBranch"
  | "directiveEndIf"
  | "directiveIfStart"
  | "doStart"
  | "elseBranch"
  | "endEnum"
  | "endIf"
  | "endSelect"
  | "endType"
  | "endWith"
  | "forStart"
  | "ifStart"
  | "loop"
  | "next"
  | "other"
  | "procedureHeader"
  | "procedureTerminator"
  | "selectStart"
  | "typeStart"
  | "enumStart"
  | "wend"
  | "whileStart"
  | "withStart";

export interface FormatModuleIndentationOptions {
  continuationIndentLevels?: number;
  fileName?: string;
  indentSize?: number;
  insertSpaces?: boolean;
}

interface ContinuationContext {
  kind: "argumentList" | "assignment" | "generic" | "methodChain";
}

type DeclarationAlignmentKind = "const" | "declare" | "dim";

interface AlignableDeclarationLine {
  comment?: string;
  head: string;
  indent: string;
  kind: DeclarationAlignmentKind;
  tail?: string;
  typeName?: string;
  value?: string;
}

export function formatModuleIndentation(text: string, options: FormatModuleIndentationOptions = {}): string {
  const lineEnding = text.includes("\r\n") ? "\r\n" : "\n";
  const source = createSourceDocument(text, { fileName: options.fileName });

  if (source.normalizedLines.length === 0) {
    return text;
  }

  const indentUnit = options.insertSpaces === false ? "\t" : " ".repeat(options.indentSize ?? 4);
  const continuationIndentLevels = Math.max(0, options.continuationIndentLevels ?? 1);
  const normalizedCodeText = normalizeBlockLayout(source.normalizedLines.join(lineEnding));
  const declarationAlignedText = normalizeDeclarationAlignment(normalizedCodeText);
  const formattingSource = createSourceDocument(declarationAlignedText, { fileName: options.fileName });
  const formattedLines = [...formattingSource.normalizedLines];
  const stack: BlockKind[] = [];

  for (const logicalLine of buildLogicalLines(formattingSource)) {
    const kind = classifyLineKind(logicalLine.codeText.trim());
    applyPreIndentation(kind, stack);

    const baseIndentLevel = getBaseIndentLevel(kind, stack);
    const physicalLines = formattingSource.normalizedLines.slice(logicalLine.startLine, logicalLine.endLine + 1);
    const continuationContext = getContinuationContext(physicalLines);

    for (let lineIndex = logicalLine.startLine; lineIndex <= logicalLine.endLine; lineIndex += 1) {
      const originalLine = formattingSource.normalizedLines[lineIndex] ?? "";
      const trimmedLine = trimLogicalLinePreservingContinuation(originalLine, lineIndex < logicalLine.endLine);

      if (trimmedLine.length === 0) {
        formattedLines[lineIndex] = "";
        continue;
      }

      const indentLevel =
        lineIndex === logicalLine.startLine
          ? getIndentLevelForFirstLine(trimmedLine, baseIndentLevel)
          : getIndentLevelForContinuationLine(trimmedLine, baseIndentLevel, continuationIndentLevels, continuationContext);
      formattedLines[lineIndex] = indentUnit.repeat(Math.max(0, indentLevel)) + trimmedLine;
    }

    applyPostIndentation(kind, stack);
  }

  const codeStartLine = source.lineMap[0] ?? source.originalLines.length;
  const nextLines = [...source.originalLines.slice(0, codeStartLine), ...formattedLines];
  const formattedText = nextLines.join(lineEnding);

  return formattedText === text ? text : formattedText;
}

function classifyLineKind(trimmedCode: string): LineKind {
  if (trimmedCode.length === 0) {
    return "blankOrComment";
  }

  if (/^#If\b.*\bThen\b/i.test(trimmedCode)) {
    return "directiveIfStart";
  }

  if (/^#Else(?:If\b.*\bThen\b)?/i.test(trimmedCode)) {
    return "directiveElseBranch";
  }

  if (/^#End\s+If\b/i.test(trimmedCode)) {
    return "directiveEndIf";
  }

  if (/^(?:(?:Public|Private|Friend)\s+)?(?:(?:Static)\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\b/i.test(trimmedCode)) {
    return "procedureHeader";
  }

  if (/^End\s+(?:Sub|Function|Property)\b/i.test(trimmedCode)) {
    return "procedureTerminator";
  }

  if (/^If\b.*\bThen\s*$/i.test(trimmedCode) && !/:/.test(trimmedCode)) {
    return "ifStart";
  }

  if (/^(?:ElseIf\b.*\bThen\b|Else\b)/i.test(trimmedCode)) {
    return "elseBranch";
  }

  if (/^End\s+If\b/i.test(trimmedCode)) {
    return "endIf";
  }

  if (/^Select\s+Case\b/i.test(trimmedCode)) {
    return "selectStart";
  }

  if (/^Case(?:\s+Else|\b)/i.test(trimmedCode)) {
    return "caseBranch";
  }

  if (/^End\s+Select\b/i.test(trimmedCode)) {
    return "endSelect";
  }

  if (/^For\b/i.test(trimmedCode)) {
    return "forStart";
  }

  if (/^Next\b/i.test(trimmedCode)) {
    return "next";
  }

  if (/^Do\b/i.test(trimmedCode)) {
    return "doStart";
  }

  if (/^Loop\b/i.test(trimmedCode)) {
    return "loop";
  }

  if (/^While\b/i.test(trimmedCode)) {
    return "whileStart";
  }

  if (/^Wend\b/i.test(trimmedCode)) {
    return "wend";
  }

  if (/^With\b/i.test(trimmedCode)) {
    return "withStart";
  }

  if (/^End\s+With\b/i.test(trimmedCode)) {
    return "endWith";
  }

  if (/^(?:(?:Public|Private)\s+)?Type\b/i.test(trimmedCode)) {
    return "typeStart";
  }

  if (/^End\s+Type\b/i.test(trimmedCode)) {
    return "endType";
  }

  if (/^(?:(?:Public|Private)\s+)?Enum\b/i.test(trimmedCode)) {
    return "enumStart";
  }

  if (/^End\s+Enum\b/i.test(trimmedCode)) {
    return "endEnum";
  }

  return "other";
}

function applyPreIndentation(kind: LineKind, stack: BlockKind[]): void {
  switch (kind) {
    case "caseBranch":
      popTop(stack, "case");
      break;
    case "directiveEndIf":
      popTop(stack, "directiveIf");
      break;
    case "endEnum":
      popTop(stack, "enum");
      break;
    case "endIf":
      popTop(stack, "if");
      break;
    case "endSelect":
      popTop(stack, "case");
      popTop(stack, "select");
      break;
    case "endType":
      popTop(stack, "type");
      break;
    case "endWith":
      popTop(stack, "with");
      break;
    case "loop":
      popTop(stack, "do");
      break;
    case "next":
      popTop(stack, "for");
      break;
    case "procedureTerminator":
      popTop(stack, "procedure");
      break;
    case "wend":
      popTop(stack, "while");
      break;
    default:
      break;
  }
}

function getBaseIndentLevel(kind: LineKind, stack: BlockKind[]): number {
  switch (kind) {
    case "directiveElseBranch":
      return stack.at(-1) === "directiveIf" ? Math.max(0, stack.length - 1) : stack.length;
    case "elseBranch":
      return stack.at(-1) === "if" ? Math.max(0, stack.length - 1) : stack.length;
    default:
      return stack.length;
  }
}

function applyPostIndentation(kind: LineKind, stack: BlockKind[]): void {
  switch (kind) {
    case "caseBranch":
      stack.push("case");
      break;
    case "directiveIfStart":
      stack.push("directiveIf");
      break;
    case "doStart":
      stack.push("do");
      break;
    case "enumStart":
      stack.push("enum");
      break;
    case "forStart":
      stack.push("for");
      break;
    case "ifStart":
      stack.push("if");
      break;
    case "procedureHeader":
      stack.push("procedure");
      break;
    case "selectStart":
      stack.push("select");
      break;
    case "typeStart":
      stack.push("type");
      break;
    case "whileStart":
      stack.push("while");
      break;
    case "withStart":
      stack.push("with");
      break;
    default:
      break;
  }
}

function normalizeBlockLayout(text: string): string {
  return text
    .split("\n")
    .flatMap(expandCompressedBlockLine)
    .join("\n");
}

function normalizeDeclarationAlignment(text: string): string {
  const lines = text.split("\n");
  const normalizedLines: string[] = [];
  let group: AlignableDeclarationLine[] = [];

  const flushGroup = (): void => {
    if (group.length === 0) {
      return;
    }

    normalizedLines.push(...formatDeclarationGroup(group));
    group = [];
  };

  for (const line of lines) {
    const declaration = parseAlignableDeclarationLine(line);

    if (!declaration) {
      flushGroup();
      normalizedLines.push(line);
      continue;
    }

    if (group.length === 0 || canJoinDeclarationGroup(group[0], declaration)) {
      group.push(declaration);
      continue;
    }

    flushGroup();
    group.push(declaration);
  }

  flushGroup();
  return normalizedLines.join("\n");
}

function expandCompressedBlockLine(line: string): string[] {
  const { code, comment } = splitCodeAndComment(line);

  if (!code.includes(":") || /\s+_\s*$/.test(code)) {
    return [line];
  }

  const rawSegments = splitColonAware(code);

  if (rawSegments.length < 2) {
    return [line];
  }

  const segments = rawSegments.map((segment) => segment.trim()).filter((segment) => segment.length > 0);

  if (segments.length < 2 || isLabelSegment(segments[0])) {
    return [line];
  }

  const kinds = segments.map((segment) => classifyLineKind(segment));
  const expandedLines: string[] = [];
  let statementBuffer = "";
  let changed = false;

  for (let index = 0; index < segments.length; index += 1) {
    const segment = segments[index];

    if (shouldOwnLine(index, kinds)) {
      if (statementBuffer.length > 0) {
        expandedLines.push(statementBuffer);
        statementBuffer = "";
      }

      expandedLines.push(segment);
      changed = true;
      continue;
    }

    statementBuffer = statementBuffer.length > 0 ? `${statementBuffer}: ${segment}` : segment;
  }

  if (statementBuffer.length > 0) {
    expandedLines.push(statementBuffer);
  }

  if (!changed || expandedLines.length < 2) {
    return [line];
  }

  if (comment) {
    const lastLineIndex = expandedLines.length - 1;
    expandedLines[lastLineIndex] = expandedLines[lastLineIndex].length > 0 ? `${expandedLines[lastLineIndex]} ${comment.trimStart()}` : comment.trimStart();
  }

  return expandedLines;
}

function splitColonAware(text: string): string[] {
  const segments: string[] = [];
  let buffer = "";
  let index = 0;

  while (index < text.length) {
    const currentCharacter = text[index];

    if (currentCharacter === "\"") {
      buffer += currentCharacter;
      index += 1;

      while (index < text.length) {
        buffer += text[index];

        if (text[index] === "\"" && text[index + 1] === "\"") {
          buffer += text[index + 1];
          index += 2;
          continue;
        }

        if (text[index] === "\"") {
          index += 1;
          break;
        }

        index += 1;
      }

      continue;
    }

    if (currentCharacter === "#" && !/^#(?:If\b|Else(?:If\b)?|End\s+If\b)/i.test(text.slice(index))) {
      buffer += currentCharacter;
      index += 1;

      while (index < text.length) {
        buffer += text[index];

        if (text[index] === "#") {
          index += 1;
          break;
        }

        index += 1;
      }

      continue;
    }

    if (currentCharacter === ":") {
      segments.push(buffer);
      buffer = "";
      index += 1;
      continue;
    }

    buffer += currentCharacter;
    index += 1;
  }

  segments.push(buffer);
  return segments;
}

function shouldOwnLine(index: number, kinds: LineKind[]): boolean {
  const kind = kinds[index];

  if (kind === "blankOrComment" || kind === "other") {
    return false;
  }

  if (kind === "ifStart") {
    return hasFollowingKind(kinds, index, ["elseBranch", "endIf"]);
  }

  if (kind === "directiveIfStart") {
    return hasFollowingKind(kinds, index, ["directiveElseBranch", "directiveEndIf"]);
  }

  return true;
}

function hasFollowingKind(kinds: LineKind[], index: number, targets: LineKind[]): boolean {
  return kinds.slice(index + 1).some((kind) => targets.includes(kind));
}

function parseAlignableDeclarationLine(line: string): AlignableDeclarationLine | undefined {
  const indent = line.match(/^\s*/u)?.[0] ?? "";
  const { code, comment } = splitCodeAndComment(line);

  if (code.trim().length === 0 || code.includes(":") || /\s+_\s*$/.test(code)) {
    return undefined;
  }

  const trimmedCode = code.trim();
  const dimMatch = /^(Dim\s+[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?(?:\([^)]*\))?)(?:\s+As\s+(.+))?$/iu.exec(trimmedCode);

  if (dimMatch) {
    return {
      comment,
      head: normalizeInlineSpacing(dimMatch[1] ?? trimmedCode),
      indent,
      kind: "dim",
      typeName: dimMatch[2]?.trim()
    };
  }

  const constMatch =
    /^(?:(Public|Private)\s+)?Const\s+([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)(?:\s+As\s+([^=]+?))?\s*=\s*(.+)$/iu.exec(trimmedCode);

  if (constMatch) {
    const modifier = constMatch[1] ? `${constMatch[1]} ` : "";
    return {
      comment,
      head: `${modifier}Const ${constMatch[2]}`,
      indent,
      kind: "const",
      typeName: constMatch[3]?.trim(),
      value: constMatch[4]?.trim()
    };
  }

  const declareMatch =
    /^((?:(?:Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+[A-Za-z_][A-Za-z0-9_]*)(\s+Lib\s+.+)$/iu.exec(trimmedCode);

  if (declareMatch) {
    return {
      comment,
      head: normalizeInlineSpacing(declareMatch[1] ?? trimmedCode),
      indent,
      kind: "declare",
      tail: normalizeDeclareTail(declareMatch[2] ?? "")
    };
  }

  return undefined;
}

function canJoinDeclarationGroup(left: AlignableDeclarationLine, right: AlignableDeclarationLine): boolean {
  return left.kind === right.kind && left.indent === right.indent;
}

function formatDeclarationGroup(group: AlignableDeclarationLine[]): string[] {
  switch (group[0]?.kind) {
    case "const":
      return formatConstGroup(group);
    case "declare":
      return formatDeclareGroup(group);
    case "dim":
      return formatDimGroup(group);
    default:
      return group.map((line) => `${line.indent}${line.head}`);
  }
}

function formatDimGroup(group: AlignableDeclarationLine[]): string[] {
  const headWidth = getMaxLength(group.map((line) => line.head));

  return group.map((line) => {
    const code = line.typeName ? `${line.head.padEnd(headWidth)} As ${line.typeName.trim()}` : line.head;
    return withTrailingComment(line.indent + code, line.comment);
  });
}

function formatConstGroup(group: AlignableDeclarationLine[]): string[] {
  const headWidth = getMaxLength(group.map((line) => line.head));
  const typedLines = group.filter((line) => line.typeName);
  const typeWidth = getMaxLength(typedLines.map((line) => line.typeName ?? ""));

  return group.map((line) => {
    const typeSegment = line.typeName ? ` As ${line.typeName.trim().padEnd(typeWidth)}` : "";
    const code = `${line.head.padEnd(headWidth)}${typeSegment} = ${line.value?.trim() ?? ""}`;
    return withTrailingComment(line.indent + code, line.comment);
  });
}

function formatDeclareGroup(group: AlignableDeclarationLine[]): string[] {
  const headWidth = getMaxLength(group.map((line) => line.head));

  return group.map((line) => {
    const code = `${line.head.padEnd(headWidth)} ${line.tail?.trimStart() ?? ""}`.trimEnd();
    return withTrailingComment(line.indent + code, line.comment);
  });
}

function getMaxLength(values: string[]): number {
  return values.reduce((maxLength, value) => Math.max(maxLength, value.length), 0);
}

function normalizeDeclareTail(tail: string): string {
  return tail
    .trimStart()
    .replace(/\)\s+As\s+/iu, ") As ")
    .replace(/\s+Alias\s+/iu, " Alias ");
}

function normalizeInlineSpacing(text: string): string {
  return text.trim().replace(/\s+/gu, " ");
}

function isLabelLine(trimmedLine: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_]*:\s*(?:$|'.*|Rem\b.*)/i.test(trimmedLine);
}

function isLabelSegment(trimmedSegment: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_]*$/i.test(trimmedSegment);
}

function popTop(stack: BlockKind[], expected: BlockKind): void {
  if (stack.at(-1) === expected) {
    stack.pop();
  }
}

function getContinuationContext(lines: string[]): ContinuationContext {
  const trimmedContinuationLines = lines.slice(1).map((line) => line.trimStart());

  if (trimmedContinuationLines.some((line) => line.startsWith("."))) {
    return { kind: "methodChain" };
  }

  const firstLine = lines[0] ?? "";
  const { code } = splitCodeAndComment(firstLine);
  const continuationAnchor = code.replace(/\s+_\s*$/, "").trimEnd();
  const scrubbedAnchor = removeStringAndDateLiterals(continuationAnchor);

  if (getParenthesisBalance(scrubbedAnchor) > 0) {
    return { kind: "argumentList" };
  }

  if (/^\s*(?:Set\s+)?[A-Za-z_][A-Za-z0-9_\.]*(?:\([^)]*\))?\s*=.*$/iu.test(continuationAnchor)) {
    return { kind: "assignment" };
  }

  return { kind: "generic" };
}

function getIndentLevelForContinuationLine(
  trimmedLine: string,
  baseIndentLevel: number,
  continuationIndentLevels: number,
  context: ContinuationContext
): number {
  if (isStandaloneClosingParenthesisLine(trimmedLine) && context.kind === "argumentList") {
    return baseIndentLevel;
  }

  return baseIndentLevel + continuationIndentLevels;
}

function getIndentLevelForFirstLine(trimmedLine: string, baseIndentLevel: number): number {
  return isLabelLine(trimmedLine) ? 0 : baseIndentLevel;
}

function getParenthesisBalance(text: string): number {
  let balance = 0;

  for (const character of text) {
    if (character === "(") {
      balance += 1;
      continue;
    }

    if (character === ")") {
      balance -= 1;
    }
  }

  return balance;
}

function isStandaloneClosingParenthesisLine(trimmedLine: string): boolean {
  const { code } = splitCodeAndComment(trimmedLine);
  return /^\)\s*(?:_)?\s*$/u.test(code.trimEnd());
}

function trimLogicalLinePreservingContinuation(line: string, preserveContinuationSuffix: boolean): string {
  const trimmedLine = line.trimStart().trimEnd();

  if (!preserveContinuationSuffix) {
    return trimmedLine;
  }

  const { code, comment } = splitCodeAndComment(trimmedLine);
  const normalizedCode = code.replace(/\s+_\s*$/, " _").trimEnd();
  return normalizedCode + (comment ?? "");
}

function withTrailingComment(code: string, comment?: string): string {
  return comment ? `${code} ${comment.trimStart()}` : code;
}
