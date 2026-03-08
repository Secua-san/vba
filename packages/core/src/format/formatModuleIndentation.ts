import { buildLogicalLines } from "../parser/text";
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

export function formatModuleIndentation(text: string, options: FormatModuleIndentationOptions = {}): string {
  const lineEnding = text.includes("\r\n") ? "\r\n" : "\n";
  const source = createSourceDocument(text, { fileName: options.fileName });

  if (source.normalizedLines.length === 0) {
    return text;
  }

  const indentUnit = options.insertSpaces === false ? "\t" : " ".repeat(options.indentSize ?? 4);
  const continuationIndentLevels = Math.max(0, options.continuationIndentLevels ?? 1);
  const formattedLines = [...source.normalizedLines];
  const stack: BlockKind[] = [];

  for (const logicalLine of buildLogicalLines(source)) {
    const kind = classifyLineKind(logicalLine.codeText.trim());
    applyPreIndentation(kind, stack);

    const baseIndentLevel = getBaseIndentLevel(kind, stack);

    for (let lineIndex = logicalLine.startLine; lineIndex <= logicalLine.endLine; lineIndex += 1) {
      const originalLine = source.normalizedLines[lineIndex] ?? "";
      const trimmedLine = originalLine.trimStart();

      if (trimmedLine.length === 0) {
        formattedLines[lineIndex] = "";
        continue;
      }

      const continuationOffset = lineIndex === logicalLine.startLine ? 0 : continuationIndentLevels;
      const indentLevel = isLabelLine(trimmedLine) ? 0 : baseIndentLevel + continuationOffset;
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

function isLabelLine(trimmedLine: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_]*:\s*(?:$|'.*|Rem\b.*)/i.test(trimmedLine);
}

function popTop(stack: BlockKind[], expected: BlockKind): void {
  if (stack.at(-1) === expected) {
    stack.pop();
  }
}
