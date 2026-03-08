import { LinePosition, SourceDocument } from "../types/model";

export interface LogicalLine {
  codeText: string;
  endLine: number;
  fullText: string;
  startLine: number;
}

export function buildLogicalLines(source: SourceDocument): LogicalLine[] {
  const lines: LogicalLine[] = [];
  let bufferCode = "";
  let bufferText = "";
  let startLine = 0;
  let hasBuffer = false;

  for (let index = 0; index < source.normalizedLines.length; index += 1) {
    const rawLine = source.normalizedLines[index];
    const { code, comment } = splitCodeAndComment(rawLine);
    const continued = isLineContinuation(code);
    const codeWithoutContinuation = continued ? code.replace(/\s+_\s*$/, "") : code;

    if (!hasBuffer) {
      startLine = index;
      hasBuffer = true;
    }

    const trimmedCode = codeWithoutContinuation.trimEnd();
    bufferCode += bufferCode.length > 0 ? ` ${trimmedCode.trimStart()}` : trimmedCode;
    bufferText += bufferText.length > 0 ? `\n${rawLine}` : rawLine;

    if (!continued) {
      lines.push({
        codeText: bufferCode.trimEnd(),
        endLine: index,
        fullText: bufferText + (comment ? "" : ""),
        startLine
      });
      bufferCode = "";
      bufferText = "";
      hasBuffer = false;
    }
  }

  if (hasBuffer) {
    lines.push({
      codeText: bufferCode.trimEnd(),
      endLine: source.normalizedLines.length - 1,
      fullText: bufferText,
      startLine
    });
  }

  return lines;
}

export function extractIdentifierAtPosition(text: string, position: LinePosition): string | undefined {
  const lines = text.split("\n");
  const line = lines[position.line];

  if (line === undefined) {
    return undefined;
  }

  const identifierPattern = /[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?/g;

  for (const match of line.matchAll(identifierPattern)) {
    const start = match.index ?? 0;
    const end = start + match[0].length;

    if (position.character >= start && position.character <= end) {
      return match[0];
    }
  }

  return undefined;
}

export function removeStringAndDateLiterals(text: string): string {
  let result = "";
  let index = 0;

  while (index < text.length) {
    const currentCharacter = text[index];

    if (currentCharacter === "\"") {
      result += " ";
      index += 1;

      while (index < text.length) {
        if (text[index] === "\"" && text[index + 1] === "\"") {
          result += "  ";
          index += 2;
          continue;
        }

        if (text[index] === "\"") {
          result += " ";
          index += 1;
          break;
        }

        result += " ";
        index += 1;
      }

      continue;
    }

    if (currentCharacter === "#") {
      result += " ";
      index += 1;

      while (index < text.length) {
        result += " ";

        if (text[index] === "#") {
          index += 1;
          break;
        }

        index += 1;
      }

      continue;
    }

    result += currentCharacter;
    index += 1;
  }

  return result;
}

export function splitCodeAndComment(line: string): { code: string; comment?: string } {
  let index = 0;
  let atTokenBoundary = true;

  while (index < line.length) {
    const currentCharacter = line[index];

    if (currentCharacter === "\"") {
      index += 1;

      while (index < line.length) {
        if (line[index] === "\"" && line[index + 1] === "\"") {
          index += 2;
          continue;
        }

        if (line[index] === "\"") {
          index += 1;
          break;
        }

        index += 1;
      }

      atTokenBoundary = false;
      continue;
    }

    if (currentCharacter === "'") {
      return {
        code: line.slice(0, index),
        comment: line.slice(index)
      };
    }

    if (
      atTokenBoundary &&
      /[Rr]/.test(currentCharacter) &&
      line.slice(index, index + 3).toLowerCase() === "rem" &&
      (index + 3 === line.length || /\s/.test(line[index + 3] ?? ""))
    ) {
      return {
        code: line.slice(0, index),
        comment: line.slice(index)
      };
    }

    atTokenBoundary = /\s|[:(,]/.test(currentCharacter);
    index += 1;
  }

  return { code: line };
}

export function splitCommaAware(text: string): string[] {
  const values: string[] = [];
  let buffer = "";
  let depth = 0;
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

    if (currentCharacter === "," && depth === 0) {
      values.push(buffer.trim());
      buffer = "";
      index += 1;
      continue;
    }

    buffer += currentCharacter;
    index += 1;
  }

  if (buffer.trim().length > 0) {
    values.push(buffer.trim());
  }

  return values;
}

function isLineContinuation(code: string): boolean {
  return /\s+_\s*$/.test(code);
}
