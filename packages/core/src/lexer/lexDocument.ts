import { VBA_KEYWORDS } from "./keywords";
import { createMappedRange, createSourceDocument } from "../types/helpers";
import { AnalyzeModuleOptions, SourceDocument, Token } from "../types/model";

export function lexDocument(text: string, options: AnalyzeModuleOptions = {}): Token[] {
  const source = createSourceDocument(text, options);
  return lexPreparedDocument(source);
}

export function lexPreparedDocument(source: SourceDocument): Token[] {
  const tokens: Token[] = [];

  for (let lineIndex = 0; lineIndex < source.normalizedLines.length; lineIndex += 1) {
    const line = source.normalizedLines[lineIndex];
    const trimmedLine = line.trimStart();
    const leadingWhitespace = line.length - trimmedLine.length;

    if (trimmedLine.length > 0 && /^Attribute\s+VB_/i.test(trimmedLine)) {
      tokens.push({
        kind: "attribute",
        range: createMappedRange(source, lineIndex, leadingWhitespace, lineIndex, line.length),
        text: trimmedLine
      });
      tokens.push({
        kind: "newline",
        range: createMappedRange(source, lineIndex, line.length, lineIndex, line.length),
        text: "\n"
      });
      continue;
    }

    if (trimmedLine.length > 0 && trimmedLine.startsWith("#")) {
      tokens.push({
        kind: "directive",
        range: createMappedRange(source, lineIndex, leadingWhitespace, lineIndex, line.length),
        text: trimmedLine
      });
      tokens.push({
        kind: "newline",
        range: createMappedRange(source, lineIndex, line.length, lineIndex, line.length),
        text: "\n"
      });
      continue;
    }

    let index = 0;
    let atTokenBoundary = true;

    while (index < line.length) {
      const currentCharacter = line[index];

      if (/\s/.test(currentCharacter)) {
        index += 1;
        atTokenBoundary = true;
        continue;
      }

      if (currentCharacter === "'") {
        tokens.push({
          kind: "comment",
          range: createMappedRange(source, lineIndex, index, lineIndex, line.length),
          text: line.slice(index)
        });
        index = line.length;
        break;
      }

      if (
        atTokenBoundary &&
        /[Rr]/.test(currentCharacter) &&
        line.slice(index, index + 3).toLowerCase() === "rem" &&
        (index + 3 === line.length || /\s/.test(line[index + 3] ?? ""))
      ) {
        tokens.push({
          kind: "comment",
          range: createMappedRange(source, lineIndex, index, lineIndex, line.length),
          text: line.slice(index)
        });
        index = line.length;
        break;
      }

      if (currentCharacter === "\"") {
        const endIndex = readStringLiteral(line, index);
        tokens.push({
          kind: "stringLiteral",
          range: createMappedRange(source, lineIndex, index, lineIndex, endIndex),
          text: line.slice(index, endIndex)
        });
        index = endIndex;
        atTokenBoundary = false;
        continue;
      }

      if (currentCharacter === "#") {
        const endIndex = readDateLiteral(line, index);
        tokens.push({
          kind: "dateLiteral",
          range: createMappedRange(source, lineIndex, index, lineIndex, endIndex),
          text: line.slice(index, endIndex)
        });
        index = endIndex;
        atTokenBoundary = false;
        continue;
      }

      if (currentCharacter === "_" && isLineContinuationMarker(line, index)) {
        tokens.push({
          kind: "lineContinuation",
          range: createMappedRange(source, lineIndex, index, lineIndex, index + 1),
          text: currentCharacter
        });
        index += 1;
        atTokenBoundary = false;
        continue;
      }

      if (/[A-Za-z_]/.test(currentCharacter)) {
        const endIndex = readIdentifier(line, index);
        const text = line.slice(index, endIndex);
        tokens.push({
          kind: VBA_KEYWORDS.has(text.replace(/[$%&!#@]$/, "").toLowerCase()) ? "keyword" : "identifier",
          range: createMappedRange(source, lineIndex, index, lineIndex, endIndex),
          text
        });
        index = endIndex;
        atTokenBoundary = false;
        continue;
      }

      if (/\d/.test(currentCharacter)) {
        const endIndex = readNumberLiteral(line, index);
        tokens.push({
          kind: "numberLiteral",
          range: createMappedRange(source, lineIndex, index, lineIndex, endIndex),
          text: line.slice(index, endIndex)
        });
        index = endIndex;
        atTokenBoundary = false;
        continue;
      }

      const tokenKind = /[(),.:]/.test(currentCharacter) ? "punctuation" : "operator";
      tokens.push({
        kind: tokenKind,
        range: createMappedRange(source, lineIndex, index, lineIndex, index + 1),
        text: currentCharacter
      });
      index += 1;
      atTokenBoundary = /[(:.,]/.test(currentCharacter);
    }

    tokens.push({
      kind: "newline",
      range: createMappedRange(source, lineIndex, line.length, lineIndex, line.length),
      text: "\n"
    });
  }

  const lastLineIndex = Math.max(0, source.normalizedLines.length - 1);
  const lastLineLength = source.normalizedLines[lastLineIndex]?.length ?? 0;
  tokens.push({
    kind: "eof",
    range: createMappedRange(source, lastLineIndex, lastLineLength, lastLineIndex, lastLineLength),
    text: ""
  });

  return tokens;
}

function readDateLiteral(line: string, start: number): number {
  let index = start + 1;

  while (index < line.length) {
    if (line[index] === "#") {
      return index + 1;
    }

    index += 1;
  }

  return line.length;
}

function isLineContinuationMarker(line: string, index: number): boolean {
  if (index === 0 || !/\s/.test(line[index - 1] ?? "")) {
    return false;
  }

  let probe = index + 1;

  while (probe < line.length && /\s/.test(line[probe])) {
    probe += 1;
  }

  const hasWhitespaceAfterMarker = probe > index + 1;

  return (
    probe === line.length ||
    line[probe] === "'" ||
    (hasWhitespaceAfterMarker &&
      line.slice(probe, probe + 3).toLowerCase() === "rem" &&
      (probe + 3 === line.length || /\s/.test(line[probe + 3] ?? "")))
  );
}

function readIdentifier(line: string, start: number): number {
  let index = start + 1;

  while (index < line.length && /[A-Za-z0-9_]/.test(line[index])) {
    index += 1;
  }

  if (/[$%&!#@]/.test(line[index] ?? "")) {
    index += 1;
  }

  return index;
}

function readNumberLiteral(line: string, start: number): number {
  let index = start + 1;

  while (index < line.length && /[\d.]/.test(line[index])) {
    index += 1;
  }

  if (/[%&!#@]/.test(line[index] ?? "")) {
    index += 1;
  }

  return index;
}

function readStringLiteral(line: string, start: number): number {
  let index = start + 1;

  while (index < line.length) {
    if (line[index] === "\"" && line[index + 1] === "\"") {
      index += 2;
      continue;
    }

    if (line[index] === "\"") {
      return index + 1;
    }

    index += 1;
  }

  return line.length;
}
