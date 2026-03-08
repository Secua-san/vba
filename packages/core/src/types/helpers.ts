import { AnalyzeModuleOptions, LinePosition, ModuleKind, SourceDocument, SourceRange } from "./model";

export function comparePositions(left: LinePosition, right: LinePosition): number {
  if (left.line !== right.line) {
    return left.line - right.line;
  }

  return left.character - right.character;
}

export function createPosition(line: number, character: number): LinePosition {
  return { line, character };
}

export function createRange(startLine: number, startCharacter: number, endLine: number, endCharacter: number): SourceRange {
  return {
    start: createPosition(startLine, startCharacter),
    end: createPosition(endLine, endCharacter)
  };
}

export function createMappedRange(
  source: SourceDocument,
  normalizedStartLine: number,
  startCharacter: number,
  normalizedEndLine: number,
  endCharacter: number
): SourceRange {
  return createRange(
    source.lineMap[normalizedStartLine] ?? normalizedStartLine,
    startCharacter,
    source.lineMap[normalizedEndLine] ?? normalizedEndLine,
    endCharacter
  );
}

export function getModuleKind(fileName?: string): ModuleKind {
  const lowerFileName = fileName?.toLowerCase();

  if (lowerFileName?.endsWith(".frm")) {
    return "form";
  }

  if (lowerFileName?.endsWith(".cls")) {
    return "class";
  }

  return "standard";
}

export function getModuleNameFromFileName(fileName?: string): string {
  if (!fileName) {
    return "Module1";
  }

  const normalizedFileName = fileName.replace(/\\/g, "/");
  const lastSegment = normalizedFileName.split("/").pop() ?? normalizedFileName;
  const extensionIndex = lastSegment.lastIndexOf(".");
  return extensionIndex >= 0 ? lastSegment.slice(0, extensionIndex) : lastSegment;
}

export function normalizeIdentifier(value: string): string {
  return value.replace(/[$%&!#@]$/, "").toLowerCase();
}

export function positionInRange(position: LinePosition, range: SourceRange): boolean {
  return comparePositions(position, range.start) >= 0 && comparePositions(position, range.end) <= 0;
}

export function typeNameFromSuffix(identifier: string): string | undefined {
  const suffix = identifier.at(-1);

  switch (suffix) {
    case "$":
      return "String";
    case "%":
      return "Integer";
    case "&":
      return "Long";
    case "!":
      return "Single";
    case "#":
      return "Double";
    case "@":
      return "Currency";
    default:
      return undefined;
  }
}

export function createSourceDocument(text: string, options: AnalyzeModuleOptions = {}): SourceDocument {
  const normalizedOriginalText = text.replace(/\r\n?/g, "\n");
  const originalLines = normalizedOriginalText.split("\n");
  const moduleKind = getModuleKind(options.fileName);
  let startLine = 0;

  if (moduleKind === "form") {
    const codeStartIndex = originalLines.findIndex((line) => {
      const trimmedLine = line.trimStart();
      return /^(Attribute\s+VB_|Option\b|Public\b|Private\b|Friend\b|Static\b|Sub\b|Function\b|Property\b|Dim\b|Const\b|Enum\b|Type\b|Declare\b|#If\b|#Else\b|#End\b)/i.test(
        trimmedLine
      );
    });

    startLine = codeStartIndex >= 0 ? codeStartIndex : originalLines.length;
  }

  const normalizedLines = originalLines.slice(startLine);
  const lineMap = normalizedLines.map((_, index) => startLine + index);

  return {
    fileName: options.fileName,
    lineMap,
    moduleKind,
    moduleName: options.moduleName ?? getModuleNameFromFileName(options.fileName),
    normalizedLines,
    normalizedText: normalizedLines.join("\n"),
    originalLines,
    originalText: normalizedOriginalText
  };
}
