import { lexPreparedDocument } from "../lexer/lexDocument";
import { createMappedRange, createSourceDocument, normalizeIdentifier, typeNameFromSuffix } from "../types/helpers";
import {
  AnalyzeModuleOptions,
  AttributeLineNode,
  ConstDeclarationNode,
  DeclareStatementNode,
  Diagnostic,
  DirectiveNode,
  EnumDeclarationNode,
  EnumMemberNode,
  ModuleMemberNode,
  ModuleNode,
  OptionStatementNode,
  ParameterNode,
  ParseResult,
  ProcedureDeclarationNode,
  ProcedureKind,
  ProcedureStatementNode,
  SourceDocument,
  Token,
  TypeDeclarationNode,
  TypeMemberNode,
  VariableDeclarationNode,
  VariableDeclaratorNode
} from "../types/model";
import { buildLogicalLines, splitCommaAware } from "./text";

export function parseModule(text: string, options: AnalyzeModuleOptions = {}): ParseResult {
  const source = createSourceDocument(text, options);
  const tokens = lexPreparedDocument(source);
  return parsePreparedModule(source, tokens);
}

export function parsePreparedModule(source: SourceDocument, tokens: Token[]): ParseResult {
  const diagnostics: Diagnostic[] = [];
  const members: ModuleMemberNode[] = [];
  const logicalLines = buildLogicalLines(source);
  let moduleName = source.moduleName;
  let index = 0;

  while (index < logicalLines.length) {
    const logicalLine = logicalLines[index];
    const trimmedText = logicalLine.codeText.trim();

    if (trimmedText.length === 0) {
      index += 1;
      continue;
    }

    if (/^Attribute\s+/i.test(trimmedText)) {
      const attribute = parseAttributeLine(source, logicalLine);
      members.push(attribute);

      if (attribute.name.toLowerCase() === "vb_name" && attribute.value) {
        moduleName = attribute.value;
      }

      index += 1;
      continue;
    }

    if (/^#/i.test(trimmedText)) {
      members.push(createDirectiveNode(source, logicalLine));
      index += 1;
      continue;
    }

    if (/^Option\b/i.test(trimmedText)) {
      members.push(parseOptionStatement(source, logicalLine));
      index += 1;
      continue;
    }

    if (isProcedureHeader(trimmedText)) {
      const procedure = parseProcedure(source, logicalLines, index, diagnostics);
      members.push(procedure.node);
      index = procedure.nextIndex;
      continue;
    }

    if (isDeclareLine(trimmedText)) {
      const declaration = parseDeclareStatement(source, logicalLine, diagnostics);
      members.push(declaration);
      index += 1;
      continue;
    }

    if (isConstDeclaration(trimmedText)) {
      members.push(parseConstDeclaration(source, logicalLine));
      index += 1;
      continue;
    }

    if (isVariableDeclaration(trimmedText)) {
      members.push(parseVariableDeclaration(source, logicalLine));
      index += 1;
      continue;
    }

    if (isEnumDeclaration(trimmedText)) {
      const enumeration = parseEnumDeclaration(source, logicalLines, index, diagnostics);
      members.push(enumeration.node);
      index = enumeration.nextIndex;
      continue;
    }

    if (isTypeDeclaration(trimmedText)) {
      const userType = parseTypeDeclaration(source, logicalLines, index, diagnostics);
      members.push(userType.node);
      index = userType.nextIndex;
      continue;
    }

    diagnostics.push({
      code: "syntax-error",
      message: "Unrecognized module-level statement.",
      range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
      severity: "error"
    });
    members.push({
      kind: "unknownStatement",
      range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
      text: logicalLine.codeText
    });
    index += 1;
  }

  source.moduleName = moduleName;

  const moduleRange =
    source.normalizedLines.length === 0
      ? createMappedRange(source, 0, 0, 0, 0)
      : createMappedRange(
          source,
          0,
          0,
          source.normalizedLines.length - 1,
          source.normalizedLines[source.normalizedLines.length - 1]?.length ?? 0
        );

  const module: ModuleNode = {
    kind: "module",
    members,
    name: moduleName,
    range: moduleRange
  };

  return {
    diagnostics,
    module,
    source,
    tokens
  };
}

function createDirectiveNode(source: SourceDocument, logicalLine: { codeText: string; endLine: number; startLine: number }): DirectiveNode {
  return {
    kind: "directive",
    range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
    text: logicalLine.codeText
  };
}

function isConstDeclaration(text: string): boolean {
  return /^(?:(?:Public|Private)\s+)?Const\b/i.test(text);
}

function isDeclareLine(text: string): boolean {
  return /^(?:(?:Public|Private)\s+)?Declare\b/i.test(text);
}

function isEnumDeclaration(text: string): boolean {
  return /^(?:(?:Public|Private)\s+)?Enum\b/i.test(text);
}

function isProcedureHeader(text: string): boolean {
  return /^(?:(?:Public|Private|Friend)\s+)?(?:(?:Static)\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\b/i.test(text);
}

function isTypeDeclaration(text: string): boolean {
  return /^(?:(?:Public|Private)\s+)?Type\b/i.test(text);
}

function isVariableDeclaration(text: string): boolean {
  return /^(?:(?:Public|Private)\s+)?Dim\b/i.test(text);
}

function parseAttributeLine(source: SourceDocument, logicalLine: { codeText: string; endLine: number; startLine: number }): AttributeLineNode {
  const match = /^Attribute\s+([A-Za-z0-9_]+)\s*=\s*(.+)$/i.exec(logicalLine.codeText.trim());
  const value = match?.[2]?.trim();

  return {
    kind: "attributeLine",
    name: match?.[1] ?? logicalLine.codeText.trim(),
    range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
    text: logicalLine.codeText,
    value: value ? value.replace(/^"|"$/g, "") : undefined
  };
}

function parseConstDeclaration(source: SourceDocument, logicalLine: { codeText: string; endLine: number; startLine: number }): ConstDeclarationNode {
  const match = /^(?:(Public|Private)\s+)?Const\s+([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)(?:\s+As\s+([A-Za-z_][A-Za-z0-9_\.]*))?/i.exec(
    logicalLine.codeText.trim()
  );

  return {
    kind: "constDeclaration",
    modifier: match?.[1],
    name: match?.[2] ?? logicalLine.codeText.trim(),
    range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
    text: logicalLine.codeText,
    typeName: match?.[3] ?? typeNameFromSuffix(match?.[2] ?? "")
  };
}

function parseDeclareStatement(
  source: SourceDocument,
  logicalLine: { codeText: string; endLine: number; startLine: number },
  diagnostics: Diagnostic[]
): DeclareStatementNode {
  const text = logicalLine.codeText.trim();
  const headerMatch = /^(?:(Public|Private)\s+)?Declare\s+(PtrSafe\s+)?(Function|Sub)\s+([A-Za-z_][A-Za-z0-9_]*)/i.exec(text);
  const openParenIndex = text.indexOf("(");
  const closeParenIndex = text.lastIndexOf(")");
  const parametersText = openParenIndex >= 0 && closeParenIndex > openParenIndex ? text.slice(openParenIndex + 1, closeParenIndex) : "";
  const returnMatch = /\)\s+As\s+([A-Za-z_][A-Za-z0-9_\.]*)/i.exec(text);
  const parameters = parseParameters(parametersText, source, logicalLine.startLine, logicalLine.endLine);
  const range = createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0);

  if (!headerMatch?.[2]) {
    diagnostics.push({
      code: "declare-missing-ptrsafe",
      message: "Declare statements must include PtrSafe on Win64.",
      range,
      severity: "error"
    });
  }

  if (openParenIndex >= 0 && closeParenIndex < openParenIndex) {
    diagnostics.push({
      code: "syntax-error",
      message: "Malformed Declare parameter list.",
      range,
      severity: "error"
    });
  }

  return {
    isPtrSafe: Boolean(headerMatch?.[2]),
    kind: "declareStatement",
    modifier: headerMatch?.[1],
    name: headerMatch?.[4] ?? text,
    parameters,
    procedureKind: (headerMatch?.[3] as "Function" | "Sub") ?? "Function",
    range,
    returnType: returnMatch?.[1],
    text
  };
}

function parseEnumDeclaration(
  source: SourceDocument,
  logicalLines: Array<{ codeText: string; endLine: number; startLine: number }>,
  startIndex: number,
  diagnostics: Diagnostic[]
): { nextIndex: number; node: EnumDeclarationNode } {
  const startLine = logicalLines[startIndex];
  const headerMatch = /^(?:(Public|Private)\s+)?Enum\s+([A-Za-z_][A-Za-z0-9_]*)/i.exec(startLine.codeText.trim());
  const members: EnumMemberNode[] = [];
  let endIndex = startIndex + 1;

  while (endIndex < logicalLines.length) {
    const currentLine = logicalLines[endIndex];
    const trimmedText = currentLine.codeText.trim();

    if (/^End\s+Enum\b/i.test(trimmedText)) {
      const node: EnumDeclarationNode = {
        kind: "enumDeclaration",
        members,
        modifier: headerMatch?.[1],
        name: headerMatch?.[2] ?? startLine.codeText.trim(),
        range: createMappedRange(source, startLine.startLine, 0, currentLine.endLine, source.normalizedLines[currentLine.endLine]?.length ?? 0),
        text: logicalLines.slice(startIndex, endIndex + 1).map((line) => line.codeText).join("\n")
      };

      return {
        nextIndex: endIndex + 1,
        node
      };
    }

    if (trimmedText.length > 0) {
      const memberMatch = /^([A-Za-z_][A-Za-z0-9_]*)/.exec(trimmedText);

      if (memberMatch) {
        members.push({
          kind: "enumMember",
          name: memberMatch[1],
          range: createMappedRange(source, currentLine.startLine, 0, currentLine.endLine, source.normalizedLines[currentLine.endLine]?.length ?? 0)
        });
      }
    }

    endIndex += 1;
  }

  diagnostics.push({
    code: "syntax-error",
    message: "Enum declaration is missing End Enum.",
    range: createMappedRange(source, startLine.startLine, 0, startLine.endLine, source.normalizedLines[startLine.endLine]?.length ?? 0),
    severity: "error"
  });

  return {
    nextIndex: logicalLines.length,
    node: {
      kind: "enumDeclaration",
      members,
      modifier: headerMatch?.[1],
      name: headerMatch?.[2] ?? startLine.codeText.trim(),
      range: createMappedRange(source, startLine.startLine, 0, logicalLines[logicalLines.length - 1]?.endLine ?? startLine.endLine, source.normalizedLines[logicalLines[logicalLines.length - 1]?.endLine ?? startLine.endLine]?.length ?? 0),
      text: logicalLines.slice(startIndex).map((line) => line.codeText).join("\n")
    }
  };
}

function parseOptionStatement(source: SourceDocument, logicalLine: { codeText: string; endLine: number; startLine: number }): OptionStatementNode {
  const match = /^Option\s+(.+)$/i.exec(logicalLine.codeText.trim());

  return {
    kind: "optionStatement",
    name: match?.[1]?.trim() ?? logicalLine.codeText.trim(),
    range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
    text: logicalLine.codeText
  };
}

function parseParameters(parametersText: string, source: SourceDocument, startLine: number, endLine: number): ParameterNode[] {
  if (parametersText.trim().length === 0) {
    return [];
  }

  return splitCommaAware(parametersText).map((parameterText) => {
    const tokens = parameterText.trim().split(/\s+/);
    let cursor = 0;
    let isOptional = false;
    let isParamArray = false;
    let direction: "byRef" | "byVal" = "byRef";

    while (cursor < tokens.length) {
      const lowerToken = tokens[cursor].toLowerCase();

      if (lowerToken === "optional") {
        isOptional = true;
        cursor += 1;
        continue;
      }

      if (lowerToken === "paramarray") {
        isParamArray = true;
        cursor += 1;
        continue;
      }

      if (lowerToken === "byval") {
        direction = "byVal";
        cursor += 1;
        continue;
      }

      if (lowerToken === "byref") {
        direction = "byRef";
        cursor += 1;
        continue;
      }

      break;
    }

    const nameToken = tokens[cursor] ?? parameterText.trim();
    const arraySuffix = /\(\)$/.test(nameToken);
    const cleanName = nameToken.replace(/\(\)$/, "");
    const asIndex = tokens.findIndex((token) => token.toLowerCase() === "as");
    const typeName = asIndex >= 0 ? tokens.slice(asIndex + 1).join(" ").split("=")[0].trim() : typeNameFromSuffix(cleanName);

    return {
      arraySuffix,
      direction,
      isOptional,
      isParamArray,
      kind: "parameter",
      name: cleanName.replace(/[$%&!#@]$/, ""),
      range: createMappedRange(source, startLine, 0, endLine, source.normalizedLines[endLine]?.length ?? 0),
      typeName: typeName || undefined
    };
  });
}

function parseProcedure(
  source: SourceDocument,
  logicalLines: Array<{ codeText: string; endLine: number; startLine: number }>,
  startIndex: number,
  diagnostics: Diagnostic[]
): { nextIndex: number; node: ProcedureDeclarationNode } {
  const headerLine = logicalLines[startIndex];
  const headerText = headerLine.codeText.trim();
  const headerMatch =
    /^(?:(Public|Private|Friend)\s+)?(?:(Static)\s+)?(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+([A-Za-z_][A-Za-z0-9_]*)(.*)$/i.exec(
      headerText
    );
  const procedureKind = normalizeProcedureKind(headerMatch?.[3] ?? "Sub");
  const openParenIndex = headerText.indexOf("(");
  const closeParenIndex = headerText.lastIndexOf(")");
  const parametersText = openParenIndex >= 0 && closeParenIndex > openParenIndex ? headerText.slice(openParenIndex + 1, closeParenIndex) : "";
  const returnTypeMatch = /\)\s+As\s+([A-Za-z_][A-Za-z0-9_\.]*)/i.exec(headerText);
  const parameters = parseParameters(parametersText, source, headerLine.startLine, headerLine.endLine);
  const statements: ProcedureStatementNode[] = [];
  let endIndex = startIndex + 1;
  let foundEnd = false;

  while (endIndex < logicalLines.length) {
    const currentLine = logicalLines[endIndex];
    const trimmedText = currentLine.codeText.trim();

    if (isProcedureTerminator(trimmedText, procedureKind)) {
      foundEnd = true;
      break;
    }

    const statement = parseProcedureStatement(source, currentLine);

    if (statement) {
      statements.push(statement);
    }

    endIndex += 1;
  }

  const endLineIndex = foundEnd ? logicalLines[endIndex].endLine : logicalLines[Math.max(startIndex, endIndex - 1)]?.endLine ?? headerLine.endLine;

  if (!foundEnd) {
    diagnostics.push({
      code: "syntax-error",
      message: `${procedureKind} ${headerMatch?.[4] ?? headerText} is missing its terminator.`,
      range: createMappedRange(source, headerLine.startLine, 0, headerLine.endLine, source.normalizedLines[headerLine.endLine]?.length ?? 0),
      severity: "error"
    });
  }

  const node: ProcedureDeclarationNode = {
    body: statements,
    headerRange: createMappedRange(source, headerLine.startLine, 0, headerLine.endLine, source.normalizedLines[headerLine.endLine]?.length ?? 0),
    isStatic: Boolean(headerMatch?.[2]),
    kind: "procedureDeclaration",
    modifier: headerMatch?.[1],
    name: headerMatch?.[4] ?? headerText,
    parameters,
    procedureKind,
    range: createMappedRange(source, headerLine.startLine, 0, endLineIndex, source.normalizedLines[endLineIndex]?.length ?? 0),
    returnType: returnTypeMatch?.[1]
  };

  validateProcedureBlocks(source, node, diagnostics);

  return {
    nextIndex: foundEnd ? endIndex + 1 : logicalLines.length,
    node
  };
}

function parseProcedureStatement(
  source: SourceDocument,
  logicalLine: { codeText: string; endLine: number; startLine: number }
): ProcedureStatementNode | undefined {
  const trimmedText = logicalLine.codeText.trim();
  const statementRange = createMappedRange(
    source,
    logicalLine.startLine,
    0,
    logicalLine.endLine,
    source.normalizedLines[logicalLine.endLine]?.length ?? 0
  );
  const inlineCharacterOffset =
    logicalLine.startLine === logicalLine.endLine
      ? source.normalizedLines[logicalLine.startLine]?.length - (source.normalizedLines[logicalLine.startLine]?.trimStart().length ?? 0)
      : 0;

  if (trimmedText.length === 0) {
    return undefined;
  }

  if (/^Const\b/i.test(trimmedText)) {
    const constant = parseConstDeclaration(source, logicalLine);
    return {
      declaredConstants: [constant],
      kind: "constStatement",
      range: constant.range,
      text: trimmedText
    };
  }

  if (/^(?:Dim|Static)\b/i.test(trimmedText)) {
    const declaration = parseProcedureVariableDeclaration(source, logicalLine);
    return {
      declaredVariables: declaration.declarators,
      kind: "declarationStatement",
      range: declaration.range,
      text: trimmedText
    };
  }

  const structuredBlockStatement = parseStructuredBlockStatement(trimmedText, statementRange, inlineCharacterOffset);

  if (structuredBlockStatement) {
    return structuredBlockStatement;
  }

  const assignmentStatement = parseAssignmentStatement(trimmedText, statementRange, inlineCharacterOffset);

  if (assignmentStatement) {
    return assignmentStatement;
  }

  const callStatement = parseCallStatement(trimmedText, statementRange, inlineCharacterOffset);

  if (callStatement) {
    return callStatement;
  }

  return {
    kind: "executableStatement",
    range: statementRange,
    text: trimmedText
  };
}

function parseStructuredBlockStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const ifBlockStatement = parseIfBlockStatement(text, statementRange, inlineCharacterOffset);

  if (ifBlockStatement) {
    return ifBlockStatement;
  }

  const elseIfClauseStatement = parseElseIfClauseStatement(text, statementRange, inlineCharacterOffset);

  if (elseIfClauseStatement) {
    return elseIfClauseStatement;
  }

  if (/^Else\s*$/iu.test(text)) {
    return {
      kind: "elseClauseStatement",
      range: statementRange,
      text
    };
  }

  if (/^End\s+If\s*$/iu.test(text)) {
    return {
      kind: "endIfStatement",
      range: statementRange,
      text
    };
  }

  const selectCaseStatement = parseSelectCaseStatement(text, statementRange, inlineCharacterOffset);

  if (selectCaseStatement) {
    return selectCaseStatement;
  }

  const caseClauseStatement = parseCaseClauseStatement(text, statementRange, inlineCharacterOffset);

  if (caseClauseStatement) {
    return caseClauseStatement;
  }

  if (/^End\s+Select\s*$/iu.test(text)) {
    return {
      kind: "endSelectStatement",
      range: statementRange,
      text
    };
  }

  const forEachStatement = parseForEachStatement(text, statementRange, inlineCharacterOffset);

  if (forEachStatement) {
    return forEachStatement;
  }

  const forStatement = parseForStatement(text, statementRange, inlineCharacterOffset);

  if (forStatement) {
    return forStatement;
  }

  const nextStatement = parseNextStatement(text, statementRange, inlineCharacterOffset);

  if (nextStatement) {
    return nextStatement;
  }

  const doBlockStatement = parseDoBlockStatement(text, statementRange, inlineCharacterOffset);

  if (doBlockStatement) {
    return doBlockStatement;
  }

  const loopStatement = parseLoopStatement(text, statementRange, inlineCharacterOffset);

  if (loopStatement) {
    return loopStatement;
  }

  const whileStatement = parseWhileStatement(text, statementRange, inlineCharacterOffset);

  if (whileStatement) {
    return whileStatement;
  }

  if (/^Wend\s*$/iu.test(text)) {
    return {
      kind: "wendStatement",
      range: statementRange,
      text
    };
  }

  const withBlockStatement = parseWithBlockStatement(text, statementRange, inlineCharacterOffset);

  if (withBlockStatement) {
    return withBlockStatement;
  }

  if (/^End\s+With\s*$/iu.test(text)) {
    return {
      kind: "endWithStatement",
      range: statementRange,
      text
    };
  }

  const onErrorStatement = parseOnErrorStatement(text, statementRange, inlineCharacterOffset);

  if (onErrorStatement) {
    return onErrorStatement;
  }

  return undefined;
}

function parseIfBlockStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  if (!isIfBlockStart(text)) {
    return undefined;
  }

  const prefixLength = /^\s*If\s+/iu.exec(text)?.[0].length ?? 0;
  const suffixLength = /\s+Then\s*$/iu.exec(text)?.[0].length ?? 0;
  const conditionText = text.slice(prefixLength, text.length - suffixLength).trim();
  const leadingWhitespace = text.slice(prefixLength, text.length - suffixLength).length - text.slice(prefixLength, text.length - suffixLength).trimStart().length;
  const conditionStartCharacter = prefixLength + leadingWhitespace;

  if (conditionText.length === 0) {
    return undefined;
  }

  return {
    conditionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + conditionStartCharacter,
      inlineCharacterOffset + conditionStartCharacter + conditionText.length
    ),
    conditionText,
    kind: "ifBlockStatement",
    range: statementRange,
    text
  };
}

function parseElseIfClauseStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  if (!/^ElseIf\b.*\bThen\s*$/iu.test(text) || /:/.test(text)) {
    return undefined;
  }

  const prefixLength = /^\s*ElseIf\s+/iu.exec(text)?.[0].length ?? 0;
  const suffixLength = /\s+Then\s*$/iu.exec(text)?.[0].length ?? 0;
  const conditionSlice = text.slice(prefixLength, text.length - suffixLength);
  const conditionText = conditionSlice.trim();
  const leadingWhitespace = conditionSlice.length - conditionSlice.trimStart().length;
  const conditionStartCharacter = prefixLength + leadingWhitespace;

  if (conditionText.length === 0) {
    return undefined;
  }

  return {
    conditionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + conditionStartCharacter,
      inlineCharacterOffset + conditionStartCharacter + conditionText.length
    ),
    conditionText,
    kind: "elseIfClauseStatement",
    range: statementRange,
    text
  };
}

function parseSelectCaseStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const prefixLength = /^\s*Select\s+Case\s+/iu.exec(text)?.[0].length;

  if (!prefixLength) {
    return undefined;
  }

  const expressionSlice = text.slice(prefixLength);
  const expressionText = expressionSlice.trim();
  const leadingWhitespace = expressionSlice.length - expressionSlice.trimStart().length;
  const expressionStartCharacter = prefixLength + leadingWhitespace;

  if (expressionText.length === 0) {
    return undefined;
  }

  return {
    expressionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + expressionStartCharacter,
      inlineCharacterOffset + expressionStartCharacter + expressionText.length
    ),
    expressionText,
    kind: "selectCaseStatement",
    range: statementRange,
    text
  };
}

function parseCaseClauseStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const prefixLength = /^\s*Case\s+/iu.exec(text)?.[0].length;

  if (!prefixLength) {
    return undefined;
  }

  const conditionSlice = text.slice(prefixLength);
  const conditionText = conditionSlice.trim();

  if (conditionText.length === 0) {
    return undefined;
  }

  if (/^Else$/iu.test(conditionText)) {
    return {
      caseKind: "else",
      kind: "caseClauseStatement",
      range: statementRange,
      text
    };
  }

  const leadingWhitespace = conditionSlice.length - conditionSlice.trimStart().length;
  const conditionStartCharacter = prefixLength + leadingWhitespace;

  return {
    caseKind: "value",
    conditionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + conditionStartCharacter,
      inlineCharacterOffset + conditionStartCharacter + conditionText.length
    ),
    conditionText,
    kind: "caseClauseStatement",
    range: statementRange,
    text
  };
}

function parseForEachStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const match = /^\s*For\s+Each\s+(.+?)\s+In\s+(.+)\s*$/iu.exec(text);

  if (!match?.[1] || !match[2]) {
    return undefined;
  }

  const itemPrefixLength = /^\s*For\s+Each\s+/iu.exec(text)?.[0].length ?? 0;
  const itemSlice = match[1];
  const itemText = itemSlice.trim();
  const itemLeadingWhitespace = itemSlice.length - itemSlice.trimStart().length;
  const itemStartCharacter = itemPrefixLength + itemLeadingWhitespace;
  const collectionPrefixLength = text.indexOf(match[2], itemStartCharacter + itemText.length);
  const collectionText = match[2].trim();
  const collectionLeadingWhitespace = match[2].length - match[2].trimStart().length;
  const collectionStartCharacter = collectionPrefixLength + collectionLeadingWhitespace;
  const simpleItemMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)$/u.exec(itemText);

  if (itemText.length === 0 || collectionText.length === 0 || collectionPrefixLength < 0) {
    return undefined;
  }

  return {
    collectionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + collectionStartCharacter,
      inlineCharacterOffset + collectionStartCharacter + collectionText.length
    ),
    collectionText,
    itemName: simpleItemMatch?.[1]?.replace(/[$%&!#@]$/, ""),
    itemRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + itemStartCharacter,
      inlineCharacterOffset + itemStartCharacter + itemText.length
    ),
    itemText,
    kind: "forEachStatement",
    range: statementRange,
    text
  };
}

function parseForStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const match = /^\s*For\s+(.+?)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+?))?\s*$/iu.exec(text);

  if (!match?.[1] || !match[2] || !match[3] || /^Each$/iu.test(match[1].trim())) {
    return undefined;
  }

  const counterText = match[1].trim();
  const startExpressionText = match[2].trim();
  const endExpressionText = match[3].trim();
  const counterStartCharacter = text.indexOf(match[1]);
  const startExpressionStartCharacter = text.indexOf(match[2], counterStartCharacter + match[1].length);
  const endExpressionStartCharacter = text.indexOf(match[3], startExpressionStartCharacter + match[2].length);
  const stepExpressionText = match[4]?.trim();
  const stepExpressionStartCharacter =
    stepExpressionText && match[4] ? text.indexOf(match[4], endExpressionStartCharacter + match[3].length) : -1;
  const simpleCounterMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)$/u.exec(counterText);

  if (
    counterText.length === 0 ||
    startExpressionText.length === 0 ||
    endExpressionText.length === 0 ||
    counterStartCharacter < 0 ||
    startExpressionStartCharacter < 0 ||
    endExpressionStartCharacter < 0
  ) {
    return undefined;
  }

  return {
    counterName: simpleCounterMatch?.[1]?.replace(/[$%&!#@]$/, ""),
    counterRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + counterStartCharacter,
      inlineCharacterOffset + counterStartCharacter + counterText.length
    ),
    counterText,
    endExpressionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + endExpressionStartCharacter,
      inlineCharacterOffset + endExpressionStartCharacter + endExpressionText.length
    ),
    endExpressionText,
    kind: "forStatement",
    range: statementRange,
    startExpressionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + startExpressionStartCharacter,
      inlineCharacterOffset + startExpressionStartCharacter + startExpressionText.length
    ),
    startExpressionText,
    ...(stepExpressionText && stepExpressionStartCharacter >= 0
      ? {
          stepExpressionRange: createInlineOrStatementRange(
            statementRange,
            inlineCharacterOffset + stepExpressionStartCharacter,
            inlineCharacterOffset + stepExpressionStartCharacter + stepExpressionText.length
          ),
          stepExpressionText
        }
      : {}),
    text
  };
}

function parseNextStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const match = /^\s*Next(?:\s+(.+?))?\s*$/iu.exec(text);

  if (!match) {
    return undefined;
  }

  const counterText = match[1]?.trim();
  const counterStartCharacter = counterText ? text.indexOf(match[1] ?? "") : -1;
  const simpleCounterMatch = counterText ? /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)$/u.exec(counterText) : undefined;

  return {
    ...(simpleCounterMatch?.[1] ? { counterName: simpleCounterMatch[1].replace(/[$%&!#@]$/, "") } : {}),
    ...(counterText
      ? {
          counterRange: createInlineOrStatementRange(
            statementRange,
            inlineCharacterOffset + counterStartCharacter,
            inlineCharacterOffset + counterStartCharacter + counterText.length
          ),
          counterText
        }
      : {}),
    kind: "nextStatement",
    range: statementRange,
    text
  };
}

function parseDoBlockStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  if (!/^\s*Do\b/iu.test(text)) {
    return undefined;
  }

  const bareDoMatch = /^\s*Do\s*$/iu.exec(text);

  if (bareDoMatch) {
    return {
      clauseKind: "none",
      kind: "doBlockStatement",
      range: statementRange,
      text
    };
  }

  const clauseMatch = /^\s*Do\s+(While|Until)\s+(.+?)\s*$/iu.exec(text);

  if (!clauseMatch?.[1] || !clauseMatch[2]) {
    return undefined;
  }

  const clauseKind = clauseMatch[1].toLowerCase() as "until" | "while";
  const conditionSlice = clauseMatch[2];
  const conditionText = conditionSlice.trim();
  const conditionStartCharacter = text.indexOf(conditionSlice);

  if (conditionText.length === 0 || conditionStartCharacter < 0) {
    return undefined;
  }

  return {
    clauseKind,
    conditionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + conditionStartCharacter,
      inlineCharacterOffset + conditionStartCharacter + conditionText.length
    ),
    conditionText,
    kind: "doBlockStatement",
    range: statementRange,
    text
  };
}

function parseLoopStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  if (!/^\s*Loop\b/iu.test(text)) {
    return undefined;
  }

  const bareLoopMatch = /^\s*Loop\s*$/iu.exec(text);

  if (bareLoopMatch) {
    return {
      clauseKind: "none",
      kind: "loopStatement",
      range: statementRange,
      text
    };
  }

  const clauseMatch = /^\s*Loop\s+(While|Until)\s+(.+?)\s*$/iu.exec(text);

  if (!clauseMatch?.[1] || !clauseMatch[2]) {
    return undefined;
  }

  const clauseKind = clauseMatch[1].toLowerCase() as "until" | "while";
  const conditionSlice = clauseMatch[2];
  const conditionText = conditionSlice.trim();
  const conditionStartCharacter = text.indexOf(conditionSlice);

  if (conditionText.length === 0 || conditionStartCharacter < 0) {
    return undefined;
  }

  return {
    clauseKind,
    conditionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + conditionStartCharacter,
      inlineCharacterOffset + conditionStartCharacter + conditionText.length
    ),
    conditionText,
    kind: "loopStatement",
    range: statementRange,
    text
  };
}

function parseWhileStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const match = /^\s*While\s+(.+?)\s*$/iu.exec(text);

  if (!match?.[1]) {
    return undefined;
  }

  const conditionSlice = match[1];
  const conditionText = conditionSlice.trim();
  const conditionStartCharacter = text.indexOf(conditionSlice);

  if (conditionText.length === 0 || conditionStartCharacter < 0) {
    return undefined;
  }

  return {
    conditionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + conditionStartCharacter,
      inlineCharacterOffset + conditionStartCharacter + conditionText.length
    ),
    conditionText,
    kind: "whileStatement",
    range: statementRange,
    text
  };
}

function parseWithBlockStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const match = /^\s*With\s+(.+?)\s*$/iu.exec(text);

  if (!match?.[1]) {
    return undefined;
  }

  const targetSlice = match[1];
  const targetText = targetSlice.trim();
  const targetStartCharacter = text.indexOf(targetSlice);

  if (targetText.length === 0 || targetStartCharacter < 0) {
    return undefined;
  }

  return {
    kind: "withBlockStatement",
    range: statementRange,
    targetRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + targetStartCharacter,
      inlineCharacterOffset + targetStartCharacter + targetText.length
    ),
    targetText,
    text
  };
}

function parseOnErrorStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const resumeNextMatch = /^\s*On\s+Error\s+Resume\s+Next\s*$/iu.exec(text);

  if (resumeNextMatch) {
    return {
      actionKind: "resumeNext",
      kind: "onErrorStatement",
      range: statementRange,
      text
    };
  }

  const gotoMatch = /^\s*On\s+Error\s+GoTo\s+(.+?)\s*$/iu.exec(text);

  if (!gotoMatch?.[1]) {
    return undefined;
  }

  const targetSlice = gotoMatch[1];
  const targetText = targetSlice.trim();
  const targetStartCharacter = text.indexOf(targetSlice);

  if (targetText.length === 0 || targetStartCharacter < 0) {
    return undefined;
  }

  return {
    actionKind: "goto",
    kind: "onErrorStatement",
    range: statementRange,
    targetRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + targetStartCharacter,
      inlineCharacterOffset + targetStartCharacter + targetText.length
    ),
    targetText,
    text
  };
}

function parseAssignmentStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const equalsIndex = findAssignmentOperatorIndex(text);

  if (equalsIndex < 0) {
    return undefined;
  }

  const leftText = text.slice(0, equalsIndex);
  const rightText = text.slice(equalsIndex + 1);
  const prefixMatch = /^\s*(Set|Let)\s+/iu.exec(leftText);
  const targetText = leftText.slice(prefixMatch?.[0].length ?? 0).trim();

  if (
    targetText.length === 0 ||
    containsWhitespaceOutsideGrouping(targetText) ||
    /^(?:Call|Case|Do|ElseIf|If|Loop|Next|Select|While|With)\b/iu.test(targetText)
  ) {
    return undefined;
  }

  const expressionText = rightText.trim();

  if (expressionText.length === 0) {
    return undefined;
  }

  const targetStartCharacter = (prefixMatch?.[0].length ?? 0) + leftText.slice(prefixMatch?.[0].length ?? 0).search(/\S/u);
  const expressionStartCharacter = equalsIndex + 1 + (rightText.length - rightText.trimStart().length);
  const simpleTargetMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)$/u.exec(targetText);

  return {
    assignmentKind: prefixMatch?.[1] ? prefixMatch[1].toLowerCase() as "let" | "set" : "implicit",
    expressionRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + expressionStartCharacter,
      inlineCharacterOffset + expressionStartCharacter + expressionText.length
    ),
    expressionText,
    kind: "assignmentStatement",
    range: statementRange,
    targetName: simpleTargetMatch?.[1]?.replace(/[$%&!#@]$/, ""),
    targetRange: createInlineOrStatementRange(
      statementRange,
      inlineCharacterOffset + targetStartCharacter,
      inlineCharacterOffset + targetStartCharacter + targetText.length
    ),
    targetText,
    text
  };
}

function parseCallStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  if (statementRange.start.line !== statementRange.end.line || findAssignmentOperatorIndex(text) >= 0) {
    return undefined;
  }

  const explicitCallMatch = /^\s*Call\s+([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)\s*\((.*)\)\s*$/iu.exec(text);

  if (explicitCallMatch?.[1]) {
    const callPrefixLength = /^\s*Call\s+/iu.exec(text)?.[0].length ?? 0;
    const openParenIndex = text.indexOf("(", callPrefixLength);
    const closeParenIndex = findMatchingCloseParen(text, openParenIndex);

    if (openParenIndex >= 0 && closeParenIndex >= 0) {
      return {
        arguments: splitInvocationArguments(
          text.slice(openParenIndex + 1, closeParenIndex),
          statementRange.start.line,
          inlineCharacterOffset + openParenIndex + 1
        ),
        callStyle: "call",
        kind: "callStatement",
        name: explicitCallMatch[1],
        nameRange: createInlineRange(
          statementRange.start.line,
          inlineCharacterOffset + callPrefixLength,
          inlineCharacterOffset + callPrefixLength + explicitCallMatch[1].length
        ),
        range: statementRange,
        text
      };
    }
  }

  const bareCallMatch = /^\s*([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)(?:\s+(.*\S))?\s*$/u.exec(text);

  if (bareCallMatch?.[1] && bareCallMatch[2] && !isStatementKeyword(bareCallMatch[1])) {
    const leadingWhitespace = /^\s*/u.exec(text)?.[0].length ?? 0;
    const nameStartCharacter = leadingWhitespace;
    const separatorLength = /\s+/u.exec(text.slice(nameStartCharacter + bareCallMatch[1].length))?.[0].length ?? 0;
    const argumentsStartCharacter = nameStartCharacter + bareCallMatch[1].length + separatorLength;

    return {
      arguments: splitInvocationArguments(
        bareCallMatch[2],
        statementRange.start.line,
        inlineCharacterOffset + argumentsStartCharacter
      ),
      callStyle: "bare",
      kind: "callStatement",
      name: bareCallMatch[1],
      nameRange: createInlineRange(
        statementRange.start.line,
        inlineCharacterOffset + nameStartCharacter,
        inlineCharacterOffset + nameStartCharacter + bareCallMatch[1].length
      ),
      range: statementRange,
      text
    };
  }

  const openParenIndex = text.indexOf("(");
  const closeParenIndex = findMatchingCloseParen(text, openParenIndex);
  const identifier = getIdentifierBeforeOpenParen(text, openParenIndex);

  if (!identifier || closeParenIndex < 0 || isStatementKeyword(identifier.text)) {
    return undefined;
  }

  return {
    arguments: splitInvocationArguments(
      text.slice(openParenIndex + 1, closeParenIndex),
      statementRange.start.line,
      inlineCharacterOffset + openParenIndex + 1
    ),
    callStyle: "parenthesized",
    kind: "callStatement",
    name: identifier.text,
    nameRange: createInlineRange(
      statementRange.start.line,
      inlineCharacterOffset + identifier.startCharacter,
      inlineCharacterOffset + identifier.startCharacter + identifier.text.length
    ),
    range: statementRange,
    text
  };
}

function parseProcedureVariableDeclaration(
  source: SourceDocument,
  logicalLine: { codeText: string; endLine: number; startLine: number }
): VariableDeclarationNode {
  const match = /^(Dim|Static)\s+(.+)$/i.exec(logicalLine.codeText.trim());
  const declarators = parseDeclarators(match?.[2] ?? "", source, logicalLine.startLine, logicalLine.endLine);

  return {
    declarators,
    kind: "variableDeclaration",
    range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
    text: logicalLine.codeText
  };
}

function containsWhitespaceOutsideGrouping(text: string): boolean {
  let depth = 0;

  for (let index = 0; index < text.length; index += 1) {
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
      continue;
    }

    if (currentCharacter === ")") {
      depth = Math.max(0, depth - 1);
      continue;
    }

    if (depth === 0 && /\s/u.test(currentCharacter)) {
      return true;
    }
  }

  return false;
}

function splitInvocationArguments(text: string, line: number, baseCharacter: number): Array<{ range: ProcedureStatementNode["range"]; text: string }> {
  const argumentsWithRanges: Array<{ range: ProcedureStatementNode["range"]; text: string }> = [];
  let startIndex = 0;
  let depth = 0;

  for (let index = 0; index < text.length; index += 1) {
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
      continue;
    }

    if (currentCharacter === ")") {
      depth = Math.max(0, depth - 1);
      continue;
    }

    if (currentCharacter === "," && depth === 0) {
      pushArgument(startIndex, index);
      startIndex = index + 1;
    }
  }

  pushArgument(startIndex, text.length);
  return argumentsWithRanges;

  function pushArgument(start: number, end: number): void {
    const rawText = text.slice(start, end);
    const trimmedText = rawText.trim();
    const leadingWhitespace = rawText.length - rawText.trimStart().length;
    const trailingWhitespace = rawText.length - rawText.trimEnd().length;
    const startCharacter = baseCharacter + start + leadingWhitespace;
    const endCharacter = baseCharacter + end - trailingWhitespace;

    argumentsWithRanges.push({
      range: createInlineRange(line, startCharacter, endCharacter),
      text: trimmedText
    });
  }
}

function parseTypeDeclaration(
  source: SourceDocument,
  logicalLines: Array<{ codeText: string; endLine: number; startLine: number }>,
  startIndex: number,
  diagnostics: Diagnostic[]
): { nextIndex: number; node: TypeDeclarationNode } {
  const startLine = logicalLines[startIndex];
  const headerMatch = /^(?:(Public|Private)\s+)?Type\s+([A-Za-z_][A-Za-z0-9_]*)/i.exec(startLine.codeText.trim());
  const members: TypeMemberNode[] = [];
  let endIndex = startIndex + 1;

  while (endIndex < logicalLines.length) {
    const currentLine = logicalLines[endIndex];
    const trimmedText = currentLine.codeText.trim();

    if (/^End\s+Type\b/i.test(trimmedText)) {
      return {
        nextIndex: endIndex + 1,
        node: {
          kind: "typeDeclaration",
          members,
          modifier: headerMatch?.[1],
          name: headerMatch?.[2] ?? startLine.codeText.trim(),
          range: createMappedRange(source, startLine.startLine, 0, currentLine.endLine, source.normalizedLines[currentLine.endLine]?.length ?? 0),
          text: logicalLines.slice(startIndex, endIndex + 1).map((line) => line.codeText).join("\n")
        }
      };
    }

    if (trimmedText.length > 0) {
      const memberMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)(?:\s+As\s+([A-Za-z_][A-Za-z0-9_\.]*))?/i.exec(trimmedText);

      if (memberMatch) {
        members.push({
          kind: "typeMember",
          name: memberMatch[1].replace(/[$%&!#@]$/, ""),
          range: createMappedRange(source, currentLine.startLine, 0, currentLine.endLine, source.normalizedLines[currentLine.endLine]?.length ?? 0),
          typeName: memberMatch[2] ?? typeNameFromSuffix(memberMatch[1])
        });
      }
    }

    endIndex += 1;
  }

  diagnostics.push({
    code: "syntax-error",
    message: "Type declaration is missing End Type.",
    range: createMappedRange(source, startLine.startLine, 0, startLine.endLine, source.normalizedLines[startLine.endLine]?.length ?? 0),
    severity: "error"
  });

  return {
    nextIndex: logicalLines.length,
    node: {
      kind: "typeDeclaration",
      members,
      modifier: headerMatch?.[1],
      name: headerMatch?.[2] ?? startLine.codeText.trim(),
      range: createMappedRange(source, startLine.startLine, 0, logicalLines[logicalLines.length - 1]?.endLine ?? startLine.endLine, source.normalizedLines[logicalLines[logicalLines.length - 1]?.endLine ?? startLine.endLine]?.length ?? 0),
      text: logicalLines.slice(startIndex).map((line) => line.codeText).join("\n")
    }
  };
}

function parseVariableDeclaration(source: SourceDocument, logicalLine: { codeText: string; endLine: number; startLine: number }): VariableDeclarationNode {
  const match = /^(?:(Public|Private)\s+)?Dim\s+(.+)$/i.exec(logicalLine.codeText.trim());
  const declarators = parseDeclarators(match?.[2] ?? "", source, logicalLine.startLine, logicalLine.endLine);

  return {
    declarators,
    kind: "variableDeclaration",
    modifier: match?.[1],
    range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
    text: logicalLine.codeText
  };
}

function parseDeclarators(
  declarationText: string,
  source: SourceDocument,
  startLine: number,
  endLine: number
): VariableDeclaratorNode[] {
  return splitCommaAware(declarationText)
    .map((value) => value.trim())
    .filter((value) => value.length > 0)
    .map((declaratorText) => {
      const match = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)(\(\))?(?:\s+As\s+([A-Za-z_][A-Za-z0-9_\.]*))?/i.exec(declaratorText);
      const rawName = match?.[1] ?? declaratorText;

      return {
        arraySuffix: Boolean(match?.[2]),
        kind: "variableDeclarator",
        name: rawName.replace(/[$%&!#@]$/, ""),
        range: createMappedRange(source, startLine, 0, endLine, source.normalizedLines[endLine]?.length ?? 0),
        typeName: match?.[3] ?? typeNameFromSuffix(rawName)
      };
    });
}

function normalizeProcedureKind(value: string): ProcedureKind {
  const normalized = value.toLowerCase().replace(/\s+/g, " ");

  switch (normalized) {
    case "function":
      return "Function";
    case "property get":
      return "PropertyGet";
    case "property let":
      return "PropertyLet";
    case "property set":
      return "PropertySet";
    default:
      return "Sub";
  }
}

function isProcedureTerminator(text: string, procedureKind: ProcedureKind): boolean {
  switch (procedureKind) {
    case "Function":
      return /^End\s+Function\b/i.test(text);
    case "PropertyGet":
    case "PropertyLet":
    case "PropertySet":
      return /^End\s+Property\b/i.test(text);
    default:
      return /^End\s+Sub\b/i.test(text);
  }
}

function validateProcedureBlocks(source: SourceDocument, procedure: ProcedureDeclarationNode, diagnostics: Diagnostic[]): void {
  const blockStack: Array<{ kind: string; statement: ProcedureStatementNode }> = [];

  for (const statement of procedure.body) {
    const trimmedText = statement.text.trim();

    if (trimmedText.length === 0 || /^#/i.test(trimmedText)) {
      continue;
    }

    if (statement.kind === "ifBlockStatement") {
      blockStack.push({ kind: "if", statement });
      continue;
    }

    if (statement.kind === "elseIfClauseStatement" || statement.kind === "elseClauseStatement") {
      requireCurrentBlock("if", statement, diagnostics);
      continue;
    }

    if (statement.kind === "selectCaseStatement") {
      blockStack.push({ kind: "select", statement });
      continue;
    }

    if (statement.kind === "caseClauseStatement") {
      requireCurrentBlock("select", statement, diagnostics);
      continue;
    }

    if (statement.kind === "forStatement" || statement.kind === "forEachStatement") {
      blockStack.push({ kind: "for", statement });
      continue;
    }

    if (statement.kind === "doBlockStatement") {
      blockStack.push({ kind: "do", statement });
      continue;
    }

    if (statement.kind === "whileStatement") {
      blockStack.push({ kind: "while", statement });
      continue;
    }

    if (statement.kind === "withBlockStatement") {
      blockStack.push({ kind: "with", statement });
      continue;
    }

    if (statement.kind === "endIfStatement") {
      popBlock("if", statement, diagnostics);
      continue;
    }

    if (statement.kind === "endSelectStatement") {
      popBlock("select", statement, diagnostics);
      continue;
    }

    if (statement.kind === "nextStatement") {
      popBlock("for", statement, diagnostics);
      continue;
    }

    if (statement.kind === "loopStatement") {
      popBlock("do", statement, diagnostics);
      continue;
    }

    if (statement.kind === "wendStatement") {
      popBlock("while", statement, diagnostics);
      continue;
    }

    if (statement.kind === "endWithStatement") {
      popBlock("with", statement, diagnostics);
      continue;
    }
  }

  for (const block of blockStack) {
    diagnostics.push({
      code: "syntax-error",
      message: `Missing terminator for ${block.kind} block in ${procedure.name}.`,
      range: block.statement.range,
      severity: "error"
    });
  }

  function popBlock(expectedKind: string, statement: ProcedureStatementNode, sink: Diagnostic[]): void {
    const lastBlock = blockStack.pop();

    if (!lastBlock || lastBlock.kind !== expectedKind) {
      sink.push({
        code: "syntax-error",
        message: `Unexpected block terminator in ${procedure.name}.`,
        range: statement.range,
        severity: "error"
      });

      if (lastBlock) {
        blockStack.push(lastBlock);
      }
    }
  }

  function requireCurrentBlock(expectedKind: string, statement: ProcedureStatementNode, sink: Diagnostic[]): void {
    const lastBlock = blockStack[blockStack.length - 1];

    if (!lastBlock || lastBlock.kind !== expectedKind) {
      sink.push({
        code: "syntax-error",
        message: `Unexpected block clause in ${procedure.name}.`,
        range: statement.range,
        severity: "error"
      });
    }
  }
}

function isIfBlockStart(text: string): boolean {
  return /^If\b.*\bThen\s*$/i.test(text) && !/:/.test(text);
}

function findAssignmentOperatorIndex(text: string): number {
  let depth = 0;

  for (let index = 0; index < text.length; index += 1) {
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
      continue;
    }

    if (currentCharacter === ")") {
      depth = Math.max(0, depth - 1);
      continue;
    }

    if (
      currentCharacter === "=" &&
      depth === 0 &&
      text[index - 1] !== "<" &&
      text[index - 1] !== ">" &&
      text[index + 1] !== ">"
    ) {
      return index;
    }
  }

  return -1;
}

function isStatementKeyword(text: string): boolean {
  return /^(?:Call|Case|Do|Else|ElseIf|End|For|If|Loop|Next|On|Select|While|With)\b/iu.test(text);
}

function getIdentifierBeforeOpenParen(text: string, openParenIndex: number): { startCharacter: number; text: string } | undefined {
  if (openParenIndex <= 0) {
    return undefined;
  }

  let identifierEnd = openParenIndex;

  while (identifierEnd > 0 && /\s/u.test(text[identifierEnd - 1] ?? "")) {
    identifierEnd -= 1;
  }

  let identifierStart = identifierEnd;

  while (identifierStart > 0 && /[A-Za-z0-9_!$%&@#]/u.test(text[identifierStart - 1] ?? "")) {
    identifierStart -= 1;
  }

  const identifierText = text.slice(identifierStart, identifierEnd);

  return /^[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?$/u.test(identifierText)
    ? { startCharacter: identifierStart, text: identifierText }
    : undefined;
}

function findMatchingCloseParen(text: string, openParenIndex: number): number {
  if (openParenIndex < 0 || text[openParenIndex] !== "(") {
    return -1;
  }

  let depth = 0;

  for (let index = openParenIndex; index < text.length; index += 1) {
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
      continue;
    }

    if (currentCharacter === ")") {
      depth -= 1;

      if (depth === 0) {
        return index;
      }
    }
  }

  return -1;
}

function skipStringLiteral(text: string, startIndex: number): number {
  let index = startIndex + 1;

  while (index < text.length) {
    if (text[index] === "\"" && text[index + 1] === "\"") {
      index += 2;
      continue;
    }

    if (text[index] === "\"") {
      return index;
    }

    index += 1;
  }

  return text.length - 1;
}

function skipDateLiteral(text: string, startIndex: number): number {
  let index = startIndex + 1;

  while (index < text.length) {
    if (text[index] === "#") {
      return index;
    }

    index += 1;
  }

  return text.length - 1;
}

function createInlineRange(line: number, startCharacter: number, endCharacter: number): ProcedureStatementNode["range"] {
  return {
    end: {
      character: endCharacter,
      line
    },
    start: {
      character: startCharacter,
      line
    }
  };
}

function createInlineOrStatementRange(
  statementRange: ProcedureStatementNode["range"],
  startCharacter: number,
  endCharacter: number
): ProcedureStatementNode["range"] {
  return statementRange.start.line === statementRange.end.line
    ? createInlineRange(statementRange.start.line, startCharacter, endCharacter)
    : statementRange;
}
