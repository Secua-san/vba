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
  LinePosition,
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
import { buildLogicalLines, hasStatementSeparatorColon, splitCodeAndComment, splitCommaAware } from "./text";

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

function parseConstDeclaration(
  source: SourceDocument,
  logicalLine: { codeText: string; endLine: number; startLine: number },
  createSegmentRange?: (startCharacter: number, endCharacter: number) => ConstDeclarationNode["range"]
): ConstDeclarationNode {
  const match = /^(?:(Public|Private)\s+)?Const\s+([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)(?:\s+As\s+([A-Za-z_][A-Za-z0-9_\.]*))?/i.exec(
    logicalLine.codeText.trim()
  );
  const text = logicalLine.codeText.trim();
  const value = parseConstValue(source, logicalLine, text, createSegmentRange);

  return {
    kind: "constDeclaration",
    modifier: match?.[1],
    name: match?.[2] ?? text,
    range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
    text: logicalLine.codeText,
    typeName: match?.[3] ?? typeNameFromSuffix(match?.[2] ?? ""),
    valueRange: value?.range,
    valueText: value?.text
  };
}

function parseConstValue(
  source: SourceDocument,
  logicalLine: { codeText: string; endLine: number; startLine: number },
  text: string,
  createSegmentRange?: (startCharacter: number, endCharacter: number) => ConstDeclarationNode["range"]
): { range: ConstDeclarationNode["range"]; text: string } | undefined {
  const equalsIndex = findAssignmentOperatorIndex(text);

  if (equalsIndex < 0) {
    return undefined;
  }

  const rawValueText = text.slice(equalsIndex + 1);
  const valueText = rawValueText.trim();
  const valueStartCharacter = findTrimmedSliceStart(text, rawValueText, equalsIndex + 1);

  if (valueText.length === 0 || valueStartCharacter < 0) {
    return undefined;
  }

  if (createSegmentRange) {
    return {
      range: createSegmentRange(valueStartCharacter, valueStartCharacter + valueText.length),
      text: valueText
    };
  }

  if (logicalLine.startLine !== logicalLine.endLine) {
    return {
      range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
      text: valueText
    };
  }

  const lineText = source.normalizedLines[logicalLine.startLine] ?? "";
  const textStartCharacter = lineText.indexOf(text);
  const inlineCharacterOffset = textStartCharacter >= 0 ? textStartCharacter : lineText.length - lineText.trimStart().length;

  return {
    range: createInlineRange(logicalLine.startLine, inlineCharacterOffset + valueStartCharacter, inlineCharacterOffset + valueStartCharacter + valueText.length),
    text: valueText
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
  const statementRange = createMappedRange(
    source,
    logicalLine.startLine,
    0,
    logicalLine.endLine,
    source.normalizedLines[logicalLine.endLine]?.length ?? 0
  );
  const statementText = createProcedureStatementText(source, logicalLine, statementRange);
  const trimmedText = statementText.text;
  const inlineCharacterOffset = statementText.inlineCharacterOffset;
  const leadingLabel = parseLeadingLabel(trimmedText);
  const parsedText = leadingLabel?.statementText ?? trimmedText;
  const parsedLogicalLine = leadingLabel ? { ...logicalLine, codeText: parsedText } : logicalLine;
  const parsedInlineCharacterOffset = inlineCharacterOffset + (leadingLabel?.statementStartCharacter ?? 0);
  const createParsedRange = leadingLabel
    ? (startCharacter: number, endCharacter: number) =>
        statementText.createRange(
          leadingLabel.statementStartCharacter + startCharacter,
          leadingLabel.statementStartCharacter + endCharacter
        )
    : statementText.createRange;

  if (parsedText.length === 0) {
    return undefined;
  }

  if (/^Const\b/i.test(parsedText)) {
    const constant = parseConstDeclaration(source, parsedLogicalLine, createParsedRange);
    return {
      declaredConstants: [constant],
      kind: "constStatement",
      range: statementRange,
      text: trimmedText
    };
  }

  if (/^(?:Dim|Static)\b/i.test(parsedText)) {
    const declaration = parseProcedureVariableDeclaration(source, parsedLogicalLine);
    return {
      declaredVariables: declaration.declarators,
      kind: "declarationStatement",
      range: statementRange,
      text: trimmedText
    };
  }

  const structuredBlockStatement = parseStructuredBlockStatement(parsedText, statementRange, parsedInlineCharacterOffset);

  if (structuredBlockStatement) {
    return structuredBlockStatement.text === trimmedText
      ? structuredBlockStatement
      : {
          ...structuredBlockStatement,
          text: trimmedText
        };
  }

  const assignmentStatement = parseAssignmentStatement(
    parsedText,
    statementRange,
    parsedInlineCharacterOffset,
    createParsedRange
  );

  if (assignmentStatement) {
    return assignmentStatement.text === trimmedText
      ? assignmentStatement
      : {
          ...assignmentStatement,
          text: trimmedText
        };
  }

  const callStatement = parseCallStatement(
    parsedText,
    statementRange,
    parsedInlineCharacterOffset,
    createParsedRange
  );

  if (callStatement) {
    return callStatement.text === trimmedText
      ? callStatement
      : {
          ...callStatement,
          text: trimmedText
        };
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

  const goToStatement = parseGoToStatement(text, statementRange, inlineCharacterOffset);

  if (goToStatement) {
    return goToStatement;
  }

  const resumeStatement = parseResumeStatement(text, statementRange, inlineCharacterOffset);

  if (resumeStatement) {
    return resumeStatement;
  }

  const terminationStatement = parseTerminationStatement(text, statementRange);

  if (terminationStatement) {
    return terminationStatement;
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
  if (!/^ElseIf\b.*\bThen\s*$/iu.test(text) || hasStatementSeparatorColon(text)) {
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
  const counterPrefixLength = /^\s*For\s+/iu.exec(text)?.[0].length ?? 0;
  const counterStartCharacter = findTrimmedSliceStart(text, match[1], counterPrefixLength);
  const startExpressionStartCharacter = findTrimmedSliceStart(
    text,
    match[2],
    counterStartCharacter + counterText.length
  );
  const endExpressionStartCharacter = findTrimmedSliceStart(
    text,
    match[3],
    startExpressionStartCharacter + startExpressionText.length
  );
  const stepExpressionText = match[4]?.trim();
  const stepExpressionStartCharacter =
    stepExpressionText && match[4]
      ? findTrimmedSliceStart(text, match[4], endExpressionStartCharacter + endExpressionText.length)
      : -1;
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
  const counterPrefixLength = /^\s*Next\s+/iu.exec(text)?.[0].length ?? 0;
  const counterStartCharacter = counterText && match[1] ? findTrimmedSliceStart(text, match[1], counterPrefixLength) : -1;
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
  const conditionPrefixLength = /^\s*Do\s+(?:While|Until)\s+/iu.exec(text)?.[0].length ?? 0;
  const conditionStartCharacter = findTrimmedSliceStart(text, conditionSlice, conditionPrefixLength);

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
  const conditionPrefixLength = /^\s*Loop\s+(?:While|Until)\s+/iu.exec(text)?.[0].length ?? 0;
  const conditionStartCharacter = findTrimmedSliceStart(text, conditionSlice, conditionPrefixLength);

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
  const conditionPrefixLength = /^\s*While\s+/iu.exec(text)?.[0].length ?? 0;
  const conditionStartCharacter = findTrimmedSliceStart(text, conditionSlice, conditionPrefixLength);

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
  const targetPrefixLength = /^\s*With\s+/iu.exec(text)?.[0].length ?? 0;
  const targetStartCharacter = findTrimmedSliceStart(text, targetSlice, targetPrefixLength);

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
  const targetPrefixLength = /^\s*On\s+Error\s+GoTo\s+/iu.exec(text)?.[0].length ?? 0;
  const targetStartCharacter = findTrimmedSliceStart(text, targetSlice, targetPrefixLength);

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

function parseGoToStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  const match = /^\s*(GoTo|GoSub)\s+(.+?)\s*$/iu.exec(text);

  if (!match?.[1] || !match[2]) {
    return undefined;
  }

  const targetSlice = match[2];
  const targetText = targetSlice.trim();
  const targetPrefixLength = /^\s*(?:GoTo|GoSub)\s+/iu.exec(text)?.[0].length ?? 0;
  const targetStartCharacter = findTrimmedSliceStart(text, targetSlice, targetPrefixLength);

  if (targetText.length === 0 || targetStartCharacter < 0) {
    return undefined;
  }

  return {
    actionKind: /^GoSub$/iu.test(match[1]) ? "goSub" : "goTo",
    kind: "goToStatement",
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

function parseResumeStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0
): ProcedureStatementNode | undefined {
  if (/^\s*Resume\s*$/iu.test(text)) {
    return {
      actionKind: "implicit",
      kind: "resumeStatement",
      range: statementRange,
      text
    };
  }

  if (/^\s*Resume\s+Next\s*$/iu.test(text)) {
    return {
      actionKind: "next",
      kind: "resumeStatement",
      range: statementRange,
      text
    };
  }

  const targetMatch = /^\s*Resume\s+(.+?)\s*$/iu.exec(text);

  if (!targetMatch?.[1]) {
    return undefined;
  }

  const targetSlice = targetMatch[1];
  const targetText = targetSlice.trim();
  const targetPrefixLength = /^\s*Resume\s+/iu.exec(text)?.[0].length ?? 0;
  const targetStartCharacter = findTrimmedSliceStart(text, targetSlice, targetPrefixLength);

  if (targetText.length === 0 || targetStartCharacter < 0) {
    return undefined;
  }

  return {
    actionKind: "target",
    kind: "resumeStatement",
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

function parseTerminationStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"]
): ProcedureStatementNode | undefined {
  if (/^\s*End\s*$/iu.test(text)) {
    return {
      kind: "endStatement",
      range: statementRange,
      text
    };
  }

  // Loop exits (`Exit For` / `Exit Do`) are not procedure termination statements.
  const exitMatch = /^\s*Exit\s+(Sub|Function|Property)\s*$/iu.exec(text);

  if (!exitMatch?.[1]) {
    return undefined;
  }

  const normalizedExitKind = exitMatch[1].toLowerCase();
  const exitKind =
    normalizedExitKind === "function" ? "Function" : normalizedExitKind === "property" ? "Property" : "Sub";

  return {
    exitKind,
    kind: "exitStatement",
    range: statementRange,
    text
  };
}

function parseAssignmentStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0,
  createSegmentRange?: (startCharacter: number, endCharacter: number) => ProcedureStatementNode["range"]
): ProcedureStatementNode | undefined {
  const equalsIndex = findAssignmentOperatorIndex(text);

  if (equalsIndex < 0) {
    return undefined;
  }

  const createRange =
    createSegmentRange ??
    ((startCharacter: number, endCharacter: number) =>
      createInlineRange(
        statementRange.start.line,
        inlineCharacterOffset + startCharacter,
        inlineCharacterOffset + endCharacter
      ));

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
    expressionRange: createRange(expressionStartCharacter, expressionStartCharacter + expressionText.length),
    expressionText,
    kind: "assignmentStatement",
    range: statementRange,
    targetName: simpleTargetMatch?.[1]?.replace(/[$%&!#@]$/, ""),
    targetRange: createRange(targetStartCharacter, targetStartCharacter + targetText.length),
    targetText,
    text
  };
}

function parseCallStatement(
  text: string,
  statementRange: ProcedureStatementNode["range"],
  inlineCharacterOffset = 0,
  createSegmentRange?: (startCharacter: number, endCharacter: number) => ProcedureStatementNode["range"]
): ProcedureStatementNode | undefined {
  if (findAssignmentOperatorIndex(text) >= 0) {
    return undefined;
  }

  const createRange =
    createSegmentRange ??
    ((startCharacter: number, endCharacter: number) =>
      createInlineRange(
        statementRange.start.line,
        inlineCharacterOffset + startCharacter,
        inlineCharacterOffset + endCharacter
      ));

  const explicitCallMatch = /^\s*Call\s+([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)\s*\((.*)\)\s*$/iu.exec(text);

  if (explicitCallMatch?.[1]) {
    const callPrefixLength = /^\s*Call\s+/iu.exec(text)?.[0].length ?? 0;
    const openParenIndex = text.indexOf("(", callPrefixLength);
    const closeParenIndex = findMatchingCloseParen(text, openParenIndex);

    if (openParenIndex >= 0 && closeParenIndex >= 0) {
      return {
        arguments: splitInvocationArguments(
          text.slice(openParenIndex + 1, closeParenIndex),
          createRange,
          openParenIndex + 1
        ),
        callStyle: "call",
        kind: "callStatement",
        name: explicitCallMatch[1],
        nameRange: createRange(callPrefixLength, callPrefixLength + explicitCallMatch[1].length),
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
        createRange,
        argumentsStartCharacter
      ),
      callStyle: "bare",
      kind: "callStatement",
      name: bareCallMatch[1],
      nameRange: createRange(nameStartCharacter, nameStartCharacter + bareCallMatch[1].length),
      range: statementRange,
      text
    };
  }

  const parenthesizedCallMatch = /^\s*([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)\s*\((.*)\)\s*$/u.exec(text);
  const openParenIndex = parenthesizedCallMatch ? text.indexOf("(", /^\s*/u.exec(text)?.[0].length ?? 0) : -1;
  const closeParenIndex = findMatchingCloseParen(text, openParenIndex);
  const identifier =
    parenthesizedCallMatch && openParenIndex >= 0
      ? {
          startCharacter: /^\s*/u.exec(text)?.[0].length ?? 0,
          text: parenthesizedCallMatch[1]
        }
      : undefined;

  if (!identifier || closeParenIndex < 0 || isStatementKeyword(identifier.text)) {
    return undefined;
  }

  return {
    arguments: splitInvocationArguments(
      text.slice(openParenIndex + 1, closeParenIndex),
      createRange,
      openParenIndex + 1
    ),
    callStyle: "parenthesized",
    kind: "callStatement",
    name: identifier.text,
    nameRange: createRange(identifier.startCharacter, identifier.startCharacter + identifier.text.length),
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

function splitInvocationArguments(
  text: string,
  createRange: (startCharacter: number, endCharacter: number) => ProcedureStatementNode["range"],
  baseCharacter: number
): Array<{ range: ProcedureStatementNode["range"]; text: string }> {
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
      range: createRange(startCharacter, endCharacter),
      text: trimmedText
    });
  }
}

function createProcedureStatementText(
  source: SourceDocument,
  logicalLine: { codeText: string; endLine: number; startLine: number },
  statementRange: ProcedureStatementNode["range"]
): {
  createRange: (startCharacter: number, endCharacter: number) => ProcedureStatementNode["range"];
  inlineCharacterOffset: number;
  text: string;
} {
  if (logicalLine.startLine === logicalLine.endLine) {
    const lineText = source.normalizedLines[logicalLine.startLine] ?? "";
    const inlineCharacterOffset = lineText.length - lineText.trimStart().length;

    return {
      createRange: (startCharacter: number, endCharacter: number) =>
        createInlineRange(
          statementRange.start.line,
          inlineCharacterOffset + startCharacter,
          inlineCharacterOffset + endCharacter
        ),
      inlineCharacterOffset,
      text: logicalLine.codeText.trim()
    };
  }

  const flattened = flattenProcedureStatementText(source, logicalLine);

  return {
    createRange: (startCharacter: number, endCharacter: number) =>
      mapFlattenedProcedureStatementRange(flattened.positions, startCharacter, endCharacter, statementRange),
    inlineCharacterOffset: 0,
    text: flattened.text
  };
}

function flattenProcedureStatementText(
  source: SourceDocument,
  logicalLine: { endLine: number; startLine: number }
): { positions: Array<LinePosition | undefined>; text: string } {
  let flattenedText = "";
  let hasOutput = false;
  const positions: Array<LinePosition | undefined> = [];

  for (let lineIndex = logicalLine.startLine; lineIndex <= logicalLine.endLine; lineIndex += 1) {
    const rawLine = source.normalizedLines[lineIndex];

    if (rawLine === undefined) {
      continue;
    }

    const { code } = splitCodeAndComment(rawLine);
    const continued = /\s+_\s*$/.test(code);
    const codeWithoutContinuation = continued ? code.replace(/\s+_\s*$/, "") : code;
    const trimmedCode = codeWithoutContinuation.trimEnd();
    const leadingTrimLength = hasOutput ? trimmedCode.length - trimmedCode.trimStart().length : 0;
    const emittedText = hasOutput ? trimmedCode.trimStart() : trimmedCode;
    const mappedLine = source.lineMap[lineIndex] ?? lineIndex;

    if (hasOutput) {
      flattenedText += " ";
      positions.push(undefined);
    }

    for (let index = 0; index < emittedText.length; index += 1) {
      flattenedText += emittedText[index];
      positions.push({
        character: leadingTrimLength + index,
        line: mappedLine
      });
    }

    hasOutput = true;
  }

  const trimmedStartIndex = flattenedText.search(/\S/u);

  if (trimmedStartIndex < 0) {
    return {
      positions: [],
      text: ""
    };
  }

  const trimmedEndIndex = flattenedText.trimEnd().length;

  return {
    positions: positions.slice(trimmedStartIndex, trimmedEndIndex),
    text: flattenedText.slice(trimmedStartIndex, trimmedEndIndex)
  };
}

function mapFlattenedProcedureStatementRange(
  positions: Array<LinePosition | undefined>,
  startCharacter: number,
  endCharacter: number,
  fallbackRange: ProcedureStatementNode["range"]
): ProcedureStatementNode["range"] {
  const mappedPositions = positions.slice(startCharacter, Math.max(endCharacter, startCharacter + 1))
    .filter((position): position is LinePosition => Boolean(position));
  const startPosition = mappedPositions[0];
  const endPosition = mappedPositions[mappedPositions.length - 1];

  if (!startPosition || !endPosition) {
    return fallbackRange;
  }

  return {
    end: {
      character: endPosition.character + (endCharacter > startCharacter ? 1 : 0),
      line: endPosition.line
    },
    start: startPosition
  };
}

function parseLeadingLabel(text: string): { statementStartCharacter: number; statementText: string } | undefined {
  const match = /^(?:[A-Za-z_][A-Za-z0-9_]*|\d+):\s*/u.exec(text);
  const statementText = match ? text.slice(match[0].length).trimStart() : "";

  if (!match || statementText.length === 0) {
    return undefined;
  }

  return {
    statementStartCharacter: match[0].length,
    statementText
  };
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
  const blockStack: Array<
    | { kind: "do" | "for" | "while" | "with"; statement: ProcedureStatementNode }
    | { hasElseClause: boolean; kind: "if"; statement: ProcedureStatementNode }
    | { hasCaseElseClause: boolean; kind: "select"; statement: ProcedureStatementNode }
  > = [];

  for (const statement of procedure.body) {
    const trimmedText = statement.text.trim();

    if (trimmedText.length === 0 || /^#/i.test(trimmedText)) {
      continue;
    }

    if (statement.kind === "ifBlockStatement") {
      blockStack.push({ hasElseClause: false, kind: "if", statement });
      continue;
    }

    if (statement.kind === "elseIfClauseStatement" || statement.kind === "elseClauseStatement") {
      const currentIfBlock = requireCurrentBlock("if", statement, diagnostics);

      if (currentIfBlock?.kind === "if") {
        if (currentIfBlock.hasElseClause) {
          pushUnexpectedBlockClauseDiagnostic(statement, diagnostics);
        } else if (statement.kind === "elseClauseStatement") {
          currentIfBlock.hasElseClause = true;
        }
      }

      continue;
    }

    if (statement.kind === "selectCaseStatement") {
      blockStack.push({ hasCaseElseClause: false, kind: "select", statement });
      continue;
    }

    if (statement.kind === "caseClauseStatement") {
      const currentSelectBlock = requireCurrentBlock("select", statement, diagnostics);

      if (currentSelectBlock?.kind === "select") {
        if (statement.caseKind === "else") {
          if (currentSelectBlock.hasCaseElseClause) {
            pushUnexpectedBlockClauseDiagnostic(statement, diagnostics);
          } else {
            currentSelectBlock.hasCaseElseClause = true;
          }
        } else if (currentSelectBlock.hasCaseElseClause) {
          pushUnexpectedBlockClauseDiagnostic(statement, diagnostics);
        }
      }

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
      return;
    }

    if (expectedKind === "for" && statement.kind === "nextStatement") {
      const activeCounterName = getForBlockCounterName(lastBlock.statement);

      if (activeCounterName && statement.counterName && normalizeIdentifier(activeCounterName) !== normalizeIdentifier(statement.counterName)) {
        sink.push({
          code: "syntax-error",
          message: `Next counter '${statement.counterText}' does not match active loop variable '${activeCounterName}' in ${procedure.name}.`,
          range: statement.range,
          severity: "error"
        });
      }
    }
  }

  function requireCurrentBlock(expectedKind: string, statement: ProcedureStatementNode, sink: Diagnostic[]) {
    const lastBlock = blockStack[blockStack.length - 1];

    if (!lastBlock || lastBlock.kind !== expectedKind) {
      pushUnexpectedBlockClauseDiagnostic(statement, sink);
      return undefined;
    }

    return lastBlock;
  }

  function pushUnexpectedBlockClauseDiagnostic(statement: ProcedureStatementNode, sink: Diagnostic[]): void {
    sink.push({
      code: "syntax-error",
      message: `Unexpected block clause in ${procedure.name}.`,
      range: statement.range,
      severity: "error"
    });
  }

  function getForBlockCounterName(statement: ProcedureStatementNode): string | undefined {
    switch (statement.kind) {
      case "forStatement":
        return statement.counterName;
      case "forEachStatement":
        return statement.itemName;
      default:
        return undefined;
    }
  }
}

function isIfBlockStart(text: string): boolean {
  return /^If\b.*\bThen\s*$/i.test(text) && !hasStatementSeparatorColon(text);
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
  return /^(?:Call|Case|Do|Else|ElseIf|End|Exit|For|GoSub|GoTo|If|Loop|Next|On|Resume|Select|While|With)\b/iu.test(text);
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

function findTrimmedSliceStart(text: string, rawSlice: string, searchStartCharacter: number): number {
  const rawIndex = text.indexOf(rawSlice, searchStartCharacter);

  if (rawIndex < 0) {
    return -1;
  }

  return rawIndex + (rawSlice.length - rawSlice.trimStart().length);
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
