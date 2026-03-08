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

  return {
    kind: "executableStatement",
    range: createMappedRange(source, logicalLine.startLine, 0, logicalLine.endLine, source.normalizedLines[logicalLine.endLine]?.length ?? 0),
    text: trimmedText
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

    if (isIfBlockStart(trimmedText)) {
      blockStack.push({ kind: "if", statement });
      continue;
    }

    if (/^Select\s+Case\b/i.test(trimmedText)) {
      blockStack.push({ kind: "select", statement });
      continue;
    }

    if (/^For\b/i.test(trimmedText)) {
      blockStack.push({ kind: "for", statement });
      continue;
    }

    if (/^Do\b/i.test(trimmedText)) {
      blockStack.push({ kind: "do", statement });
      continue;
    }

    if (/^While\b/i.test(trimmedText)) {
      blockStack.push({ kind: "while", statement });
      continue;
    }

    if (/^With\b/i.test(trimmedText)) {
      blockStack.push({ kind: "with", statement });
      continue;
    }

    if (/^End\s+If\b/i.test(trimmedText)) {
      popBlock("if", statement, diagnostics);
      continue;
    }

    if (/^End\s+Select\b/i.test(trimmedText)) {
      popBlock("select", statement, diagnostics);
      continue;
    }

    if (/^Next\b/i.test(trimmedText)) {
      popBlock("for", statement, diagnostics);
      continue;
    }

    if (/^Loop\b/i.test(trimmedText)) {
      popBlock("do", statement, diagnostics);
      continue;
    }

    if (/^Wend\b/i.test(trimmedText)) {
      popBlock("while", statement, diagnostics);
      continue;
    }

    if (/^End\s+With\b/i.test(trimmedText)) {
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
}

function isIfBlockStart(text: string): boolean {
  return /^If\b.*\bThen\s*$/i.test(text) && !/:/.test(text);
}
