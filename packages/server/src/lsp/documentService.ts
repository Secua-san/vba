import {
  analyzeModule,
  areTypesCompatible,
  extractIdentifierAtPosition,
  findDefinition,
  getCompletionSymbols,
  getDocumentOutline,
  inferExpressionTypeAtLine,
  getSymbolTypeName,
  normalizeIdentifier,
  removeStringAndDateLiterals,
  splitCodeAndComment,
  type AnalysisResult,
  type Diagnostic,
  type LinePosition,
  type SourceRange,
  type SymbolInfo
} from "../../../core/src/index";

export interface DocumentState {
  analysis: AnalysisResult;
  languageId: string;
  text: string;
  uri: string;
  version: number;
}

export interface WorkspaceSymbolResolution {
  moduleName: string;
  symbol: SymbolInfo;
  typeName?: string;
  uri: string;
}

export interface WorkspaceReference {
  range: SourceRange;
  uri: string;
}

export interface SignatureParameterHint {
  documentation?: string;
  label: string;
}

export interface SignatureHint {
  activeParameter?: number;
  activeSignature: number;
  documentation?: string;
  label: string;
  parameters: SignatureParameterHint[];
}

export interface DocumentService {
  analyzeText: (uri: string, languageId: string, version: number, text: string) => DocumentState;
  getCompletionSymbols: (uri: string, position: LinePosition) => WorkspaceSymbolResolution[];
  getDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined;
  getDiagnostics: (uri: string) => Diagnostic[];
  getDocumentSymbols: (uri: string) => ReturnType<typeof getDocumentOutline>;
  getReferences: (uri: string, position: LinePosition, includeDeclaration: boolean) => WorkspaceReference[];
  getSignatureHelp: (uri: string, position: LinePosition) => SignatureHint | undefined;
  getState: (uri: string) => DocumentState | undefined;
  remove: (uri: string) => void;
}

export function createDocumentService(): DocumentService {
  const documentStates = new Map<string, DocumentState>();
  let workspaceIndex = createWorkspaceIndex([]);

  function resolveDefinition(uri: string, position: LinePosition): WorkspaceSymbolResolution | undefined {
    const state = documentStates.get(uri);

    if (!state) {
      return undefined;
    }

    const localDefinition = findDefinition(state.analysis, position);

    if (localDefinition) {
      return createResolution(state, localDefinition, uri);
    }

    const identifier = extractIdentifierAtPosition(state.text.replace(/\r\n?/g, "\n"), position);

    if (!identifier) {
      return undefined;
    }

    const matches = workspaceIndex.byNormalizedName
      .get(normalizeIdentifier(identifier))
      ?.filter((resolution) => resolution.uri !== uri) ?? [];

    return matches.length === 1 ? matches[0] : undefined;
  }

  function getFilteredDiagnostics(uri: string): Diagnostic[] {
    const state = documentStates.get(uri);

    if (!state) {
      return [];
    }

    return state.analysis.diagnostics.filter((diagnostic) => {
      if (diagnostic.code !== "undeclared-variable") {
        return true;
      }

      const identifier = getDiagnosticIdentifier(state.text, diagnostic);

      if (!identifier) {
        return true;
      }

      const matches = workspaceIndex.byNormalizedName
        .get(normalizeIdentifier(identifier))
        ?.filter((resolution) => resolution.uri !== uri && resolution.symbol.kind !== "module") ?? [];

      return matches.length !== 1;
    });
  }

  function getReferenceMatches(uri: string, position: LinePosition, includeDeclaration: boolean): WorkspaceReference[] {
    const target = resolveDefinition(uri, position);

    if (!target) {
      return [];
    }

    const references = includeDeclaration
      ? [
          {
            range: getDeclarationRange(documentStates.get(target.uri), target, resolveDefinition),
            uri: target.uri
          }
        ]
      : [];

    for (const state of documentStates.values()) {
      references.push(...collectReferencesForState(state, target, resolveDefinition));
    }

    return deduplicateReferences(references);
  }

  return {
    analyzeText(uri: string, languageId: string, version: number, text: string): DocumentState {
      const analysis = analyzeModule(text, { fileName: getFileNameFromUri(uri) });
      const state: DocumentState = {
        analysis,
        languageId,
        text,
        uri,
        version
      };
      documentStates.set(uri, state);
      workspaceIndex = createWorkspaceIndex([...documentStates.values()]);
      return state;
    },
    getCompletionSymbols(uri: string, position: LinePosition): WorkspaceSymbolResolution[] {
      const state = documentStates.get(uri);

      if (!state) {
        return [];
      }

      const localSymbols = getCompletionSymbols(state.analysis, position).map((symbol) => ({
        ...createResolution(state, symbol, uri)
      }));
      const deduplicated = new Map<string, WorkspaceSymbolResolution>();

      for (const resolution of localSymbols) {
        deduplicated.set(`${resolution.uri}:${resolution.symbol.kind}:${resolution.symbol.normalizedName}`, resolution);
      }

      for (const resolution of workspaceIndex.entries) {
        if (resolution.uri === uri) {
          continue;
        }

        const key = `${resolution.uri}:${resolution.symbol.kind}:${resolution.symbol.normalizedName}`;

        if (!deduplicated.has(key)) {
          deduplicated.set(key, resolution);
        }
      }

      return narrowCompletionByAssignmentTarget(state.text, uri, position, [...deduplicated.values()], resolveDefinition);
    },
    getDefinition(uri: string, position: LinePosition): WorkspaceSymbolResolution | undefined {
      return resolveDefinition(uri, position);
    },
    getDiagnostics(uri: string): Diagnostic[] {
      return getFilteredDiagnostics(uri);
    },
    getDocumentSymbols(uri: string) {
      const state = documentStates.get(uri);
      return state ? getDocumentOutline(state.analysis) : [];
    },
    getReferences(uri: string, position: LinePosition, includeDeclaration: boolean): WorkspaceReference[] {
      return getReferenceMatches(uri, position, includeDeclaration);
    },
    getSignatureHelp(uri: string, position: LinePosition): SignatureHint | undefined {
      const state = documentStates.get(uri);

      if (!state) {
        return undefined;
      }

      const callContext = getCallContext(state.text, position);

      if (!callContext) {
        return undefined;
      }

      const target = resolveDefinition(uri, {
        character: callContext.identifierStartCharacter,
        line: position.line
      });

      if (!target || !isCallableSymbol(target.symbol)) {
        return undefined;
      }

      const targetState = documentStates.get(target.uri);

      if (!targetState) {
        return undefined;
      }

      const callable = findCallableMember(targetState.analysis, target.symbol);

      if (!callable) {
        return undefined;
      }

      return createSignatureHint(state.analysis, uri, target, callable, position.line, callContext, resolveDefinition);
    },
    getState(uri: string): DocumentState | undefined {
      return documentStates.get(uri);
    },
    remove(uri: string): void {
      documentStates.delete(uri);
      workspaceIndex = createWorkspaceIndex([...documentStates.values()]);
    }
  };
}

interface WorkspaceIndex {
  byNormalizedName: Map<string, WorkspaceSymbolResolution[]>;
  entries: WorkspaceSymbolResolution[];
}

type CallableMember = Extract<AnalysisResult["module"]["members"][number], { kind: "declareStatement" | "procedureDeclaration" }>;

interface CallContext {
  activeParameter: number;
  currentArgumentStartCharacter: number;
  currentArgumentText: string;
  identifierStartCharacter: number;
}

function createWorkspaceIndex(states: DocumentState[]): WorkspaceIndex {
  const entries = states.flatMap(collectWorkspaceSymbols);
  const byNormalizedName = new Map<string, WorkspaceSymbolResolution[]>();

  for (const entry of entries) {
    const currentEntries = byNormalizedName.get(entry.symbol.normalizedName);

    if (currentEntries) {
      currentEntries.push(entry);
    } else {
      byNormalizedName.set(entry.symbol.normalizedName, [entry]);
    }
  }

  return {
    byNormalizedName,
    entries
  };
}

function collectWorkspaceSymbols(state: DocumentState): WorkspaceSymbolResolution[] {
  const entries: WorkspaceSymbolResolution[] = [
    createResolution(state, state.analysis.symbols.moduleSymbol, state.uri)
  ];

  if (state.analysis.source.moduleKind !== "standard") {
    return entries;
  }

  const moduleSymbolsByName = new Map<string, SymbolInfo[]>();

  for (const symbol of state.analysis.symbols.moduleSymbols) {
    const currentEntries = moduleSymbolsByName.get(symbol.normalizedName);

    if (currentEntries) {
      currentEntries.push(symbol);
    } else {
      moduleSymbolsByName.set(symbol.normalizedName, [symbol]);
    }
  }

  for (const member of state.analysis.module.members) {
    switch (member.kind) {
      case "constDeclaration":
        if (isWorkspaceVisible(member.modifier)) {
          entries.push(...findModuleSymbols(moduleSymbolsByName, member.name, state));
        }
        break;
      case "declareStatement":
        if (isWorkspaceVisible(member.modifier)) {
          entries.push(...findModuleSymbols(moduleSymbolsByName, member.name, state));
        }
        break;
      case "enumDeclaration":
        if (isWorkspaceVisible(member.modifier)) {
          entries.push(...findModuleSymbols(moduleSymbolsByName, member.name, state));

          for (const enumMember of member.members) {
            entries.push(...findModuleSymbols(moduleSymbolsByName, enumMember.name, state));
          }
        }
        break;
      case "procedureDeclaration":
        if (isWorkspaceVisible(member.modifier)) {
          entries.push(...findModuleSymbols(moduleSymbolsByName, member.name, state));
        }
        break;
      case "typeDeclaration":
        if (isWorkspaceVisible(member.modifier)) {
          entries.push(...findModuleSymbols(moduleSymbolsByName, member.name, state));
        }
        break;
      case "variableDeclaration":
        if (isWorkspaceVisible(member.modifier)) {
          for (const declarator of member.declarators) {
            entries.push(...findModuleSymbols(moduleSymbolsByName, declarator.name, state));
          }
        }
        break;
      default:
        break;
    }
  }

  return deduplicateWorkspaceEntries(entries);
}

function collectReferencesForState(
  state: DocumentState,
  target: WorkspaceSymbolResolution,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined
): WorkspaceReference[] {
  const lines = state.text.replace(/\r\n?/g, "\n").split("\n");
  const references: WorkspaceReference[] = [];

  for (const member of state.analysis.module.members) {
    if (member.kind !== "procedureDeclaration") {
      continue;
    }

    for (const statement of member.body) {
      if (statement.kind !== "executableStatement") {
        continue;
      }

      for (let lineIndex = statement.range.start.line; lineIndex <= statement.range.end.line; lineIndex += 1) {
        const line = lines[lineIndex];

        if (line === undefined) {
          continue;
        }

        const { code } = splitCodeAndComment(line);
        const scrubbed = removeStringAndDateLiterals(code);

        for (const match of scrubbed.matchAll(/[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?/g)) {
          const identifier = match[0];
          const normalizedIdentifier = normalizeIdentifier(identifier);
          const startIndex = match.index ?? 0;
          const nextCharacter = scrubbed[startIndex + identifier.length] ?? "";

          if (normalizedIdentifier !== target.symbol.normalizedName) {
            continue;
          }

          if (nextCharacter === ":" && startIndex === 0) {
            continue;
          }

          const resolution = resolveDefinition(state.uri, { character: startIndex, line: lineIndex });

          if (resolution && isSameResolution(resolution, target)) {
            references.push({
              range: {
                start: {
                  character: startIndex,
                  line: lineIndex
                },
                end: {
                  character: startIndex + identifier.length,
                  line: lineIndex
                }
              },
              uri: state.uri
            });
          }
        }
      }
    }
  }

  return references;
}

function deduplicateWorkspaceEntries(entries: WorkspaceSymbolResolution[]): WorkspaceSymbolResolution[] {
  const deduplicated = new Map<string, WorkspaceSymbolResolution>();

  for (const entry of entries) {
    const key = `${entry.uri}:${entry.symbol.kind}:${entry.symbol.normalizedName}`;

    if (!deduplicated.has(key)) {
      deduplicated.set(key, entry);
    }
  }

  return [...deduplicated.values()];
}

function deduplicateReferences(references: WorkspaceReference[]): WorkspaceReference[] {
  const deduplicated = new Map<string, WorkspaceReference>();

  for (const reference of references) {
    const key = `${reference.uri}:${reference.range.start.line}:${reference.range.start.character}:${reference.range.end.line}:${reference.range.end.character}`;

    if (!deduplicated.has(key)) {
      deduplicated.set(key, reference);
    }
  }

  return [...deduplicated.values()];
}

function narrowCompletionByAssignmentTarget(
  text: string,
  uri: string,
  position: LinePosition,
  completions: WorkspaceSymbolResolution[],
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined
): WorkspaceSymbolResolution[] {
  const targetTypeName = getAssignmentTargetTypeName(text, uri, position, resolveDefinition);

  if (!targetTypeName) {
    return completions;
  }

  const narrowed = completions.filter(
    (resolution) => !resolution.typeName || areTypesCompatible(targetTypeName, resolution.typeName)
  );

  return narrowed.length > 0 && narrowed.length < completions.length ? narrowed : completions;
}

function createSignatureHint(
  analysis: AnalysisResult,
  sourceUri: string,
  target: WorkspaceSymbolResolution,
  callable: CallableMember,
  line: number,
  callContext: CallContext,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined
): SignatureHint {
  const currentArgumentTypeName = getCurrentArgumentTypeName(analysis, sourceUri, line, callContext, resolveDefinition);
  const activeParameter =
    callable.parameters.length === 0 ? undefined : Math.min(callContext.activeParameter, callable.parameters.length - 1);

  return {
    activeParameter,
    activeSignature: 0,
    documentation: target.moduleName === analysis.module.name ? undefined : `${target.moduleName} モジュール`,
    label: buildSignatureLabel(target.symbol.name, callable),
    parameters: callable.parameters.map((parameter, index) => ({
      documentation: buildParameterDocumentation(parameter, index === activeParameter ? currentArgumentTypeName : undefined),
      label: buildParameterLabel(parameter)
    }))
  };
}

function getCurrentArgumentTypeName(
  analysis: AnalysisResult,
  uri: string,
  line: number,
  callContext: CallContext,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined
): string | undefined {
  const expressionText = callContext.currentArgumentText.trim();

  if (expressionText.length === 0) {
    return undefined;
  }

  const inferredTypeName = inferExpressionTypeAtLine(analysis, line, expressionText);

  if (inferredTypeName) {
    return inferredTypeName;
  }

  const expressionOffset = callContext.currentArgumentText.length - callContext.currentArgumentText.trimStart().length;
  const simpleReferenceMatch = /^([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)(?:\s*\(.*\))?$/iu.exec(expressionText);

  if (!simpleReferenceMatch?.[1]) {
    return undefined;
  }

  return resolveDefinition(uri, {
    character: callContext.currentArgumentStartCharacter + expressionOffset,
    line
  })?.typeName;
}

function getCallContext(text: string, position: LinePosition): CallContext | undefined {
  const line = text.replace(/\r\n?/g, "\n").split("\n")[position.line];

  if (line === undefined) {
    return undefined;
  }

  const { code } = splitCodeAndComment(line.slice(0, position.character));
  const frames: Array<{
    commaCount: number;
    identifier?: string;
    identifierStartCharacter?: number;
    lastCommaIndex: number;
  }> = [];
  let index = 0;

  while (index < code.length) {
    const character = code[index];

    if (character === "\"") {
      index = skipStringLiteral(code, index);
      continue;
    }

    if (character === "#") {
      index = skipDateLiteral(code, index);
      continue;
    }

    if (character === "(") {
      const identifier = getIdentifierBeforeOpenParen(code, index);

      frames.push({
        commaCount: 0,
        identifier: identifier?.text,
        identifierStartCharacter: identifier?.startCharacter,
        lastCommaIndex: index
      });
      index += 1;
      continue;
    }

    if (character === ",") {
      const currentFrame = frames[frames.length - 1];

      if (currentFrame) {
        currentFrame.commaCount += 1;
        currentFrame.lastCommaIndex = index;
      }

      index += 1;
      continue;
    }

    if (character === ")") {
      frames.pop();
    }

    index += 1;
  }

  const currentFrame = [...frames].reverse().find((frame) => frame.identifier && frame.identifierStartCharacter !== undefined);

  if (!currentFrame?.identifier || currentFrame.identifierStartCharacter === undefined) {
    return undefined;
  }

  return {
    activeParameter: currentFrame.commaCount,
    currentArgumentStartCharacter: currentFrame.lastCommaIndex + 1,
    currentArgumentText: code.slice(currentFrame.lastCommaIndex + 1),
    identifierStartCharacter: currentFrame.identifierStartCharacter
  };
}

function findCallableMember(analysis: AnalysisResult, symbol: SymbolInfo): CallableMember | undefined {
  return analysis.module.members.find((member): member is CallableMember => {
    if ((member.kind !== "declareStatement" && member.kind !== "procedureDeclaration") || member.name !== symbol.name) {
      return false;
    }

    const selectionRange = member.kind === "procedureDeclaration" ? member.headerRange : member.range;

    return (
      selectionRange.start.line === symbol.selectionRange.start.line &&
      selectionRange.start.character === symbol.selectionRange.start.character &&
      selectionRange.end.line === symbol.selectionRange.end.line &&
      selectionRange.end.character === symbol.selectionRange.end.character
    );
  });
}

function isCallableSymbol(symbol: SymbolInfo): boolean {
  return symbol.kind === "declare" || symbol.kind === "procedure";
}

function buildSignatureLabel(name: string, callable: CallableMember): string {
  const parameters = callable.parameters.map(buildParameterLabel).join(", ");
  const returnType = callable.kind === "declareStatement" ? callable.returnType : callable.returnType;
  const procedureKind = callable.kind === "declareStatement" ? callable.procedureKind : callable.procedureKind;

  if (procedureKind === "Sub") {
    return `${name}(${parameters})`;
  }

  return `${name}(${parameters}) As ${returnType ?? "Variant"}`;
}

function buildParameterLabel(parameter: CallableMember["parameters"][number]): string {
  const modifiers = [
    parameter.isOptional ? "Optional" : "",
    parameter.isParamArray ? "ParamArray" : "",
    parameter.direction === "byVal" ? "ByVal" : "ByRef"
  ]
    .filter((value) => value.length > 0)
    .join(" ");
  const arraySuffix = parameter.arraySuffix ? "()" : "";
  const typeSuffix = parameter.typeName ? ` As ${parameter.typeName}` : "";

  return `${modifiers} ${parameter.name}${arraySuffix}${typeSuffix}`.trim();
}

function buildParameterDocumentation(
  parameter: CallableMember["parameters"][number],
  currentArgumentTypeName?: string
): string | undefined {
  const lines = [];

  if (parameter.typeName) {
    lines.push(`想定型: ${parameter.typeName}`);
  }

  if (currentArgumentTypeName) {
    lines.push(`現在の引数型: ${currentArgumentTypeName}`);
  }

  return lines.length > 0 ? lines.join("\n") : undefined;
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

function findModuleSymbols(
  moduleSymbolsByName: Map<string, SymbolInfo[]>,
  name: string,
  state: DocumentState
): WorkspaceSymbolResolution[] {
  return (moduleSymbolsByName.get(normalizeIdentifier(name)) ?? []).map((symbol) => createResolution(state, symbol, state.uri));
}

function getTextInRange(text: string, range: Diagnostic["range"]): string {
  const normalizedText = text.replace(/\r\n?/g, "\n");
  const lines = normalizedText.split("\n");
  const line = lines[range.start.line];

  if (line === undefined || range.start.line !== range.end.line) {
    return "";
  }

  return line.slice(range.start.character, range.end.character);
}

function getDiagnosticIdentifier(text: string, diagnostic: Diagnostic): string {
  const messageMatch = diagnostic.message.match(/'([^']+)'/);

  if (messageMatch?.[1]) {
    return messageMatch[1];
  }

  return getTextInRange(text, diagnostic.range);
}

function getAssignmentTargetTypeName(
  text: string,
  uri: string,
  position: LinePosition,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined
): string | undefined {
  const line = text.replace(/\r\n?/g, "\n").split("\n")[position.line];

  if (line === undefined) {
    return undefined;
  }

  const beforeCursor = line.slice(0, position.character);
  const { code } = splitCodeAndComment(beforeCursor);
  const assignmentTarget = parseAssignmentTarget(code, position.line);

  if (!assignmentTarget) {
    return undefined;
  }

  return resolveDefinition(uri, assignmentTarget)?.typeName;
}

function parseAssignmentTarget(code: string, line: number): LinePosition | undefined {
  const match = /^\s*(?:Set\s+)?([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)\s*=\s*(?:[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)?$/iu.exec(code);

  if (!match?.[1]) {
    return undefined;
  }

  const prefixMatch = /^\s*(?:Set\s+)?/iu.exec(code);
  const identifierStart = prefixMatch?.[0].length ?? 0;

  return {
    character: identifierStart,
    line
  };
}

function getFileNameFromUri(uri: string): string | undefined {
  const normalizedUri = uri.startsWith("file:///") ? decodeURIComponent(uri.replace("file:///", "")) : uri;
  const segments = normalizedUri.split(/[\\/]/);
  return segments[segments.length - 1];
}

function createResolution(state: DocumentState, symbol: SymbolInfo, uri: string): WorkspaceSymbolResolution {
  return {
    moduleName: state.analysis.module.name,
    symbol,
    typeName: getSymbolTypeName(state.analysis, symbol),
    uri
  };
}

function getDeclarationRange(
  state: DocumentState | undefined,
  target: WorkspaceSymbolResolution,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined
): SourceRange {
  if (!state) {
    return target.symbol.selectionRange;
  }

  const lines = state.text.replace(/\r\n?/g, "\n").split("\n");

  for (let lineIndex = target.symbol.selectionRange.start.line; lineIndex <= target.symbol.selectionRange.end.line; lineIndex += 1) {
    const line = lines[lineIndex];

    if (line === undefined) {
      continue;
    }

    const { code } = splitCodeAndComment(line);
    const scrubbed = removeStringAndDateLiterals(code);

    for (const match of scrubbed.matchAll(/[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?/g)) {
      const identifier = match[0];
      const startIndex = match.index ?? 0;

      if (normalizeIdentifier(identifier) !== target.symbol.normalizedName) {
        continue;
      }

      if (lineIndex === target.symbol.selectionRange.start.line && startIndex < target.symbol.selectionRange.start.character) {
        continue;
      }

      if (
        lineIndex === target.symbol.selectionRange.end.line &&
        startIndex + identifier.length > target.symbol.selectionRange.end.character
      ) {
        continue;
      }

      const range = {
        start: {
          character: startIndex,
          line: lineIndex
        },
        end: {
          character: startIndex + identifier.length,
          line: lineIndex
        }
      };

      const resolution = resolveDefinition(state.uri, range.start);

      if (!resolution || !isSameResolution(resolution, target)) {
        continue;
      }

      return range;
    }
  }

  return target.symbol.selectionRange;
}

function isSameResolution(left: WorkspaceSymbolResolution, right: WorkspaceSymbolResolution): boolean {
  return (
    left.uri === right.uri &&
    left.symbol.kind === right.symbol.kind &&
    left.symbol.normalizedName === right.symbol.normalizedName &&
    left.symbol.selectionRange.start.line === right.symbol.selectionRange.start.line &&
    left.symbol.selectionRange.start.character === right.symbol.selectionRange.start.character &&
    left.symbol.selectionRange.end.line === right.symbol.selectionRange.end.line &&
    left.symbol.selectionRange.end.character === right.symbol.selectionRange.end.character
  );
}

function isWorkspaceVisible(modifier?: string): boolean {
  return /^(public|friend)$/i.test(modifier ?? "");
}
