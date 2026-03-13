import {
  analyzeModule,
  areTypesCompatible,
  collectByRefArgumentDiagnostics,
  extractIdentifierAtPosition,
  findDefinition,
  formatModuleIndentation,
  getBuiltinMemberSignature,
  getBuiltinCompletionItems,
  getBuiltinMemberCompletionItems,
  getBuiltinMemberReferenceItem,
  getBuiltinReferenceItem,
  getCompletionSymbols,
  getDocumentOutline,
  inferExpressionTypeAtLine,
  isReservedOrBuiltinIdentifier,
  getSymbolTypeName,
  normalizeIdentifier,
  removeStringAndDateLiterals,
  resolveBuiltinMemberOwner,
  splitCodeAndComment,
  type AnalysisResult,
  type BuiltinCallableSignature,
  type BuiltinCompletionKind,
  type BuiltinMemberReferenceItem,
  type BuiltinReferenceItem,
  type BuiltinSemanticModifier,
  type BuiltinSemanticType,
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
  completionItemKind?: BuiltinCompletionKind;
  documentation?: string;
  isBuiltIn?: boolean;
  moduleName: string;
  semanticModifiers?: BuiltinSemanticModifier[];
  semanticType?: BuiltinSemanticType;
  symbol: SymbolInfo;
  typeName?: string;
  uri: string;
}

export interface WorkspaceReference {
  range: SourceRange;
  uri: string;
}

export interface RenameTarget {
  placeholder: string;
  range: SourceRange;
}

export interface RenameTextEdit {
  newText: string;
  range: SourceRange;
  uri: string;
}

export interface DocumentCodeAction {
  edit: RenameTextEdit;
  kind: "quickfix";
  title: string;
}

export interface HoverHint {
  contents: string;
  range?: SourceRange;
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

export const SEMANTIC_TOKEN_TYPES = ["variable", "parameter", "function", "type", "enumMember", "keyword"] as const;
export const SEMANTIC_TOKEN_MODIFIERS = ["declaration", "readonly"] as const;

export type SemanticTokenTypeName = (typeof SEMANTIC_TOKEN_TYPES)[number];
export type SemanticTokenModifierName = (typeof SEMANTIC_TOKEN_MODIFIERS)[number];

export interface SemanticTokenEntry {
  modifiers: SemanticTokenModifierName[];
  range: SourceRange;
  type: SemanticTokenTypeName;
}

export interface DocumentService {
  analyzeText: (uri: string, languageId: string, version: number, text: string) => DocumentState;
  formatDocument: (uri: string, options?: { insertSpaces?: boolean; tabSize?: number }) => string | undefined;
  getCodeActions: (uri: string) => DocumentCodeAction[];
  getCompletionSymbols: (uri: string, position: LinePosition) => WorkspaceSymbolResolution[];
  getDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined;
  getDiagnostics: (uri: string) => Diagnostic[];
  getDocumentSymbols: (uri: string) => ReturnType<typeof getDocumentOutline>;
  getHover: (uri: string, position: LinePosition) => HoverHint | undefined;
  getRenameEdits: (uri: string, position: LinePosition, newName: string) => RenameTextEdit[] | undefined;
  getReferences: (uri: string, position: LinePosition, includeDeclaration: boolean) => WorkspaceReference[];
  getSemanticTokens: (uri: string) => SemanticTokenEntry[];
  prepareRename: (uri: string, position: LinePosition) => RenameTarget | undefined;
  getSignatureHelp: (uri: string, position: LinePosition) => SignatureHint | undefined;
  getState: (uri: string) => DocumentState | undefined;
  remove: (uri: string) => void;
}

export function createDocumentService(): DocumentService {
  const documentStates = new Map<string, DocumentState>();
  let workspaceIndex = createWorkspaceIndex([]);
  const getDocumentState = (uri: string): DocumentState | undefined => documentStates.get(uri);

  function resolveLocalRenameTarget(uri: string, position: LinePosition): {
    range: SourceRange;
    resolution: WorkspaceSymbolResolution;
    state: DocumentState;
    scope: LocalProcedureScope;
  } | undefined {
    const state = documentStates.get(uri);

    if (!state) {
      return undefined;
    }

    const resolution = resolveDefinition(uri, position);
    const range = getIdentifierRangeAtPosition(state.text, position);

    if (
      !resolution ||
      !range ||
      resolution.uri !== uri ||
      resolution.symbol.scope !== "procedure" ||
      resolution.symbol.kind !== "variable"
    ) {
      return undefined;
    }

    const scope = state.analysis.symbols.procedureScopes.find((item) =>
      item.symbols.some((symbol) => isSameSymbol(symbol, resolution.symbol))
    );

    return scope ? { range, resolution, scope, state } : undefined;
  }

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

    const workspaceByRefDiagnostics = collectByRefArgumentDiagnostics(state.analysis, (position) => {
      const target = resolveDefinition(uri, position);

      if (!target || target.uri === uri) {
        return undefined;
      }

      const targetState = documentStates.get(target.uri);

      if (!targetState) {
        return undefined;
      }

      const callable = findCallableMember(targetState.analysis, target.symbol);
      return callable ? { callable, symbol: target.symbol } : undefined;
    });
    const diagnostics = deduplicateDiagnostics([...state.analysis.diagnostics, ...workspaceByRefDiagnostics]);

    return diagnostics.filter((diagnostic) => {
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
    formatDocument(uri: string, options?: { insertSpaces?: boolean; tabSize?: number }): string | undefined {
      const state = documentStates.get(uri);

      if (!state) {
        return undefined;
      }

      return formatModuleIndentation(state.text, {
        fileName: getFileNameFromUri(uri),
        indentSize: options?.tabSize,
        insertSpaces: options?.insertSpaces
      });
    },
    getCodeActions(uri: string): DocumentCodeAction[] {
      const state = documentStates.get(uri);

      if (!state) {
        return [];
      }

      const optionExplicitAction = createOptionExplicitCodeAction(state);
      return optionExplicitAction ? [optionExplicitAction] : [];
    },
    getCompletionSymbols(uri: string, position: LinePosition): WorkspaceSymbolResolution[] {
      const state = documentStates.get(uri);

      if (!state) {
        return [];
      }

      const completionContext = getCompletionContext(state.text, position);
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

      const userAndWorkspaceCompletions = filterCompletionsByPrefix([...deduplicated.values()], completionContext.prefix);

      if (completionContext.isMemberAccess) {
        const memberOwnerName = resolveConfirmedBuiltinMemberOwner(
          state,
          position.line,
          completionContext,
          resolveDefinition,
          getDocumentState
        );

        return memberOwnerName
          ? getBuiltinMemberCompletionItems(memberOwnerName, completionContext.prefix).map(createBuiltinResolution)
          : [];
      }

      const builtInCompletions =
        completionContext.prefix.length > 0
          ? getBuiltinCompletionItems(completionContext.prefix)
              .filter((item) => !userAndWorkspaceCompletions.some((resolution) => resolution.symbol.normalizedName === item.normalizedName))
              .map(createBuiltinResolution)
          : [];

      return narrowCompletionByAssignmentTarget(
        state.text,
        uri,
        position,
        [...userAndWorkspaceCompletions, ...builtInCompletions],
        resolveDefinition
      );
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
    getHover(uri: string, position: LinePosition): HoverHint | undefined {
      const state = documentStates.get(uri);

      if (!state) {
        return undefined;
      }

      return getBuiltinMemberHover(state, uri, position, resolveDefinition, getDocumentState);
    },
    getRenameEdits(uri: string, position: LinePosition, newName: string): RenameTextEdit[] | undefined {
      const renameTarget = resolveLocalRenameTarget(uri, position);

      if (!renameTarget || !isValidRenameIdentifier(newName) || hasRenameConflict(renameTarget, newName)) {
        return undefined;
      }

      return getReferenceMatches(uri, position, true)
        .filter(
          (reference) =>
            reference.uri === renameTarget.resolution.uri &&
            rangeIsWithin(renameTarget.scope.procedure.range, reference.range)
        )
        .map((reference) => ({
          newText: newName,
          range: reference.range,
          uri: reference.uri
        }));
    },
    getReferences(uri: string, position: LinePosition, includeDeclaration: boolean): WorkspaceReference[] {
      return getReferenceMatches(uri, position, includeDeclaration);
    },
    getSemanticTokens(uri: string): SemanticTokenEntry[] {
      const state = documentStates.get(uri);
      return state ? collectSemanticTokensForState(state, resolveDefinition, documentStates) : [];
    },
    prepareRename(uri: string, position: LinePosition): RenameTarget | undefined {
      const renameTarget = resolveLocalRenameTarget(uri, position);

      return renameTarget
        ? {
            placeholder: renameTarget.resolution.symbol.name,
            range: renameTarget.range
          }
        : undefined;
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

      const builtinMember = resolveBuiltinCallableMember(uri, position.line, callContext, resolveDefinition, getDocumentState);

      if (builtinMember) {
        return createBuiltinSignatureHint(state.analysis, uri, position.line, callContext, builtinMember, resolveDefinition);
      }

      const target = resolveDefinition(uri, {
        character: callContext.identifierStartCharacter,
        line: position.line
      });

      if (target && isCallableSymbol(target.symbol)) {
        const targetState = documentStates.get(target.uri);

        if (!targetState) {
          return undefined;
        }

        const callable = findCallableMember(targetState.analysis, target.symbol);

        if (!callable) {
          return undefined;
        }

        return createSignatureHint(state.analysis, uri, target, callable, position.line, callContext, resolveDefinition);
      }

      return undefined;
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

type LocalProcedureScope = DocumentState["analysis"]["symbols"]["procedureScopes"][number];
type CallableMember = Extract<AnalysisResult["module"]["members"][number], { kind: "declareStatement" | "procedureDeclaration" }>;
type SemanticTokenShape = Pick<SemanticTokenEntry, "modifiers" | "type">;

interface CallContext {
  activeParameter: number;
  callPath: string[];
  callPathStartCharacter: number;
  currentArgumentStartCharacter: number;
  currentArgumentText: string;
  identifierStartCharacter: number;
}

interface CompletionContext {
  isMemberAccess: boolean;
  memberPath: string[];
  memberPathStartCharacter?: number;
  prefix: string;
}

const OPTION_EXPLICIT_TITLE = "Option Explicit を追加";

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

function createBuiltinResolution(item: BuiltinMemberReferenceItem | BuiltinReferenceItem): WorkspaceSymbolResolution {
  return {
    completionItemKind: item.completionKind,
    documentation: item.documentation,
    isBuiltIn: true,
    moduleName: item.detail,
    semanticModifiers: item.modifiers,
    semanticType: item.semanticType,
    symbol: {
      kind: mapBuiltinSymbolKind(item.completionKind),
      name: item.name,
      normalizedName: item.normalizedName,
      range: createZeroRange(),
      scope: "module",
      selectionRange: createZeroRange(),
      typeName: item.typeName
    },
    typeName: item.typeName,
    uri: "builtin://reference"
  };
}

function createZeroRange(): SourceRange {
  return {
    start: {
      character: 0,
      line: 0
    },
    end: {
      character: 0,
      line: 0
    }
  };
}

function createOptionExplicitCodeAction(state: DocumentState): DocumentCodeAction | undefined {
  if (hasOptionExplicit(state)) {
    return undefined;
  }

  const edit = createOptionExplicitEdit(state);

  return edit
    ? {
        edit,
        kind: "quickfix",
        title: OPTION_EXPLICIT_TITLE
      }
    : undefined;
}

function createOptionExplicitEdit(state: DocumentState): RenameTextEdit | undefined {
  const originalLines = state.analysis.source.originalLines;
  const codeStartLine = state.analysis.source.lineMap[0] ?? originalLines.length;
  const eol = getDocumentEol(state.text);
  let anchoredInsertionLine: number | undefined;
  let insertionLine = codeStartLine;

  for (let lineIndex = codeStartLine; lineIndex < originalLines.length; lineIndex += 1) {
    const line = originalLines[lineIndex] ?? "";
    const trimmedLine = line.trim();

    if (trimmedLine.length === 0 || isFullLineComment(trimmedLine)) {
      continue;
    }

    if (/^Attribute\s+VB_/iu.test(trimmedLine) || /^Option\b/iu.test(trimmedLine)) {
      anchoredInsertionLine = lineIndex + 1;
      continue;
    }

    insertionLine = lineIndex;
    break;
  }

  insertionLine = anchoredInsertionLine ?? insertionLine;

  if (insertionLine >= originalLines.length) {
    const lastLineIndex = Math.max(0, originalLines.length - 1);
    const lastLineText = originalLines[lastLineIndex] ?? "";
    const insertAtDocumentEnd = state.text.length === 0 ? "" : lastLineText.length > 0 ? eol : "";

    return {
      newText: `${insertAtDocumentEnd}Option Explicit${eol}`,
      range: {
        start: {
          character: lastLineText.length,
          line: lastLineIndex
        },
        end: {
          character: lastLineText.length,
          line: lastLineIndex
        }
      },
      uri: state.uri
    };
  }

  const nextLine = originalLines[insertionLine] ?? "";

  return {
    newText: `Option Explicit${eol}${nextLine.trim().length > 0 ? eol : ""}`,
    range: {
      start: {
        character: 0,
        line: insertionLine
      },
      end: {
        character: 0,
        line: insertionLine
      }
    },
    uri: state.uri
  };
}

function getDocumentEol(text: string): "\n" | "\r\n" {
  return text.includes("\r\n") ? "\r\n" : "\n";
}

function hasOptionExplicit(state: DocumentState): boolean {
  return state.analysis.module.members.some(
    (member) => member.kind === "optionStatement" && normalizeIdentifier(member.name) === "explicit"
  );
}

function isFullLineComment(trimmedLine: string): boolean {
  return trimmedLine.startsWith("'") || /^Rem\b/iu.test(trimmedLine);
}

function collectSemanticTokensForState(
  state: DocumentState,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined,
  documentStates: ReadonlyMap<string, DocumentState>
): SemanticTokenEntry[] {
  const tokens = new Map<string, SemanticTokenEntry>();
  const getDocumentState = (uri: string): DocumentState | undefined => documentStates.get(uri);
  const declarationResolutions = [
    ...state.analysis.symbols.moduleSymbols.map((symbol) => createResolution(state, symbol, state.uri)),
    ...state.analysis.symbols.procedureScopes.flatMap((scope) => scope.symbols.map((symbol) => createResolution(state, symbol, state.uri)))
  ];
  const lines = state.text.replace(/\r\n?/g, "\n").split("\n");

  for (const resolution of declarationResolutions) {
    const tokenShape = mapSemanticToken(resolution.symbol);

    if (!tokenShape) {
      continue;
    }

    const declarationRange = getDeclarationRange(documentStates.get(resolution.uri), resolution, resolveDefinition);
    addSemanticToken(tokens, declarationRange, {
      modifiers: addUniqueModifier(tokenShape.modifiers, "declaration"),
      type: tokenShape.type
    });
  }

  for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
    const line = lines[lineIndex];
    const { code } = splitCodeAndComment(line);
    const scrubbed = removeStringAndDateLiterals(code);

    for (const match of scrubbed.matchAll(/[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?/g)) {
      const identifier = match[0];
      const startCharacter = match.index ?? 0;
      const previousCharacter = scrubbed[startCharacter - 1] ?? "";
      const nextCharacter = scrubbed[startCharacter + identifier.length] ?? "";

      if (nextCharacter === ":" && startCharacter === 0) {
        continue;
      }

      const range = {
        start: {
          character: startCharacter,
          line: lineIndex
        },
        end: {
          character: startCharacter + identifier.length,
          line: lineIndex
        }
      };

      if (previousCharacter === ".") {
        const memberTokenShape = mapBuiltinMemberSemanticToken(
          state.uri,
          lineIndex,
          scrubbed,
          startCharacter,
          identifier,
          resolveDefinition,
          getDocumentState
        );

        if (memberTokenShape) {
          addSemanticToken(tokens, range, memberTokenShape);
        }

        continue;
      }

      const resolution = resolveDefinition(state.uri, range.start);
      const builtinAliasTokenShape =
        resolution && isBuiltinAliasDocumentModule(resolution, identifier, getDocumentState)
          ? mapBuiltinSemanticToken(identifier)
          : undefined;
      const tokenShape = builtinAliasTokenShape
        ? builtinAliasTokenShape
        : resolution
          ? mapSemanticToken(resolution.symbol, resolution.semanticType, resolution.semanticModifiers)
          : mapBuiltinSemanticToken(identifier);

      if (!tokenShape) {
        continue;
      }

      const declarationRange = resolution && !builtinAliasTokenShape
        ? getDeclarationRange(documentStates.get(resolution.uri), resolution, resolveDefinition)
        : undefined;
      const isDeclaration =
        resolution && declarationRange ? resolution.uri === state.uri && rangesEqual(range, declarationRange) : false;

      addSemanticToken(tokens, range, {
        modifiers: isDeclaration ? addUniqueModifier(tokenShape.modifiers, "declaration") : tokenShape.modifiers,
        type: tokenShape.type
      });
    }
  }

  return [...tokens.values()].sort(compareSemanticTokens);
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

function deduplicateDiagnostics(diagnostics: Diagnostic[]): Diagnostic[] {
  const deduplicated = new Map<string, Diagnostic>();

  for (const diagnostic of diagnostics) {
    const key = `${diagnostic.code}:${diagnostic.range.start.line}:${diagnostic.range.start.character}:${diagnostic.range.end.line}:${diagnostic.range.end.character}:${diagnostic.message}`;

    if (!deduplicated.has(key)) {
      deduplicated.set(key, diagnostic);
    }
  }

  return [...deduplicated.values()];
}

function addSemanticToken(
  tokens: Map<string, SemanticTokenEntry>,
  range: SourceRange,
  token: SemanticTokenShape
): void {
  const key = `${range.start.line}:${range.start.character}:${range.end.line}:${range.end.character}:${token.type}:${token.modifiers.join(".")}`;

  if (!tokens.has(key)) {
    tokens.set(key, {
      modifiers: token.modifiers,
      range,
      type: token.type
    });
  }
}

function addUniqueModifier(
  modifiers: SemanticTokenModifierName[],
  modifier: SemanticTokenModifierName
): SemanticTokenModifierName[] {
  return modifiers.includes(modifier) ? modifiers : [...modifiers, modifier];
}

function hasRenameConflict(
  renameTarget: {
    resolution: WorkspaceSymbolResolution;
    state: DocumentState;
    scope: LocalProcedureScope;
  },
  newName: string
): boolean {
  const normalizedName = normalizeIdentifier(newName);
  const accessibleSymbols = [
    renameTarget.state.analysis.symbols.moduleSymbol,
    ...renameTarget.state.analysis.symbols.moduleSymbols,
    ...renameTarget.scope.symbols
  ];

  return accessibleSymbols.some(
    (symbol) => symbol.normalizedName === normalizedName && !isSameSymbol(symbol, renameTarget.resolution.symbol)
  );
}

function getIdentifierRangeAtPosition(text: string, position: LinePosition): SourceRange | undefined {
  const line = text.replace(/\r\n?/g, "\n").split("\n")[position.line];

  if (line === undefined) {
    return undefined;
  }

  for (const match of line.matchAll(/[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?/g)) {
    const startCharacter = match.index ?? 0;
    const endCharacter = startCharacter + match[0].length;

    if (position.character < startCharacter || position.character > endCharacter) {
      continue;
    }

    return {
      start: {
        character: startCharacter,
        line: position.line
      },
      end: {
        character: endCharacter,
        line: position.line
      }
    };
  }

  return undefined;
}

function getCompletionContext(text: string, position: LinePosition): CompletionContext {
  const line = text.replace(/\r\n?/g, "\n").split("\n")[position.line] ?? "";
  const { code } = splitCodeAndComment(line.slice(0, position.character));
  const memberAccess = parseTrailingMemberAccess(code);

  if (memberAccess) {
    return {
      isMemberAccess: true,
      memberPath: memberAccess.memberPath,
      memberPathStartCharacter: memberAccess.memberPathStartCharacter,
      prefix: memberAccess.prefix
    };
  }

  return {
    isMemberAccess: false,
    memberPath: [],
    memberPathStartCharacter: undefined,
    prefix: getTrailingIdentifier(code)
  };
}

function resolveConfirmedBuiltinMemberOwner(
  state: DocumentState,
  line: number,
  completionContext: CompletionContext,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined,
  getDocumentState: (uri: string) => DocumentState | undefined
): string | undefined {
  if (
    !completionContext.isMemberAccess ||
    completionContext.memberPath.length === 0 ||
    completionContext.memberPathStartCharacter === undefined
  ) {
    return undefined;
  }

  const rootResolution = resolveDefinition(state.uri, {
    character: completionContext.memberPathStartCharacter,
    line
  });

  if (
    rootResolution &&
    !isBuiltinAliasDocumentModule(rootResolution, stripIndexedAccessMarker(completionContext.memberPath[0]), getDocumentState)
  ) {
    return undefined;
  }

  return resolveBuiltinMemberOwner(completionContext.memberPath);
}

function isValidRenameIdentifier(name: string): boolean {
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/u.test(name)) {
    return false;
  }

  return !isReservedOrBuiltinIdentifier(name);
}

function rangeIsWithin(outer: SourceRange, inner: SourceRange): boolean {
  return comparePositions(outer.start, inner.start) <= 0 && comparePositions(outer.end, inner.end) >= 0;
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

function filterCompletionsByPrefix(
  completions: WorkspaceSymbolResolution[],
  prefix: string
): WorkspaceSymbolResolution[] {
  if (prefix.length === 0) {
    return completions;
  }

  const normalizedPrefix = normalizeIdentifier(prefix);
  return completions.filter((resolution) => resolution.symbol.normalizedName.startsWith(normalizedPrefix));
}

function resolveBuiltinCallableMember(
  uri: string,
  line: number,
  callContext: CallContext,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined,
  getDocumentState: (uri: string) => DocumentState | undefined
): BuiltinMemberReferenceItem | undefined {
  if (callContext.callPath.length < 2) {
    return undefined;
  }

  const rootResolution = resolveDefinition(uri, {
    character: callContext.callPathStartCharacter,
    line
  });

  if (
    rootResolution &&
    !isBuiltinAliasDocumentModule(rootResolution, stripIndexedAccessMarker(callContext.callPath[0]), getDocumentState)
  ) {
    return undefined;
  }

  const ownerName = resolveBuiltinMemberOwner(callContext.callPath.slice(0, -1));
  const memberName = callContext.callPath[callContext.callPath.length - 1];
  const memberReference = ownerName ? getBuiltinMemberReferenceItem(ownerName, memberName) : undefined;

  if (!memberReference) {
    return undefined;
  }

  return memberReference.signature ||
    (memberReference.completionKind === "function" && memberReference.memberKind === "method")
    ? memberReference
    : undefined;
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

function createBuiltinSignatureHint(
  analysis: AnalysisResult,
  sourceUri: string,
  line: number,
  callContext: CallContext,
  memberReference: BuiltinMemberReferenceItem,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined
): SignatureHint {
  const signature = memberReference.signature;
  const currentArgumentTypeName = getCurrentArgumentTypeName(analysis, sourceUri, line, callContext, resolveDefinition);
  const activeParameter =
    !signature || signature.parameters.length === 0 ? undefined : Math.min(callContext.activeParameter, signature.parameters.length - 1);

  return {
    activeParameter,
    activeSignature: 0,
    documentation: buildBuiltinSignatureDocumentation(memberReference),
    label: signature?.label ?? `${memberReference.ownerName}.${memberReference.name}()`,
    parameters:
      signature?.parameters.map((parameter, index) => ({
        documentation: buildBuiltinSignatureParameterDocumentation(
          parameter,
          index === activeParameter ? currentArgumentTypeName : undefined
        ),
        label: parameter.label
      })) ?? []
  };
}

function getBuiltinMemberHover(
  state: DocumentState,
  uri: string,
  position: LinePosition,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined,
  getDocumentState: (uri: string) => DocumentState | undefined
): HoverHint | undefined {
  const builtinMember = resolveBuiltinMemberAtPosition(state.text, uri, position, resolveDefinition, getDocumentState);

  if (!builtinMember) {
    return undefined;
  }

  return {
    contents: buildBuiltinHoverMarkdown(builtinMember.reference),
    range: builtinMember.range
  };
}

function resolveBuiltinMemberAtPosition(
  text: string,
  uri: string,
  position: LinePosition,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined,
  getDocumentState: (uri: string) => DocumentState | undefined
): { range: SourceRange; reference: BuiltinMemberReferenceItem } | undefined {
  const range = getIdentifierRangeAtPosition(text, position);

  if (!range) {
    return undefined;
  }

  const line = text.replace(/\r\n?/g, "\n").split("\n")[position.line] ?? "";
  const { code } = splitCodeAndComment(line.slice(0, range.end.character));
  const memberAccess = parseTrailingMemberAccess(code);

  if (
    !memberAccess ||
    normalizeIdentifier(memberAccess.prefix) !== normalizeIdentifier(line.slice(range.start.character, range.end.character)) ||
    !canResolveBuiltinAliasMemberAccess(
      uri,
      position.line,
      memberAccess.memberPath[0],
      memberAccess.memberPathStartCharacter,
      resolveDefinition,
      getDocumentState
    )
  ) {
    return undefined;
  }

  const ownerName = resolveBuiltinMemberOwner(memberAccess.memberPath);
  const memberReference = ownerName ? getBuiltinMemberReferenceItem(ownerName, memberAccess.prefix) : undefined;

  return memberReference ? { range, reference: memberReference } : undefined;
}

function buildBuiltinSignatureDocumentation(memberReference: BuiltinMemberReferenceItem): string | undefined {
  const lines = [];

  if (memberReference.summary) {
    lines.push(memberReference.summary);
  }

  if (memberReference.learnUrl) {
    lines.push(memberReference.learnUrl);
  }

  return lines.length > 0 ? lines.join("\n") : undefined;
}

function buildBuiltinSignatureParameterDocumentation(
  parameter: BuiltinCallableSignature["parameters"][number],
  currentArgumentTypeName?: string
): string | undefined {
  const lines = [];

  if (parameter.dataType) {
    lines.push(`想定型: ${parameter.dataType}`);
  }

  if (parameter.isRequired !== undefined) {
    lines.push(parameter.isRequired ? "必須引数" : "省略可能");
  }

  if (parameter.description) {
    lines.push(parameter.description);
  }

  if (currentArgumentTypeName) {
    lines.push(`現在の引数型: ${currentArgumentTypeName}`);
  }

  return lines.length > 0 ? lines.join("\n") : undefined;
}

function buildBuiltinHoverMarkdown(memberReference: BuiltinMemberReferenceItem): string {
  const lines = ["```vb", buildBuiltinMemberLabel(memberReference), "```"];

  if (memberReference.summary) {
    lines.push(memberReference.summary);
  }

  if (memberReference.learnUrl) {
    lines.push(`[Microsoft Learn](${memberReference.learnUrl})`);
  }

  return lines.join("\n\n");
}

function buildBuiltinMemberLabel(memberReference: BuiltinMemberReferenceItem): string {
  if (memberReference.signature?.label) {
    return memberReference.signature.label;
  }

  return memberReference.memberKind === "method"
    ? `${memberReference.ownerName}.${memberReference.name}()`
    : `${memberReference.ownerName}.${memberReference.name}`;
}

function canResolveBuiltinAliasMemberAccess(
  uri: string,
  line: number,
  rootSegment: string,
  rootStartCharacter: number,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined,
  getDocumentState: (uri: string) => DocumentState | undefined
): boolean {
  const rootResolution = resolveDefinition(uri, {
    character: rootStartCharacter,
    line
  });

  return !rootResolution || isBuiltinAliasDocumentModule(rootResolution, rootSegment, getDocumentState);
}

function isBuiltinAliasDocumentModule(
  resolution: WorkspaceSymbolResolution,
  rootSegment: string,
  getDocumentState: (uri: string) => DocumentState | undefined
): boolean {
  if (normalizeIdentifier(rootSegment) !== "thisworkbook" || resolution.symbol.kind !== "module") {
    return false;
  }

  return isWorkbookDocumentState(getDocumentState(resolution.uri));
}

function isWorkbookDocumentState(state: DocumentState | undefined): boolean {
  return (
    !!state &&
    state.analysis.source.moduleKind === "class" &&
    normalizeIdentifier(state.analysis.module.name) === "thisworkbook" &&
    hasModuleAttribute(state, "VB_PredeclaredId", (value) => normalizeIdentifier(value ?? "") === "true") &&
    hasModuleAttribute(state, "VB_Base")
  );
}

function hasModuleAttribute(
  state: DocumentState,
  attributeName: string,
  predicate?: (value: string | undefined) => boolean
): boolean {
  return state.analysis.module.members.some(
    (member) =>
      member.kind === "attributeLine" &&
      normalizeIdentifier(member.name) === normalizeIdentifier(attributeName) &&
      (!predicate || predicate(member.value))
  );
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
    callPath?: string[];
    callPathStartCharacter?: number;
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
      const callable = getCallableBeforeOpenParen(code, index);

      frames.push({
        callPath: callable?.path,
        callPathStartCharacter: callable?.pathStartCharacter,
        commaCount: 0,
        identifier: callable?.identifier,
        identifierStartCharacter: callable?.identifierStartCharacter,
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
    callPath: currentFrame.callPath ?? [currentFrame.identifier],
    callPathStartCharacter: currentFrame.callPathStartCharacter ?? currentFrame.identifierStartCharacter,
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

function getCallableBeforeOpenParen(
  text: string,
  openParenIndex: number
):
  | {
      identifier: string;
      identifierStartCharacter: number;
      path: string[];
      pathStartCharacter: number;
    }
  | undefined {
  const memberAccess = parseTrailingMemberAccess(text.slice(0, openParenIndex));

  if (memberAccess) {
    return {
      identifier: memberAccess.prefix,
      identifierStartCharacter: memberAccess.prefixStartCharacter,
      path: [...memberAccess.memberPath, memberAccess.prefix],
      pathStartCharacter: memberAccess.memberPathStartCharacter
    };
  }

  const identifier = getIdentifierBeforeOpenParen(text, openParenIndex);

  return identifier
    ? {
        identifier: identifier.text,
        identifierStartCharacter: identifier.startCharacter,
        path: [identifier.text],
        pathStartCharacter: identifier.startCharacter
      }
    : undefined;
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

function isSameSymbol(left: SymbolInfo, right: SymbolInfo): boolean {
  return (
    left.scope === right.scope &&
    left.kind === right.kind &&
    left.normalizedName === right.normalizedName &&
    left.selectionRange.start.line === right.selectionRange.start.line &&
    left.selectionRange.start.character === right.selectionRange.start.character &&
    left.selectionRange.end.line === right.selectionRange.end.line &&
    left.selectionRange.end.character === right.selectionRange.end.character
  );
}

function isSameResolution(left: WorkspaceSymbolResolution, right: WorkspaceSymbolResolution): boolean {
  return left.uri === right.uri && isSameSymbol(left.symbol, right.symbol);
}

function comparePositions(left: LinePosition, right: LinePosition): number {
  if (left.line !== right.line) {
    return left.line - right.line;
  }

  return left.character - right.character;
}

function isWorkspaceVisible(modifier?: string): boolean {
  return /^(public|friend)$/i.test(modifier ?? "");
}

function mapSemanticToken(
  symbol: SymbolInfo,
  explicitType?: BuiltinSemanticType,
  explicitModifiers?: BuiltinSemanticModifier[]
): SemanticTokenShape | undefined {
  if (explicitType) {
    return {
      modifiers: [...(explicitModifiers ?? [])],
      type: explicitType as SemanticTokenTypeName
    };
  }

  switch (symbol.kind) {
    case "constant":
      return {
        modifiers: ["readonly"],
        type: "variable"
      };
    case "declare":
    case "procedure":
      return {
        modifiers: [],
        type: "function"
      };
    case "enum":
    case "type":
      return {
        modifiers: [],
        type: "type"
      };
    case "enumMember":
      return {
        modifiers: ["readonly"],
        type: "enumMember"
      };
    case "parameter":
      return {
        modifiers: [],
        type: "parameter"
      };
    case "variable":
      return {
        modifiers: [],
        type: "variable"
      };
    default:
      return undefined;
  }
}

function mapBuiltinSemanticToken(identifier: string): SemanticTokenShape | undefined {
  const referenceItem = getBuiltinReferenceItem(identifier);

  return referenceItem
    ? {
        modifiers: [...referenceItem.modifiers],
        type: referenceItem.semanticType as SemanticTokenTypeName
      }
    : undefined;
}

function mapBuiltinMemberSemanticToken(
  uri: string,
  line: number,
  lineText: string,
  startCharacter: number,
  identifier: string,
  resolveDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined,
  getDocumentState: (uri: string) => DocumentState | undefined
): SemanticTokenShape | undefined {
  const memberAccess = parseTrailingMemberAccess(lineText.slice(0, startCharacter + identifier.length));

  if (
    !memberAccess ||
    memberAccess.memberPathStartCharacter === undefined ||
    normalizeIdentifier(memberAccess.prefix) !== normalizeIdentifier(identifier) ||
    !canResolveBuiltinAliasMemberAccess(
      uri,
      line,
      stripIndexedAccessMarker(memberAccess.memberPath[0]),
      memberAccess.memberPathStartCharacter,
      resolveDefinition,
      getDocumentState
    )
  ) {
    return undefined;
  }

  const ownerName = resolveBuiltinMemberOwner(memberAccess.memberPath);
  const memberReference = ownerName ? getBuiltinMemberReferenceItem(ownerName, identifier) : undefined;

  return memberReference
    ? {
        modifiers: [...memberReference.modifiers],
        type: memberReference.semanticType as SemanticTokenTypeName
      }
    : undefined;
}

function mapBuiltinSymbolKind(completionKind: BuiltinCompletionKind): SymbolInfo["kind"] {
  switch (completionKind) {
    case "constant":
      return "constant";
    case "function":
      return "procedure";
    case "type":
      return "type";
    case "variable":
      return "variable";
    case "keyword":
    default:
      return "variable";
  }
}

function compareSemanticTokens(left: SemanticTokenEntry, right: SemanticTokenEntry): number {
  return (
    comparePositions(left.range.start, right.range.start) ||
    comparePositions(left.range.end, right.range.end) ||
    left.type.localeCompare(right.type) ||
    left.modifiers.join(".").localeCompare(right.modifiers.join("."))
  );
}

function parseTrailingMemberAccess(
  text: string
): {
  memberPath: string[];
  memberPathStartCharacter: number;
  prefix: string;
  prefixStartCharacter: number;
} | undefined {
  let index = text.length - 1;

  while (index >= 0 && /\s/u.test(text[index] ?? "")) {
    index -= 1;
  }

  const prefixEnd = index + 1;

  while (index >= 0 && /[A-Za-z0-9_$%&!#@]/u.test(text[index] ?? "")) {
    index -= 1;
  }

  const prefix = text.slice(index + 1, prefixEnd);
  const prefixStartCharacter = index + 1;

  if (prefix.length > 0 && !/^[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?$/u.test(prefix)) {
    return undefined;
  }

  while (index >= 0 && /\s/u.test(text[index] ?? "")) {
    index -= 1;
  }

  if (text[index] !== ".") {
    return undefined;
  }

  const memberPath: string[] = [];
  let memberPathStartCharacter: number | undefined;

  while (index >= 0 && text[index] === ".") {
    index -= 1;

    while (index >= 0 && /\s/u.test(text[index] ?? "")) {
      index -= 1;
    }

    const indexedAccess = skipTrailingIndexedAccess(text, index);
    index = indexedAccess.index;

    const identifierEnd = index + 1;

    while (index >= 0 && /[A-Za-z0-9_$%&!#@]/u.test(text[index] ?? "")) {
      index -= 1;
    }

    const identifier = text.slice(index + 1, identifierEnd);

    if (!/^[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?$/u.test(identifier)) {
      return undefined;
    }

    memberPath.unshift(indexedAccess.isIndexed ? `${identifier}()` : identifier);
    memberPathStartCharacter = index + 1;

    while (index >= 0 && /\s/u.test(text[index] ?? "")) {
      index -= 1;
    }
  }

  return memberPath.length > 0
    ? {
        memberPath,
        memberPathStartCharacter: memberPathStartCharacter ?? 0,
        prefix,
        prefixStartCharacter
      }
    : undefined;
}

function getTrailingIdentifier(text: string): string {
  return /[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?$/u.exec(text)?.[0] ?? "";
}

function skipTrailingIndexedAccess(
  text: string,
  startIndex: number
): {
  index: number;
  isIndexed: boolean;
} {
  let index = startIndex;

  while (index >= 0 && /\s/u.test(text[index] ?? "")) {
    index -= 1;
  }

  if (text[index] !== ")") {
    return {
      index,
      isIndexed: false
    };
  }

  let depth = 0;

  while (index >= 0) {
    const character = text[index];

    if (character === "\"") {
      index = skipStringLiteralBackward(text, index);
      continue;
    }

    if (character === ")") {
      depth += 1;
      index -= 1;
      continue;
    }

    if (character === "(") {
      depth -= 1;
      index -= 1;

      if (depth === 0) {
        break;
      }

      continue;
    }

    index -= 1;
  }

  while (index >= 0 && /\s/u.test(text[index] ?? "")) {
    index -= 1;
  }

  return {
    index,
    isIndexed: true
  };
}

function skipStringLiteralBackward(text: string, endQuoteIndex: number): number {
  let index = endQuoteIndex - 1;

  while (index >= 0) {
    if (text[index] !== "\"") {
      index -= 1;
      continue;
    }

    if (index - 1 >= 0 && text[index - 1] === "\"") {
      index -= 2;
      continue;
    }

    return index - 1;
  }

  return -1;
}

function stripIndexedAccessMarker(pathSegment: string): string {
  return pathSegment.endsWith("()") ? pathSegment.slice(0, -2) : pathSegment;
}

function rangesEqual(left: SourceRange, right: SourceRange): boolean {
  return (
    left.start.line === right.start.line &&
    left.start.character === right.start.character &&
    left.end.line === right.end.line &&
    left.end.character === right.end.character
  );
}
