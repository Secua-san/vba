import {
  analyzeModule,
  extractIdentifierAtPosition,
  findDefinition,
  getCompletionSymbols,
  getDocumentOutline,
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

export interface DocumentService {
  analyzeText: (uri: string, languageId: string, version: number, text: string) => DocumentState;
  getCompletionSymbols: (uri: string, position: LinePosition) => WorkspaceSymbolResolution[];
  getDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined;
  getDiagnostics: (uri: string) => Diagnostic[];
  getDocumentSymbols: (uri: string) => ReturnType<typeof getDocumentOutline>;
  getReferences: (uri: string, position: LinePosition, includeDeclaration: boolean) => WorkspaceReference[];
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

      return [...deduplicated.values()];
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
