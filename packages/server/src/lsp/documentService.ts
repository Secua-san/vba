import {
  analyzeModule,
  extractIdentifierAtPosition,
  findDefinition,
  getCompletionSymbols,
  getDocumentOutline,
  normalizeIdentifier,
  type AnalysisResult,
  type Diagnostic,
  type LinePosition,
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
  uri: string;
}

export interface DocumentService {
  analyzeText: (uri: string, languageId: string, version: number, text: string) => DocumentState;
  getCompletionSymbols: (uri: string, position: LinePosition) => WorkspaceSymbolResolution[];
  getDefinition: (uri: string, position: LinePosition) => WorkspaceSymbolResolution | undefined;
  getDiagnostics: (uri: string) => Diagnostic[];
  getDocumentSymbols: (uri: string) => ReturnType<typeof getDocumentOutline>;
  getState: (uri: string) => DocumentState | undefined;
  remove: (uri: string) => void;
}

export function createDocumentService(): DocumentService {
  const documentStates = new Map<string, DocumentState>();
  let workspaceIndex = createWorkspaceIndex([]);

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
        moduleName: state.analysis.module.name,
        symbol,
        uri
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
      const state = documentStates.get(uri);

      if (!state) {
        return undefined;
      }

      const localDefinition = findDefinition(state.analysis, position);

      if (localDefinition) {
        return {
          moduleName: state.analysis.module.name,
          symbol: localDefinition,
          uri
        };
      }

      const identifier = extractIdentifierAtPosition(state.text.replace(/\r\n?/g, "\n"), position);

      if (!identifier) {
        return undefined;
      }

      const matches = workspaceIndex.byNormalizedName
        .get(normalizeIdentifier(identifier))
        ?.filter((resolution) => resolution.uri !== uri) ?? [];

      return matches.length === 1 ? matches[0] : undefined;
    },
    getDiagnostics(uri: string): Diagnostic[] {
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
    },
    getDocumentSymbols(uri: string) {
      const state = documentStates.get(uri);
      return state ? getDocumentOutline(state.analysis) : [];
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
    {
      moduleName: state.analysis.module.name,
      symbol: state.analysis.symbols.moduleSymbol,
      uri: state.uri
    }
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

function findModuleSymbols(
  moduleSymbolsByName: Map<string, SymbolInfo[]>,
  name: string,
  state: DocumentState
): WorkspaceSymbolResolution[] {
  return (moduleSymbolsByName.get(normalizeIdentifier(name)) ?? []).map((symbol) => ({
    moduleName: state.analysis.module.name,
    symbol,
    uri: state.uri
  }));
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

function isWorkspaceVisible(modifier?: string): boolean {
  return /^(public|friend)$/i.test(modifier ?? "");
}
