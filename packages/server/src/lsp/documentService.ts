import {
  analyzeModule,
  findDefinition,
  getCompletionSymbols,
  getDocumentOutline,
  type AnalysisResult,
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

export interface DocumentService {
  analyzeText: (uri: string, languageId: string, version: number, text: string) => DocumentState;
  getCompletionSymbols: (uri: string, position: LinePosition) => SymbolInfo[];
  getDefinition: (uri: string, position: LinePosition) => SymbolInfo | undefined;
  getDocumentSymbols: (uri: string) => ReturnType<typeof getDocumentOutline>;
  getState: (uri: string) => DocumentState | undefined;
  remove: (uri: string) => void;
}

export function createDocumentService(): DocumentService {
  const documentStates = new Map<string, DocumentState>();

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
      return state;
    },
    getCompletionSymbols(uri: string, position: LinePosition): SymbolInfo[] {
      const state = documentStates.get(uri);
      return state ? getCompletionSymbols(state.analysis, position) : [];
    },
    getDefinition(uri: string, position: LinePosition): SymbolInfo | undefined {
      const state = documentStates.get(uri);
      return state ? findDefinition(state.analysis, position) : undefined;
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
    }
  };
}

function getFileNameFromUri(uri: string): string | undefined {
  const normalizedUri = uri.startsWith("file:///") ? decodeURIComponent(uri.replace("file:///", "")) : uri;
  const segments = normalizedUri.split(/[\\/]/);
  return segments[segments.length - 1];
}
