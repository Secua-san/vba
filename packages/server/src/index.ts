import {
  CompletionItem,
  CompletionItemKind,
  createConnection,
  Definition,
  DiagnosticSeverity,
  DocumentSymbol,
  InitializeParams,
  InitializeResult,
  Location,
  Position,
  ProposedFeatures,
  Range,
  SymbolKind,
  TextDocumentSyncKind
} from "vscode-languageserver/node";
import { TextDocuments } from "vscode-languageserver";
import { TextDocument } from "vscode-languageserver-textdocument";
import { createDocumentService } from "./lsp/documentService";
import type { DocumentState } from "./lsp/documentService";
import type { Diagnostic, OutlineSymbol, SymbolInfo } from "../../core/src/index";

export { createDocumentService } from "./lsp/documentService";

interface ServerSettings {
  analysisDebounceMs: number;
}

export function startServer(): void {
  const connection = createConnection(ProposedFeatures.all);
  const documents = new TextDocuments(TextDocument);
  const documentService = createDocumentService();
  const pendingTimers = new Map<string, NodeJS.Timeout>();
  let settings: ServerSettings = {
    analysisDebounceMs: 300
  };
  let canReadConfiguration = false;

  connection.onInitialize((params: InitializeParams): InitializeResult => {
    canReadConfiguration = Boolean(params.capabilities.workspace?.configuration);

    return {
      capabilities: {
        completionProvider: {},
        definitionProvider: true,
        documentSymbolProvider: true,
        textDocumentSync: TextDocumentSyncKind.Full
      }
    };
  });

  connection.onInitialized(async () => {
    if (canReadConfiguration) {
      settings = await readSettings(connection);
    }
  });

  documents.onDidOpen((event) => {
    analyzeAndPublish(event.document);
  });

  documents.onDidChangeContent((event) => {
    const pendingTimer = pendingTimers.get(event.document.uri);

    if (pendingTimer) {
      clearTimeout(pendingTimer);
    }

    pendingTimers.set(
      event.document.uri,
      setTimeout(() => {
        analyzeAndPublish(event.document);
        pendingTimers.delete(event.document.uri);
      }, settings.analysisDebounceMs)
    );
  });

  documents.onDidClose((event) => {
    const pendingTimer = pendingTimers.get(event.document.uri);

    if (pendingTimer) {
      clearTimeout(pendingTimer);
      pendingTimers.delete(event.document.uri);
    }

    documentService.remove(event.document.uri);
    connection.sendDiagnostics({ diagnostics: [], uri: event.document.uri });
  });

  connection.onDidChangeConfiguration(async () => {
    if (canReadConfiguration) {
      settings = await readSettings(connection);
    }

    for (const document of documents.all()) {
      analyzeAndPublish(document);
    }
  });

  connection.onCompletion((params): CompletionItem[] => {
    return documentService.getCompletionSymbols(params.textDocument.uri, toCorePosition(params.position)).map(toCompletionItem);
  });

  connection.onDefinition((params): Definition | undefined => {
    const symbol = documentService.getDefinition(params.textDocument.uri, toCorePosition(params.position));

    if (!symbol) {
      return undefined;
    }

    return Location.create(params.textDocument.uri, toLspRange(symbol.selectionRange));
  });

  connection.onDocumentSymbol((params): DocumentSymbol[] => {
    return documentService.getDocumentSymbols(params.textDocument.uri).map(toDocumentSymbol);
  });

  documents.listen(connection);
  connection.listen();

  function analyzeAndPublish(document: TextDocument): DocumentState {
    const state = documentService.analyzeText(document.uri, document.languageId, document.version, document.getText());
    connection.sendDiagnostics({
      diagnostics: state.analysis.diagnostics.map(toLspDiagnostic),
      uri: document.uri
    });
    return state;
  }
}

async function readSettings(connection: ReturnType<typeof createConnection>): Promise<ServerSettings> {
  const configuration = await connection.workspace.getConfiguration("vba");
  const rawDebounce = configuration?.analysis?.debounceMs;

  return {
    analysisDebounceMs: typeof rawDebounce === "number" ? Math.max(50, Math.min(2000, rawDebounce)) : 300
  };
}

function toCompletionItem(symbol: SymbolInfo): CompletionItem {
  return {
    kind: mapCompletionItemKind(symbol.kind),
    label: symbol.name
  };
}

function toCorePosition(position: Position): { character: number; line: number } {
  return {
    character: position.character,
    line: position.line
  };
}

function toDocumentSymbol(symbol: OutlineSymbol): DocumentSymbol {
  return DocumentSymbol.create(
    symbol.name,
    "",
    mapDocumentSymbolKind(symbol.kind),
    toLspRange(symbol.range),
    toLspRange(symbol.selectionRange),
    symbol.children?.map(toDocumentSymbol) ?? []
  );
}

function toLspDiagnostic(diagnostic: Diagnostic) {
  return {
    message: diagnostic.message,
    range: toLspRange(diagnostic.range),
    severity: diagnostic.severity === "warning" ? DiagnosticSeverity.Warning : DiagnosticSeverity.Error,
    source: "vba"
  };
}

function toLspRange(range: { end: Position; start: Position } | { end: { character: number; line: number }; start: { character: number; line: number } }): Range {
  return Range.create(range.start.line, range.start.character, range.end.line, range.end.character);
}

function mapCompletionItemKind(kind: SymbolInfo["kind"]): CompletionItemKind {
  switch (kind) {
    case "constant":
    case "enumMember":
      return CompletionItemKind.Constant;
    case "declare":
    case "procedure":
      return CompletionItemKind.Function;
    case "enum":
    case "type":
      return CompletionItemKind.Enum;
    case "parameter":
      return CompletionItemKind.Field;
    case "variable":
      return CompletionItemKind.Variable;
    default:
      return CompletionItemKind.Text;
  }
}

function mapDocumentSymbolKind(kind: OutlineSymbol["kind"]): SymbolKind {
  switch (kind) {
    case "constant":
    case "enumMember":
      return SymbolKind.Constant;
    case "declare":
    case "procedure":
      return SymbolKind.Function;
    case "enum":
      return SymbolKind.Enum;
    case "module":
      return SymbolKind.Module;
    case "type":
      return SymbolKind.Struct;
    case "variable":
      return SymbolKind.Variable;
    default:
      return SymbolKind.Object;
  }
}

if (typeof require !== "undefined" && require.main === module) {
  startServer();
}
