import {
  CodeAction,
  CodeActionKind,
  CompletionItem,
  CompletionItemKind,
  createConnection,
  Definition,
  DiagnosticSeverity,
  DocumentSymbol,
  FileChangeType,
  Hover,
  InitializeParams,
  InitializeResult,
  Location,
  MarkupKind,
  ParameterInformation,
  Position,
  ProposedFeatures,
  Range,
  SemanticTokensBuilder,
  SignatureHelp,
  SignatureInformation,
  SymbolKind,
  TextEdit,
  TextDocumentSyncKind
} from "vscode-languageserver/node";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { TextDocuments } from "vscode-languageserver";
import { TextDocument } from "vscode-languageserver-textdocument";
import {
  createDocumentService,
  SEMANTIC_TOKEN_MODIFIERS,
  SEMANTIC_TOKEN_TYPES
} from "./lsp/documentService";
import type {
  DocumentCodeAction,
  DocumentState,
  SemanticTokenEntry,
  SignatureHint,
  WorkspaceSymbolResolution
} from "./lsp/documentService";
import type { Diagnostic, OutlineSymbol, SymbolInfo } from "../../core/src/index";
import { ACTIVE_WORKBOOK_IDENTITY_NOTIFICATION_METHOD } from "../../core/src/index";

export { createDocumentService } from "./lsp/documentService";

interface ServerSettings {
  analysisDebounceMs: number;
}

export function startServer(): void {
  const connection = createConnection(ProposedFeatures.all);
  const documents = new TextDocuments(TextDocument);
  const documentService = createDocumentService({
    logger: (entry) => {
      const channel = entry.code.startsWith("active-workbook-identity.")
        ? "active-workbook-identity"
        : entry.code.startsWith("worksheet-control-metadata.")
          ? "worksheet-control-metadata"
          : "vba";
      const message = `[${channel}] ${entry.message}`;

      if (entry.level === "warn") {
        connection.console.warn(message);
        return;
      }

      connection.console.info(message);
    }
  });
  const pendingTimers = new Map<string, NodeJS.Timeout>();
  let settings: ServerSettings = {
    analysisDebounceMs: 300
  };
  let canReadConfiguration = false;

  connection.onInitialize((params: InitializeParams): InitializeResult => {
    canReadConfiguration = Boolean(params.capabilities.workspace?.configuration);
    documentService.setWorkspaceRoots([
      ...(params.workspaceFolders?.map((workspaceFolder) => workspaceFolder.uri) ?? []),
      ...(params.rootUri ? [params.rootUri] : []),
      ...(params.rootPath ? [params.rootPath] : [])
    ]);

    return {
      capabilities: {
        codeActionProvider: true,
        completionProvider: {},
        definitionProvider: true,
        documentFormattingProvider: true,
        documentSymbolProvider: true,
        hoverProvider: true,
        renameProvider: {
          prepareProvider: true
        },
        referencesProvider: true,
        semanticTokensProvider: {
          full: true,
          legend: {
            tokenModifiers: [...SEMANTIC_TOKEN_MODIFIERS],
            tokenTypes: [...SEMANTIC_TOKEN_TYPES]
          }
        },
        signatureHelpProvider: {
          retriggerCharacters: [","],
          triggerCharacters: ["(", ","]
        },
        textDocumentSync: TextDocumentSyncKind.Full
      }
    };
  });

  connection.onInitialized(async () => {
    if (canReadConfiguration) {
      settings = await readSettings(connection);
    }

    await primeWorkspaceIndex();

    for (const document of documents.all()) {
      analyzeAndPublish(document);
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

    connection.sendDiagnostics({ diagnostics: [], uri: event.document.uri });
    void restoreWorkspaceDocument(event.document.uri);
  });

  connection.onDidChangeConfiguration(async () => {
    if (canReadConfiguration) {
      settings = await readSettings(connection);
    }

    for (const document of documents.all()) {
      analyzeAndPublish(document);
    }
  });

  connection.onNotification(ACTIVE_WORKBOOK_IDENTITY_NOTIFICATION_METHOD, (snapshot) => {
    documentService.setActiveWorkbookIdentitySnapshot(snapshot);
  });

  connection.onDidChangeWatchedFiles(async (params) => {
    const openDocumentUris = new Set(documents.all().map((document) => document.uri));

    for (const change of params.changes) {
      if (openDocumentUris.has(change.uri)) {
        continue;
      }

      if (change.type === FileChangeType.Deleted) {
        documentService.remove(change.uri);
      } else {
        await restoreWorkspaceDocument(change.uri);
      }
    }

    for (const document of documents.all()) {
      analyzeAndPublish(document);
    }
  });

  connection.onCompletion((params): CompletionItem[] => {
    return documentService.getCompletionSymbols(params.textDocument.uri, toCorePosition(params.position)).map(toCompletionItem);
  });

  connection.onCodeAction((params): CodeAction[] => {
    if (
      params.context.only &&
      !params.context.only.some(
        (requestedKind) =>
          requestedKind === CodeActionKind.QuickFix || requestedKind.startsWith(`${CodeActionKind.QuickFix}.`)
      )
    ) {
      return [];
    }

    return documentService.getCodeActions(params.textDocument.uri).map(toCodeAction);
  });

  connection.onDefinition((params): Definition | undefined => {
    const resolution = documentService.getDefinition(params.textDocument.uri, toCorePosition(params.position));

    if (!resolution) {
      return undefined;
    }

    return Location.create(resolution.uri, toLspRange(resolution.symbol.selectionRange));
  });

  connection.onHover((params): Hover | undefined => {
    const hover = documentService.getHover(params.textDocument.uri, toCorePosition(params.position));

    return hover
      ? {
          contents: {
            kind: MarkupKind.Markdown,
            value: hover.contents
          },
          range: hover.range ? toLspRange(hover.range) : undefined
        }
      : undefined;
  });

  connection.onDocumentSymbol((params): DocumentSymbol[] => {
    return documentService.getDocumentSymbols(params.textDocument.uri).map(toDocumentSymbol);
  });

  connection.onDocumentFormatting((params): TextEdit[] => {
    const state = documentService.getState(params.textDocument.uri);
    const formattedText = documentService.formatDocument(params.textDocument.uri, params.options);

    if (!state || formattedText === undefined || formattedText === state.text) {
      return [];
    }

    return [TextEdit.replace(toLspRange(getFullDocumentRange(state.text)), formattedText)];
  });

  connection.onReferences((params): Location[] => {
    return documentService
      .getReferences(params.textDocument.uri, toCorePosition(params.position), params.context.includeDeclaration)
      .map((reference) => Location.create(reference.uri, toLspRange(reference.range)));
  });

  connection.languages.semanticTokens.on((params) => {
    const builder = new SemanticTokensBuilder();

    for (const token of documentService.getSemanticTokens(params.textDocument.uri)) {
      builder.push(
        token.range.start.line,
        token.range.start.character,
        token.range.end.character - token.range.start.character,
        getSemanticTokenTypeIndex(token),
        getSemanticTokenModifierMask(token)
      );
    }

    return builder.build();
  });

  connection.onPrepareRename((params) => {
    const target = documentService.prepareRename(params.textDocument.uri, toCorePosition(params.position));

    return target
      ? {
          placeholder: target.placeholder,
          range: toLspRange(target.range)
        }
      : null;
  });

  connection.onRenameRequest((params) => {
    const edits = documentService.getRenameEdits(params.textDocument.uri, toCorePosition(params.position), params.newName);

    if (!edits || edits.length === 0) {
      return null;
    }

    const changes = edits.reduce<Record<string, TextEdit[]>>((accumulator, edit) => {
      const currentEdits = accumulator[edit.uri] ?? [];
      currentEdits.push(TextEdit.replace(toLspRange(edit.range), edit.newText));
      accumulator[edit.uri] = currentEdits;
      return accumulator;
    }, {});

    return { changes };
  });

  connection.onSignatureHelp((params): SignatureHelp | undefined => {
    const signatureHelp = documentService.getSignatureHelp(params.textDocument.uri, toCorePosition(params.position));
    return signatureHelp ? toSignatureHelp(signatureHelp) : undefined;
  });

  documents.listen(connection);
  connection.listen();

  function analyzeAndPublish(document: TextDocument): DocumentState {
    const state = documentService.analyzeText(document.uri, document.languageId, document.version, document.getText());
    connection.sendDiagnostics({
      diagnostics: documentService.getDiagnostics(document.uri).map(toLspDiagnostic),
      uri: document.uri
    });
    return state;
  }

  async function primeWorkspaceIndex(): Promise<void> {
    try {
      const openDocumentUris = new Set(documents.all().map((document) => document.uri));
      const workspaceUris = (
        await Promise.all([
          connection.workspace.findFiles("**/*.bas"),
          connection.workspace.findFiles("**/*.cls"),
          connection.workspace.findFiles("**/*.frm")
        ])
      ).flat();

      for (const uri of [...new Set(workspaceUris)]) {
        if (openDocumentUris.has(uri)) {
          continue;
        }

        await restoreWorkspaceDocument(uri);
      }
    } catch (error) {
      connection.console.error(`Failed to index VBA workspace files: ${String(error)}`);
    }
  }

  async function restoreWorkspaceDocument(uri: string): Promise<void> {
    if (!uri.startsWith("file:")) {
      documentService.remove(uri);
      return;
    }

    try {
      const filePath = fileURLToPath(uri);
      const text = await readFile(filePath, "utf8");
      documentService.analyzeText(uri, "vba", 0, text);
    } catch {
      documentService.remove(uri);
    }
  }
}

async function readSettings(connection: ReturnType<typeof createConnection>): Promise<ServerSettings> {
  const configuration = await connection.workspace.getConfiguration("vba");
  const rawDebounce = configuration?.analysis?.debounceMs;

  return {
    analysisDebounceMs: typeof rawDebounce === "number" ? Math.max(50, Math.min(2000, rawDebounce)) : 300
  };
}

function toCompletionItem(resolution: WorkspaceSymbolResolution): CompletionItem {
  return {
    detail: resolution.typeName ? `${resolution.moduleName} : ${resolution.typeName}` : resolution.moduleName,
    documentation: resolution.documentation,
    kind: mapCompletionItemKind(resolution.symbol.kind, resolution.completionItemKind),
    label: resolution.symbol.name,
    sortText: resolution.isBuiltIn ? `~${resolution.symbol.name}` : undefined
  };
}

function toCodeAction(action: DocumentCodeAction): CodeAction {
  return {
    edit: {
      changes: {
        [action.edit.uri]: [TextEdit.replace(toLspRange(action.edit.range), action.edit.newText)]
      }
    },
    isPreferred: true,
    kind: CodeActionKind.QuickFix,
    title: action.title
  };
}

function toSignatureHelp(signature: SignatureHint): SignatureHelp {
  return {
    activeParameter: signature.activeParameter,
    activeSignature: signature.activeSignature,
    signatures: [
      SignatureInformation.create(
        signature.label,
        signature.documentation,
        ...signature.parameters.map((parameter) => ParameterInformation.create(parameter.label, parameter.documentation))
      )
    ]
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

function getFullDocumentRange(text: string): { end: { character: number; line: number }; start: { character: number; line: number } } {
  const normalizedText = text.replace(/\r\n?/g, "\n");
  const lines = normalizedText.split("\n");
  const endLine = Math.max(0, lines.length - 1);

  return {
    start: {
      character: 0,
      line: 0
    },
    end: {
      character: lines[endLine]?.length ?? 0,
      line: endLine
    }
  };
}

function mapCompletionItemKind(
  kind: SymbolInfo["kind"],
  explicitKind?: WorkspaceSymbolResolution["completionItemKind"]
): CompletionItemKind {
  switch (explicitKind) {
    case "constant":
      return CompletionItemKind.Constant;
    case "function":
      return CompletionItemKind.Function;
    case "keyword":
      return CompletionItemKind.Keyword;
    case "type":
      return CompletionItemKind.Class;
    case "variable":
      return CompletionItemKind.Variable;
    default:
      break;
  }

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

function getSemanticTokenModifierMask(token: SemanticTokenEntry): number {
  return token.modifiers.reduce((mask, modifier) => {
    const modifierIndex = SEMANTIC_TOKEN_MODIFIERS.indexOf(modifier);
    return modifierIndex >= 0 ? mask | (1 << modifierIndex) : mask;
  }, 0);
}

function getSemanticTokenTypeIndex(token: SemanticTokenEntry): number {
  return SEMANTIC_TOKEN_TYPES.indexOf(token.type);
}

if (typeof require !== "undefined" && require.main === module) {
  startServer();
}
