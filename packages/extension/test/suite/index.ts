import assert from "node:assert/strict";
import path from "node:path";
import * as vscode from "vscode";

export async function run(): Promise<void> {
  const extension = vscode.extensions.getExtension("tagi0.vba-extension");
  assert.ok(extension, "extension must be discoverable");

  await extension.activate();
  await vscode.workspace.getConfiguration("editor").update("snippetSuggestions", "top", vscode.ConfigurationTarget.Global);
  await vscode.workspace.getConfiguration("editor").update("insertSpaces", true, vscode.ConfigurationTarget.Global);
  await vscode.workspace.getConfiguration("editor").update("tabSize", 4, vscode.ConfigurationTarget.Global);

  const fixturesPath = path.resolve(__dirname, "..", "..", "test", "fixtures");
  const sampleDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "sample.bas"));
  await vscode.window.showTextDocument(sampleDocument);

  assert.equal(sampleDocument.languageId, "vba");

  const symbols = await waitForSymbols(sampleDocument);
  assert.ok(symbols.length > 0, "document symbols should be available");

  const libraryDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "PublicApi.bas"));
  await vscode.window.showTextDocument(libraryDocument);
  const numberDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "NumberApi.bas"));
  await vscode.window.showTextDocument(numberDocument);
  const formatterDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "FormatterApi.bas"));
  await vscode.window.showTextDocument(formatterDocument);

  const consumerDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "Consumer.bas"));
  await vscode.window.showTextDocument(consumerDocument);

  const completionItems = await waitForCompletions(
    consumerDocument,
    new vscode.Position(5, 4),
    (items) => items.some((item) => item.label === "PublicMessage")
  );
  const publicMessageCompletion = completionItems.find((item) => item.label === "PublicMessage");

  assert.ok(
    completionItems.some((item) => item.label === "PublicMessage"),
    "cross-file completion should include exported workspace symbols"
  );
  assert.ok(publicMessageCompletion?.detail?.includes("String"), "completion detail should include inferred type information");

  const definitions = await waitForDefinitions(
    consumerDocument,
    new vscode.Position(5, 18),
    (locations) => locations.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "PublicApi.bas")))
  );
  assert.ok(definitions.length > 0, "definition should resolve across files");
  assert.ok(definitions.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "PublicApi.bas"))));

  const references = await waitForReferences(
    consumerDocument,
    new vscode.Position(5, 18),
    (locations) =>
      locations.length >= 2 &&
      locations.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "Consumer.bas"))) &&
      locations.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "PublicApi.bas")))
  );
  assert.ok(references.length >= 2, "references should include declaration and usage");
  assert.ok(references.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "Consumer.bas"))));
  assert.ok(references.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "PublicApi.bas"))));

  const consumerCompletionDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "ConsumerCompletion.bas"));
  await vscode.window.showTextDocument(consumerCompletionDocument);

  const narrowedCompletionItems = await waitForCompletions(
    consumerCompletionDocument,
    new vscode.Position(5, 17),
    (items) => items.some((item) => item.label === "PublicMessage")
  );
  assert.ok(
    narrowedCompletionItems.some((item) => item.label === "PublicMessage"),
    "type-aware completion should keep compatible candidates"
  );
  assert.equal(
    narrowedCompletionItems.some((item) => item.label === "PublicNumber"),
    false,
    "type-aware completion should hide incompatible candidates"
  );

  const consumerSignatureDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "ConsumerSignature.bas"));
  await vscode.window.showTextDocument(consumerSignatureDocument);

  const signatureHelp = await waitForSignatureHelp(
    consumerSignatureDocument,
    new vscode.Position(5, 38),
    (help) => help.signatures.length > 0 && help.activeParameter === 1
  );
  assert.ok(signatureHelp.signatures.length > 0, "signature help should be available across files");
  assert.equal(signatureHelp.activeParameter, 1);
  assert.equal(signatureHelp.signatures[0]?.label, "FormatMessage(ByVal value As String, ByVal count As Long) As String");
  assert.equal(getSignatureDocumentation(signatureHelp.signatures[0]?.documentation), "FormatterApi モジュール");
  assert.ok(
    getSignatureDocumentation(signatureHelp.signatures[0]?.parameters[1]?.documentation).includes("現在の引数型: Long"),
    "signature help should include inferred argument type information"
  );

  const renameDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "RenameLocal.bas"));
  await vscode.window.showTextDocument(renameDocument);

  const renameEdit = await waitForRename(
    renameDocument,
    new vscode.Position(6, 6),
    "currentCount",
    (edit) => (edit.get(renameDocument.uri)?.length ?? 0) === 4
  );
  const renameEntries = renameEdit.get(renameDocument.uri) ?? [];

  assert.equal(renameEntries.length, 4, "rename should update declaration and local references only");
  assert.ok(renameEntries.every((edit) => edit.newText === "currentCount"));
  assert.ok(renameEntries.every((edit) => edit.range.start.line >= 4 && edit.range.start.line <= 8));

  const semanticDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "SemanticTokens.bas"));
  await vscode.window.showTextDocument(semanticDocument);

  const semanticLegend = await waitForSemanticTokensLegend(
    semanticDocument,
    (legend) =>
      legend.tokenTypes.includes("variable") &&
      legend.tokenTypes.includes("parameter") &&
      legend.tokenTypes.includes("function") &&
      legend.tokenTypes.includes("type")
  );
  const semanticTokens = await waitForSemanticTokens(
    semanticDocument,
    (tokens) => tokens.data.length > 0
  );

  assert.ok(semanticLegend.tokenTypes.includes("variable"), "semantic token legend should include variable");
  assert.ok(semanticLegend.tokenTypes.includes("parameter"), "semantic token legend should include parameter");
  assert.ok(semanticLegend.tokenTypes.includes("function"), "semantic token legend should include function");
  assert.ok(semanticLegend.tokenTypes.includes("type"), "semantic token legend should include type");
  assert.ok(semanticTokens.data.length > 0, "semantic tokens should be available");

  const formatDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "FormatDocument.bas"));
  await vscode.window.showTextDocument(formatDocument);

  const formattedText = await waitForFormattedDocument(
    formatDocument,
    `Attribute VB_Name = "FormatDocument"
Option Explicit

Public Sub Demo()
    If True Then
        Debug.Print "ready"
    Else
        Select Case 1
            Case Else
                Debug.Print "fallback"
        End Select
    End If
End Sub`
  );

  assert.equal(normalizeText(formattedText), normalizeText(`Attribute VB_Name = "FormatDocument"
Option Explicit

Public Sub Demo()
    If True Then
        Debug.Print "ready"
    Else
        Select Case 1
            Case Else
                Debug.Print "fallback"
        End Select
    End If
End Sub`));

  const continuationDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "ContinuationFormatting.bas"));
  await vscode.window.showTextDocument(continuationDocument);

  const formattedContinuationText = await waitForFormattedDocument(
    continuationDocument,
    `Attribute VB_Name = "ContinuationFormatting"
Option Explicit

Public Sub Demo()
    Dim message As String
    message = _
        "prefix" & _
        "suffix"

    Debug.Print JoinValues( _
        message, _
        "tail" _
    )

    message = CreateBuilder() _
        .WithName(message) _
        .WithSuffix("!")
End Sub`
  );

  assert.equal(normalizeText(formattedContinuationText), normalizeText(`Attribute VB_Name = "ContinuationFormatting"
Option Explicit

Public Sub Demo()
    Dim message As String
    message = _
        "prefix" & _
        "suffix"

    Debug.Print JoinValues( _
        message, _
        "tail" _
    )

    message = CreateBuilder() _
        .WithName(message) _
        .WithSuffix("!")
End Sub`));

  const snippetDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "SnippetCompletions.bas"));
  await vscode.window.showTextDocument(snippetDocument);

  const subSnippetItems = await waitForCompletions(
    snippetDocument,
    new vscode.Position(3, 3),
    (items) => hasSnippetCompletion(items, "sub")
  );
  const selectSnippetItems = await waitForCompletions(
    snippetDocument,
    new vscode.Position(4, 6),
    (items) => hasSnippetCompletion(items, "select")
  );

  assert.ok(
    hasSnippetCompletion(subSnippetItems, "sub"),
    "snippet completion should include the Sub Procedure template"
  );
  assert.ok(
    hasSnippetCompletion(selectSnippetItems, "select"),
    "snippet completion should include the Select Case template"
  );

  const commands = await vscode.commands.getCommands(true);
  assert.equal(commands.includes("vba.extract"), false);
  assert.equal(commands.includes("vba.combine"), false);
}

async function waitForSymbols(document: vscode.TextDocument): Promise<readonly vscode.DocumentSymbol[]> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const symbols = await vscode.commands.executeCommand<readonly vscode.DocumentSymbol[]>(
      "vscode.executeDocumentSymbolProvider",
      document.uri
    );

    if (symbols && symbols.length > 0) {
      return symbols;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return [];
}

async function waitForCompletions(
  document: vscode.TextDocument,
  position: vscode.Position,
  predicate: (items: readonly vscode.CompletionItem[]) => boolean
): Promise<readonly vscode.CompletionItem[]> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const completionList = await vscode.commands.executeCommand<vscode.CompletionList>(
      "vscode.executeCompletionItemProvider",
      document.uri,
      position
    );

    if (completionList?.items?.length && predicate(completionList.items)) {
      return completionList.items;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return [];
}

async function waitForDefinitions(
  document: vscode.TextDocument,
  position: vscode.Position,
  predicate: (locations: readonly vscode.Location[]) => boolean
): Promise<readonly vscode.Location[]> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const definitions = await vscode.commands.executeCommand<readonly vscode.Location[]>(
      "vscode.executeDefinitionProvider",
      document.uri,
      position
    );

    if (definitions?.length && predicate(definitions)) {
      return definitions;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return [];
}

async function waitForReferences(
  document: vscode.TextDocument,
  position: vscode.Position,
  predicate: (locations: readonly vscode.Location[]) => boolean
): Promise<readonly vscode.Location[]> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const references = await vscode.commands.executeCommand<readonly vscode.Location[]>(
      "vscode.executeReferenceProvider",
      document.uri,
      position
    );

    if (references?.length && predicate(references)) {
      return references;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return [];
}

async function waitForSignatureHelp(
  document: vscode.TextDocument,
  position: vscode.Position,
  predicate: (help: vscode.SignatureHelp) => boolean
): Promise<vscode.SignatureHelp> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const signatureHelp = await vscode.commands.executeCommand<vscode.SignatureHelp>(
      "vscode.executeSignatureHelpProvider",
      document.uri,
      position
    );

    if (signatureHelp && predicate(signatureHelp)) {
      return signatureHelp;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return {
    activeParameter: 0,
    activeSignature: 0,
    signatures: []
  };
}

async function waitForRename(
  document: vscode.TextDocument,
  position: vscode.Position,
  newName: string,
  predicate: (edit: vscode.WorkspaceEdit) => boolean
): Promise<vscode.WorkspaceEdit> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const renameEdit = await vscode.commands.executeCommand<vscode.WorkspaceEdit>(
      "vscode.executeDocumentRenameProvider",
      document.uri,
      position,
      newName
    );

    if (renameEdit && predicate(renameEdit)) {
      return renameEdit;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return new vscode.WorkspaceEdit();
}

async function waitForSemanticTokensLegend(
  document: vscode.TextDocument,
  predicate: (legend: vscode.SemanticTokensLegend) => boolean
): Promise<vscode.SemanticTokensLegend> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const legend = await vscode.commands.executeCommand<vscode.SemanticTokensLegend>(
      "vscode.provideDocumentSemanticTokensLegend",
      document.uri
    );

    if (legend && predicate(legend)) {
      return legend;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return new vscode.SemanticTokensLegend([]);
}

async function waitForSemanticTokens(
  document: vscode.TextDocument,
  predicate: (tokens: vscode.SemanticTokens) => boolean
): Promise<vscode.SemanticTokens> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const tokens = await vscode.commands.executeCommand<vscode.SemanticTokens>(
      "vscode.provideDocumentSemanticTokens",
      document.uri
    );

    if (tokens && predicate(tokens)) {
      return tokens;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return new vscode.SemanticTokens(new Uint32Array());
}

async function waitForFormattedDocument(document: vscode.TextDocument, expectedText: string): Promise<string> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    await vscode.commands.executeCommand("editor.action.formatDocument");
    const currentText = document.getText();

    if (normalizeText(currentText) === normalizeText(expectedText)) {
      return currentText;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return document.getText();
}

function normalizeText(text: string): string {
  return text.replace(/\r\n?/g, "\n").trimEnd();
}

function getSignatureDocumentation(
  documentation: vscode.MarkdownString | vscode.ParameterInformation["documentation"] | vscode.SignatureInformation["documentation"] | undefined
): string {
  if (!documentation) {
    return "";
  }

  return typeof documentation === "string" ? documentation : documentation.value;
}

function getCompletionItemLabel(item: vscode.CompletionItem): string {
  return typeof item.label === "string" ? item.label : item.label.label;
}

function hasSnippetCompletion(items: readonly vscode.CompletionItem[], label: string): boolean {
  return items.some(
    (item) =>
      item.kind === vscode.CompletionItemKind.Snippet &&
      getCompletionItemLabel(item).toLowerCase() === label.toLowerCase()
  );
}
