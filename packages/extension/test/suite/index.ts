import assert from "node:assert/strict";
import path from "node:path";
import * as vscode from "vscode";

export async function run(): Promise<void> {
  const extension = vscode.extensions.getExtension("tagi0.vba-extension");
  assert.ok(extension, "extension must be discoverable");

  await extension.activate();

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

function getSignatureDocumentation(
  documentation: vscode.MarkdownString | vscode.ParameterInformation["documentation"] | vscode.SignatureInformation["documentation"] | undefined
): string {
  if (!documentation) {
    return "";
  }

  return typeof documentation === "string" ? documentation : documentation.value;
}
