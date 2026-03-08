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

  const consumerDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "Consumer.bas"));
  await vscode.window.showTextDocument(consumerDocument);

  const completionItems = await waitForCompletions(
    consumerDocument,
    new vscode.Position(5, 4),
    (items) => items.some((item) => item.label === "PublicMessage")
  );
  assert.ok(
    completionItems.some((item) => item.label === "PublicMessage"),
    "cross-file completion should include exported workspace symbols"
  );

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
