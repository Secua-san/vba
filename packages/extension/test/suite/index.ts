import assert from "node:assert/strict";
import path from "node:path";
import * as vscode from "vscode";

export async function run(): Promise<void> {
  const extension = vscode.extensions.getExtension("tagi0.vba-extension");
  assert.ok(extension, "extension must be discoverable");

  await extension.activate();

  const samplePath = path.resolve(__dirname, "..", "..", "test", "fixtures", "sample.bas");
  const document = await vscode.workspace.openTextDocument(samplePath);
  await vscode.window.showTextDocument(document);

  assert.equal(document.languageId, "vba");

  const symbols = await waitForSymbols(document);
  assert.ok(symbols.length > 0, "document symbols should be available");

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
