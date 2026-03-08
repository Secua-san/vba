import path from "node:path";
import * as vscode from "vscode";
import { LanguageClient, LanguageClientOptions, ServerOptions, TransportKind } from "vscode-languageclient/node";

let client: LanguageClient | undefined;

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  const serverModule = context.asAbsolutePath(path.join("dist", "server", "index.js"));
  const serverOptions: ServerOptions = {
    debug: {
      module: serverModule,
      options: {
        execArgv: ["--nolazy", "--inspect=6010"]
      },
      transport: TransportKind.ipc
    },
    run: {
      module: serverModule,
      transport: TransportKind.ipc
    }
  };
  const clientOptions: LanguageClientOptions = {
    documentSelector: [
      {
        language: "vba",
        scheme: "file"
      }
    ],
    synchronize: {
      configurationSection: "vba"
    }
  };

  client = new LanguageClient("excelVbaLanguageServer", "Excel VBA Language Server", serverOptions, clientOptions);
  context.subscriptions.push(client);
  await client.start();
}

export async function deactivate(): Promise<void> {
  if (client) {
    await client.stop();
    client = undefined;
  }
}
