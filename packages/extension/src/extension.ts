import path from "node:path";
import * as vscode from "vscode";
import { LanguageClient, LanguageClientOptions, ServerOptions, TransportKind } from "vscode-languageclient/node";

let client: LanguageClient | undefined;
const ACTIVE_WORKBOOK_IDENTITY_NOTIFICATION = "vba/activeWorkbookIdentity";
const TEST_SET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND = "vba.test.setActiveWorkbookIdentitySnapshot";

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  const serverModule = context.asAbsolutePath(path.join("dist", "server", "index.js"));
  const fileWatcher = vscode.workspace.createFileSystemWatcher("**/*.{bas,cls,frm}");
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
      configurationSection: "vba",
      fileEvents: fileWatcher
    }
  };

  client = new LanguageClient("excelVbaLanguageServer", "Excel VBA Language Server", serverOptions, clientOptions);
  context.subscriptions.push(fileWatcher);
  context.subscriptions.push(client);
  await client.start();

  if (context.extensionMode === vscode.ExtensionMode.Test) {
    context.subscriptions.push(
      vscode.commands.registerCommand(TEST_SET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND, async (snapshot: unknown) => {
        if (!client) {
          throw new Error("language client is not ready");
        }

        await Promise.resolve(client.sendNotification(ACTIVE_WORKBOOK_IDENTITY_NOTIFICATION, snapshot));
      })
    );
  }
}

export async function deactivate(): Promise<void> {
  if (client) {
    await client.stop();
    client = undefined;
  }
}
