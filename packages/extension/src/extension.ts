import path from "node:path";
import * as vscode from "vscode";
import { LanguageClient, LanguageClientOptions, ServerOptions, TransportKind } from "vscode-languageclient/node";
import {
  ACTIVE_WORKBOOK_IDENTITY_NOTIFICATION_METHOD,
  ACTIVE_WORKBOOK_IDENTITY_TEST_STATE_REQUEST_METHOD
} from "../../core/src/index";
import {
  TEST_GET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND,
  TEST_SET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND
} from "./testCommands";

let client: LanguageClient | undefined;

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
  await (client as LanguageClient & { onReady?: () => Promise<void> }).onReady?.();

  if (context.extensionMode === vscode.ExtensionMode.Test) {
    context.subscriptions.push(
      vscode.commands.registerCommand(TEST_SET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND, async (snapshot: unknown) => {
        if (!client) {
          throw new Error("language client is not ready");
        }

        await Promise.resolve(client.sendNotification(ACTIVE_WORKBOOK_IDENTITY_NOTIFICATION_METHOD, snapshot));
      })
    );
    context.subscriptions.push(
      vscode.commands.registerCommand(TEST_GET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND, async () => {
        if (!client) {
          throw new Error("language client is not ready");
        }

        for (let attempt = 0; attempt < 30; attempt += 1) {
          try {
            return await client.sendRequest(ACTIVE_WORKBOOK_IDENTITY_TEST_STATE_REQUEST_METHOD);
          } catch {
            await new Promise((resolve) => setTimeout(resolve, 100));
          }
        }

        return null;
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
