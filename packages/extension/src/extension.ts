import * as vscode from "vscode";
import { vbacExtract } from "./commands/vbacExtract";
import { vbacCombine } from "./commands/vbacCombine";

export function activate(context: vscode.ExtensionContext): void {
  context.subscriptions.push(
    vscode.commands.registerCommand("vba.extract", vbacExtract),
    vscode.commands.registerCommand("vba.combine", vbacCombine)
  );
}

export function deactivate(): void {}
