import * as vscode from "vscode";
import { runVbacCombine } from "./vbacShared";

export async function vbacCombine(context: vscode.ExtensionContext): Promise<void> {
  await runVbacCombine(context);
}
