import * as vscode from "vscode";
import { runVbacExtract } from "./vbacShared";

export async function vbacExtract(context: vscode.ExtensionContext): Promise<void> {
  await runVbacExtract(context);
}
