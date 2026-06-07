import { spawn } from "node:child_process";
import path from "node:path";
import * as vscode from "vscode";
import {
  ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND,
  ACTIVE_WORKBOOK_IDENTITY_VERSION,
  parseActiveWorkbookIdentitySnapshot,
  type ActiveWorkbookIdentitySnapshot
} from "../../../core/src/index";

const ACTIVE_WORKBOOK_IDENTITY_HELPER_TIMEOUT_MS = 10 * 1000;
const ACTIVE_WORKBOOK_IDENTITY_OBSERVED_AT_TOLERANCE_MS = 1000;

export const activeWorkbookIdentityOutputChannel = vscode.window.createOutputChannel("VBA Active Workbook");

export async function refreshActiveWorkbookIdentity(
  context: vscode.ExtensionContext,
  sendSnapshot: (snapshot: ActiveWorkbookIdentitySnapshot) => Promise<void>
): Promise<void> {
  await vscode.window.withProgress(
    {
      cancellable: true,
      location: vscode.ProgressLocation.Notification,
      title: "Refreshing VBA active workbook identity"
    },
    async (_progress, cancellationToken) => {
      let snapshot: ActiveWorkbookIdentitySnapshot;

      try {
        snapshot = await readActiveWorkbookIdentitySnapshot(context, cancellationToken);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        activeWorkbookIdentityOutputChannel.appendLine(`result=failed error=${message}`);
        activeWorkbookIdentityOutputChannel.show(true);

        if (isCancellationErrorMessage(message)) {
          return;
        }

        await clearActiveWorkbookIdentity(sendSnapshot);
        await vscode.window.showErrorMessage(message);
        return;
      }

      try {
        await sendSnapshot(snapshot);
        activeWorkbookIdentityOutputChannel.appendLine(
          `result=success state=${snapshot.state}${"reason" in snapshot ? ` reason=${snapshot.reason}` : ""}`
        );
        await showSnapshotMessage(snapshot);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        activeWorkbookIdentityOutputChannel.appendLine(`result=failed error=${message}`);
        activeWorkbookIdentityOutputChannel.show(true);
        await vscode.window.showErrorMessage(message);
      }
    }
  );
}

async function clearActiveWorkbookIdentity(sendSnapshot: (snapshot: ActiveWorkbookIdentitySnapshot) => Promise<void>): Promise<void> {
  try {
    await sendSnapshot({
      observedAt: new Date().toISOString(),
      providerKind: ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND,
      reason: "host-error",
      state: "unavailable",
      version: ACTIVE_WORKBOOK_IDENTITY_VERSION
    });
    activeWorkbookIdentityOutputChannel.appendLine("result=cleared state=unavailable reason=host-error");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    activeWorkbookIdentityOutputChannel.appendLine(`result=clear-failed error=${message}`);
  }
}

function isCancellationErrorMessage(message: string): boolean {
  return message === "Active workbook identity refresh was cancelled.";
}

async function readActiveWorkbookIdentitySnapshot(
  context: vscode.ExtensionContext,
  cancellationToken: vscode.CancellationToken
): Promise<ActiveWorkbookIdentitySnapshot> {
  if (process.platform !== "win32") {
    throw new Error("Active workbook identity refresh requires Windows and cscript.exe.");
  }

  activeWorkbookIdentityOutputChannel.appendLine("operation=refreshActiveWorkbookIdentity");
  activeWorkbookIdentityOutputChannel.appendLine(`timeoutMs=${ACTIVE_WORKBOOK_IDENTITY_HELPER_TIMEOUT_MS}`);

  const helperPath = context.asAbsolutePath(path.join("resources", "host", "activeWorkbookIdentity.js"));
  const startedAtMs = Date.now();
  const result = await runCscriptHelper(helperPath, cancellationToken);
  const completedAtMs = Date.now();

  activeWorkbookIdentityOutputChannel.appendLine(`exitCode=${result.exitCode ?? "null"}`);
  if (result.signal) {
    activeWorkbookIdentityOutputChannel.appendLine(`signal=${result.signal}`);
  }
  if (result.stderr.trim().length > 0) {
    activeWorkbookIdentityOutputChannel.appendLine(`stderrLength=${result.stderr.length}`);
  }

  if (result.cancelled) {
    throw new Error("Active workbook identity refresh was cancelled.");
  }
  if (result.timedOut) {
    throw new Error(`Active workbook identity helper timed out after ${ACTIVE_WORKBOOK_IDENTITY_HELPER_TIMEOUT_MS}ms.`);
  }
  if (result.exitCode !== 0 || result.signal) {
    throw new Error(`Active workbook identity helper failed with exit code ${result.exitCode ?? "null"}.`);
  }

  const snapshot = parseHelperSnapshot(result.stdout);
  assertFreshObservedAt(snapshot, startedAtMs, completedAtMs);
  return snapshot;
}

function parseHelperSnapshot(stdout: string): ActiveWorkbookIdentitySnapshot {
  const payload = stdout.replace(/^\uFEFF/u, "").trim();

  if (!payload) {
    throw new Error("Active workbook identity helper returned no payload.");
  }

  let value: unknown;
  try {
    value = JSON.parse(payload);
  } catch {
    throw new Error("Active workbook identity helper returned invalid JSON.");
  }

  const parseResult = parseActiveWorkbookIdentitySnapshot(value);
  if (!parseResult.snapshot) {
    const issueSummary = parseResult.issues.map((issue) => `${issue.path}:${issue.code}`).join(", ");
    throw new Error(`Active workbook identity helper returned invalid snapshot${issueSummary ? `: ${issueSummary}` : "."}`);
  }

  return parseResult.snapshot;
}

function assertFreshObservedAt(snapshot: ActiveWorkbookIdentitySnapshot, startedAtMs: number, completedAtMs: number): void {
  const observedAtMs = Date.parse(snapshot.observedAt);

  if (
    observedAtMs < startedAtMs - ACTIVE_WORKBOOK_IDENTITY_OBSERVED_AT_TOLERANCE_MS ||
    observedAtMs > completedAtMs + ACTIVE_WORKBOOK_IDENTITY_OBSERVED_AT_TOLERANCE_MS
  ) {
    throw new Error("Active workbook identity helper returned a stale observedAt value.");
  }
}

async function runCscriptHelper(
  helperPath: string,
  cancellationToken: vscode.CancellationToken
): Promise<{
  cancelled: boolean;
  exitCode: number | null;
  signal: string | null;
  stderr: string;
  stdout: string;
  timedOut: boolean;
}> {
  const args = ["//nologo", "//U", helperPath];
  activeWorkbookIdentityOutputChannel.appendLine("command=cscript.exe //nologo //U <helper>");

  return new Promise((resolve, reject) => {
    const child = spawn("cscript.exe", args, {
      cwd: path.dirname(helperPath),
      windowsHide: true
    });
    const stdout: Buffer[] = [];
    const stderr: Buffer[] = [];
    let cancelled = false;
    let timedOut = false;

    const timeout = setTimeout(() => {
      timedOut = true;
      child.kill();
    }, ACTIVE_WORKBOOK_IDENTITY_HELPER_TIMEOUT_MS);
    const cancellation = cancellationToken.onCancellationRequested(() => {
      cancelled = true;
      child.kill();
    });

    child.stdout.on("data", (chunk: Buffer) => stdout.push(chunk));
    child.stderr.on("data", (chunk: Buffer) => stderr.push(chunk));
    child.on("error", (error) => {
      clearTimeout(timeout);
      cancellation.dispose();
      reject(error);
    });
    child.on("close", (exitCode, signal) => {
      clearTimeout(timeout);
      cancellation.dispose();
      resolve({
        cancelled,
        exitCode,
        signal,
        stderr: Buffer.concat(stderr).toString("utf16le"),
        stdout: Buffer.concat(stdout).toString("utf16le"),
        timedOut
      });
    });
  });
}

async function showSnapshotMessage(snapshot: ActiveWorkbookIdentitySnapshot): Promise<void> {
  switch (snapshot.state) {
    case "available":
      await vscode.window.showInformationMessage("VBA active workbook identity refreshed.");
      return;
    case "protected-view":
      await vscode.window.showWarningMessage("VBA active workbook is in Protected View. Workbook binding remains disabled.");
      return;
    case "unavailable":
      await vscode.window.showWarningMessage(`VBA active workbook identity unavailable: ${snapshot.reason}.`);
      return;
    case "unsupported":
      await vscode.window.showWarningMessage(`VBA active workbook identity unsupported: ${snapshot.reason}.`);
      return;
    default:
      return;
  }
}
