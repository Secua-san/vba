import { spawn } from "node:child_process";
import { copyFile, cp, mkdir, mkdtemp, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import * as vscode from "vscode";

const EXCEL_WORKBOOK_EXTENSIONS = new Set([".xls", ".xla", ".xlsb", ".xlsm", ".xlam", ".xlt", ".xltm"]);
const SOURCE_COMPONENT_EXTENSIONS = new Set([".bas", ".cls", ".dcm", ".frm"]);
export const vbacOutputChannel = vscode.window.createOutputChannel("VBA vbac");

type VbacCommand = "combine" | "decombine";

class VbacLogger {
  private readonly lines: string[] = [];

  public constructor(private readonly logPath: string) {}

  public append(line: string): void {
    this.lines.push(line);
    vbacOutputChannel.appendLine(line);
  }

  public async flush(): Promise<void> {
    await mkdir(path.dirname(this.logPath), { recursive: true });
    await writeFile(this.logPath, `${this.lines.join(os.EOL)}${os.EOL}`, "utf8");
  }
}

export async function runVbacExtract(context: vscode.ExtensionContext): Promise<void> {
  await runVbacOperation("extract", async () => {
    const workbookPath = await pickWorkbook("Select workbook to extract");
    if (!workbookPath) {
      return;
    }

    validateWorkbookPath(workbookPath);
    const sourceRoot = await pickSourceRoot("Select vbac source root");
    if (!sourceRoot) {
      return;
    }

    const timestamp = createTimestamp();
    const logPath = getLogPath("extract", workbookPath, sourceRoot, timestamp);
    const logger = new VbacLogger(logPath);
    const scriptPath = getVbacScriptPath(context);
    const workbookName = path.basename(workbookPath);
    const targetSourceDir = path.join(sourceRoot, workbookName);
    const tempRoot = await mkdtemp(path.join(os.tmpdir(), "vba-vbac-extract-"));

    try {
      await ensureVbacScript(scriptPath);
      logger.append(`operation=extract`);
      logger.append(`workbook=${workbookPath}`);
      logger.append(`sourceRoot=${sourceRoot}`);
      logger.append(`log=${logPath}`);

      if (await pathExists(targetSourceDir)) {
        const confirmed = await vscode.window.showWarningMessage(
          `Extract will replace existing VBA source folder: ${targetSourceDir}`,
          { detail: "The existing source folder will be backed up first.", modal: true },
          "Extract"
        );

        if (confirmed !== "Extract") {
          logger.append("result=cancelled");
          await logger.flush();
          return;
        }

        const sourceBackupPath = await backupSourceDirectory(targetSourceDir, sourceRoot, timestamp, logger);
        logger.append(`sourceBackup=${sourceBackupPath}`);
      }

      const tempBinaryDir = path.join(tempRoot, "bin");
      const tempSourceRoot = path.join(tempRoot, "src");
      await mkdir(tempBinaryDir, { recursive: true });
      await mkdir(tempSourceRoot, { recursive: true });
      await copyFile(workbookPath, path.join(tempBinaryDir, workbookName));

      await runVbacProcess("decombine", scriptPath, tempBinaryDir, tempSourceRoot, logger);

      const extractedSourceDir = path.join(tempSourceRoot, workbookName);
      const componentFiles = await assertSourceComponents(extractedSourceDir, "Extracted source");
      logger.append(`extractedComponentCount=${componentFiles.length}`);

      await replaceDirectory(extractedSourceDir, targetSourceDir, sourceRoot);
      const copiedComponentFiles = await assertSourceComponents(targetSourceDir, "Copied source");
      logger.append(`copiedComponentCount=${copiedComponentFiles.length}`);
      logger.append("result=success");
      await logger.flush();

      await vscode.window.showInformationMessage(`VBA extract completed. Log: ${logPath}`);
    } catch (error) {
      logger.append("result=failed");
      logger.append(`error=${error instanceof Error ? error.message : String(error)}`);
      await logger.flush().catch((flushError: unknown) => {
        vbacOutputChannel.appendLine(`failed to write vbac log: ${flushError instanceof Error ? flushError.message : String(flushError)}`);
      });
      throw error;
    } finally {
      await rm(tempRoot, { force: true, recursive: true });
    }
  });
}

export async function runVbacCombine(context: vscode.ExtensionContext): Promise<void> {
  await runVbacOperation("combine", async () => {
    const workbookPath = await pickWorkbook("Select workbook to update");
    if (!workbookPath) {
      return;
    }

    validateWorkbookPath(workbookPath);
    const sourceRoot = await pickSourceRoot("Select vbac source root");
    if (!sourceRoot) {
      return;
    }

    const workbookName = path.basename(workbookPath);
    const sourceProjectDir = path.join(sourceRoot, workbookName);
    const timestamp = createTimestamp();
    const logPath = getLogPath("combine", workbookPath, sourceRoot, timestamp);
    const logger = new VbacLogger(logPath);
    const scriptPath = getVbacScriptPath(context);
    const tempRoot = await mkdtemp(path.join(os.tmpdir(), "vba-vbac-combine-"));

    try {
      await ensureVbacScript(scriptPath);
      const componentFiles = await assertSourceComponents(sourceProjectDir, "Source");
      logger.append(`operation=combine`);
      logger.append(`workbook=${workbookPath}`);
      logger.append(`sourceRoot=${sourceRoot}`);
      logger.append(`sourceProject=${sourceProjectDir}`);
      logger.append(`sourceComponentCount=${componentFiles.length}`);
      logger.append(`log=${logPath}`);

      const confirmed = await vscode.window.showWarningMessage(
        `Combine will overwrite workbook: ${workbookPath}`,
        { detail: "A workbook backup will be created before the original file is replaced.", modal: true },
        "Combine"
      );

      if (confirmed !== "Combine") {
        logger.append("result=cancelled");
        await logger.flush();
        return;
      }

      const backupPath = await backupWorkbook(workbookPath, timestamp, logger);
      logger.append(`workbookBackup=${backupPath}`);

      const tempBinaryDir = path.join(tempRoot, "bin");
      const tempSourceRoot = path.join(tempRoot, "src");
      const tempSourceProjectDir = path.join(tempSourceRoot, workbookName);
      const tempWorkbookPath = path.join(tempBinaryDir, workbookName);
      await mkdir(tempBinaryDir, { recursive: true });
      await mkdir(tempSourceRoot, { recursive: true });
      await copyFile(workbookPath, tempWorkbookPath);
      await cp(sourceProjectDir, tempSourceProjectDir, { force: true, recursive: true });

      await runVbacProcess("combine", scriptPath, tempBinaryDir, tempSourceRoot, logger);
      await assertNonEmptyFile(tempWorkbookPath, "Combined workbook");

      const verifyBinaryDir = path.join(tempRoot, "verify-bin");
      const verifySourceRoot = path.join(tempRoot, "verify-src");
      await mkdir(verifyBinaryDir, { recursive: true });
      await mkdir(verifySourceRoot, { recursive: true });
      await copyFile(tempWorkbookPath, path.join(verifyBinaryDir, workbookName));
      await runVbacProcess("decombine", scriptPath, verifyBinaryDir, verifySourceRoot, logger);
      const verifiedComponentFiles = await assertSourceComponents(path.join(verifySourceRoot, workbookName), "Verified source");
      logger.append(`verifiedComponentCount=${verifiedComponentFiles.length}`);

      await copyFile(tempWorkbookPath, workbookPath);
      await assertNonEmptyFile(workbookPath, "Updated workbook");
      logger.append("result=success");
      await logger.flush();

      await vscode.window.showInformationMessage(`VBA combine completed. Backup: ${backupPath}. Log: ${logPath}`);
    } catch (error) {
      logger.append("result=failed");
      logger.append(`error=${error instanceof Error ? error.message : String(error)}`);
      await logger.flush().catch((flushError: unknown) => {
        vbacOutputChannel.appendLine(`failed to write vbac log: ${flushError instanceof Error ? flushError.message : String(flushError)}`);
      });
      throw error;
    } finally {
      await rm(tempRoot, { force: true, recursive: true });
    }
  });
}

async function runVbacOperation(operation: string, action: () => Promise<void>): Promise<void> {
  try {
    if (process.platform !== "win32") {
      throw new Error("vbac commands require Windows and cscript.exe.");
    }

    await action();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    vbacOutputChannel.appendLine(`${operation} failed: ${message}`);
    vbacOutputChannel.show(true);
    await vscode.window.showErrorMessage(message);
  }
}

function getVbacScriptPath(context: vscode.ExtensionContext): string {
  return context.asAbsolutePath(path.join("resources", "vbac", "vbac.wsf"));
}

async function ensureVbacScript(scriptPath: string): Promise<void> {
  const content = await readFile(scriptPath, "utf8");
  if (!content.includes("Usage: cscript vbac.wsf") || !content.includes("decombine")) {
    throw new Error(`vbac.wsf is missing or invalid: ${scriptPath}`);
  }
}

async function pickWorkbook(openLabel: string): Promise<string | undefined> {
  const selection = await vscode.window.showOpenDialog({
    canSelectFiles: true,
    canSelectFolders: false,
    canSelectMany: false,
    defaultUri: vscode.workspace.workspaceFolders?.[0]?.uri,
    filters: {
      "Excel VBA workbooks": ["xls", "xla", "xlsb", "xlsm", "xlam", "xlt", "xltm"]
    },
    openLabel
  });

  return selection?.[0]?.fsPath;
}

async function pickSourceRoot(openLabel: string): Promise<string | undefined> {
  const selection = await vscode.window.showOpenDialog({
    canSelectFiles: false,
    canSelectFolders: true,
    canSelectMany: false,
    defaultUri: vscode.workspace.workspaceFolders?.[0]?.uri,
    openLabel
  });

  return selection?.[0]?.fsPath;
}

function validateWorkbookPath(workbookPath: string): void {
  const extension = path.extname(workbookPath).toLowerCase();
  if (!EXCEL_WORKBOOK_EXTENSIONS.has(extension)) {
    throw new Error(`Unsupported Excel VBA workbook extension: ${extension || "(none)"}`);
  }
}

async function backupWorkbook(workbookPath: string, timestamp: string, logger: VbacLogger): Promise<string> {
  await assertNonEmptyFile(workbookPath, "Workbook");
  const parsed = path.parse(workbookPath);
  const backupDir = path.join(parsed.dir, ".vscode-vba", "backups");
  const backupPath = path.join(backupDir, `${parsed.name}.${timestamp}${parsed.ext}`);
  await mkdir(backupDir, { recursive: true });
  await copyFile(workbookPath, backupPath);
  await assertNonEmptyFile(backupPath, "Workbook backup");
  logger.append(`backupCreated=${backupPath}`);
  return backupPath;
}

async function backupSourceDirectory(
  sourceDir: string,
  sourceRoot: string,
  timestamp: string,
  logger: VbacLogger
): Promise<string> {
  if (!isPathInside(sourceRoot, sourceDir)) {
    throw new Error(`Refusing to back up source outside selected root: ${sourceDir}`);
  }

  const backupDir = path.join(sourceRoot, ".vscode-vba", "backups");
  const backupPath = path.join(backupDir, `${path.basename(sourceDir)}.${timestamp}`);
  await mkdir(backupDir, { recursive: true });
  await cp(sourceDir, backupPath, { force: true, recursive: true });
  logger.append(`sourceBackupCreated=${backupPath}`);
  return backupPath;
}

async function runVbacProcess(
  command: VbacCommand,
  scriptPath: string,
  binaryDir: string,
  sourceDir: string,
  logger: VbacLogger
): Promise<void> {
  const args = ["//nologo", scriptPath, command, `/binary:${binaryDir}`, `/source:${sourceDir}`];
  logger.append(`command=cscript.exe ${args.map(quoteForLog).join(" ")}`);

  const result = await new Promise<{ exitCode: number | null; signal: string | null; stderr: string; stdout: string }>(
    (resolve, reject) => {
      const child = spawn("cscript.exe", args, {
        cwd: path.dirname(scriptPath),
        windowsHide: true
      });
      const stdout: Buffer[] = [];
      const stderr: Buffer[] = [];

      child.stdout.on("data", (chunk: Buffer) => stdout.push(chunk));
      child.stderr.on("data", (chunk: Buffer) => stderr.push(chunk));
      child.on("error", reject);
      child.on("close", (exitCode, signal) => {
        resolve({
          exitCode,
          signal,
          stderr: Buffer.concat(stderr).toString("utf8"),
          stdout: Buffer.concat(stdout).toString("utf8")
        });
      });
    }
  );

  appendProcessOutput("stdout", result.stdout, logger);
  appendProcessOutput("stderr", result.stderr, logger);
  logger.append(`exitCode=${result.exitCode ?? "null"}`);
  if (result.signal) {
    logger.append(`signal=${result.signal}`);
  }

  if (result.exitCode !== 0 || result.signal) {
    throw new Error(`vbac ${command} failed with exit code ${result.exitCode ?? "null"}.`);
  }

  if (/directory '.+' not exists\.|command '.+' is undefined\./i.test(`${result.stdout}\n${result.stderr}`)) {
    throw new Error(`vbac ${command} did not complete successfully. See log for details.`);
  }
}

function appendProcessOutput(label: string, value: string, logger: VbacLogger): void {
  const trimmed = value.trim();
  if (!trimmed) {
    return;
  }

  logger.append(`${label}:`);
  for (const line of trimmed.split(/\r?\n/)) {
    logger.append(`  ${line}`);
  }
}

async function assertSourceComponents(sourceProjectDir: string, label: string): Promise<string[]> {
  const sourceStat = await stat(sourceProjectDir).catch(() => undefined);
  if (!sourceStat?.isDirectory()) {
    throw new Error(`${label} folder does not exist: ${sourceProjectDir}`);
  }

  const componentFiles = await collectSourceComponents(sourceProjectDir);
  if (componentFiles.length === 0) {
    throw new Error(`${label} folder has no VBA component files: ${sourceProjectDir}`);
  }

  return componentFiles;
}

async function collectSourceComponents(dir: string): Promise<string[]> {
  const entries = await readdir(dir, { withFileTypes: true });
  const files: string[] = [];

  for (const entry of entries) {
    const entryPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await collectSourceComponents(entryPath)));
      continue;
    }

    if (SOURCE_COMPONENT_EXTENSIONS.has(path.extname(entry.name).toLowerCase())) {
      files.push(entryPath);
    }
  }

  return files;
}

async function assertNonEmptyFile(filePath: string, label: string): Promise<void> {
  const fileStat = await stat(filePath).catch(() => undefined);
  if (!fileStat?.isFile() || fileStat.size === 0) {
    throw new Error(`${label} is missing or empty: ${filePath}`);
  }
}

async function replaceDirectory(sourceDir: string, targetDir: string, sourceRoot: string): Promise<void> {
  if (!isPathInside(sourceRoot, targetDir)) {
    throw new Error(`Refusing to replace source outside selected root: ${targetDir}`);
  }

  await rm(targetDir, { force: true, recursive: true });
  await cp(sourceDir, targetDir, { force: true, recursive: true });
}

async function pathExists(targetPath: string): Promise<boolean> {
  return stat(targetPath)
    .then(() => true)
    .catch(() => false);
}

function isPathInside(parent: string, child: string): boolean {
  const relative = path.relative(parent, child);
  return relative.length > 0 && !relative.startsWith("..") && !path.isAbsolute(relative);
}

function getLogPath(operation: string, workbookPath: string, sourceRoot: string, timestamp: string): string {
  const workspaceRoot =
    vscode.workspace.getWorkspaceFolder(vscode.Uri.file(workbookPath))?.uri.fsPath ??
    vscode.workspace.getWorkspaceFolder(vscode.Uri.file(sourceRoot))?.uri.fsPath ??
    vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ??
    path.dirname(workbookPath);
  return path.join(workspaceRoot, ".vscode-vba", "logs", `vbac-${operation}-${timestamp}.log`);
}

function createTimestamp(): string {
  return new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
}

function quoteForLog(value: string): string {
  return /\s/.test(value) ? `"${value.replace(/"/g, '\\"')}"` : value;
}
