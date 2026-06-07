import { spawn } from "node:child_process";
import path from "node:path";
import { pathToFileURL } from "node:url";
import {
  normalizeWorkbookFullNameForComparison,
  parseActiveWorkbookIdentitySnapshot
} from "../packages/core/dist/index.js";

const HELPER_TIMEOUT_MS = 10 * 1000;
const OBSERVED_AT_TOLERANCE_MS = 1000;
const VALID_STATES = new Set(["available", "protected-view", "unavailable", "unsupported"]);
const DEFAULT_HELPER_PATH = path.resolve("packages/extension/resources/host/activeWorkbookIdentity.js");

export async function main(argv = process.argv.slice(2)) {
  if (process.platform !== "win32") {
    throw new Error("Active workbook identity smoke requires Windows and cscript.exe.");
  }

  const options = parseSmokeOptions(argv);
  const startedAtMs = Date.now();
  const result = await runHelper(options.helperPath);
  const completedAtMs = Date.now();

  if (result.timedOut) {
    throw new Error(`Active workbook identity helper timed out after ${HELPER_TIMEOUT_MS}ms.`);
  }

  if (result.exitCode !== 0 || result.signal) {
    throw new Error(`Active workbook identity helper failed with exit code ${result.exitCode ?? "null"}.`);
  }

  const payload = result.stdout.replace(/^\uFEFF/u, "").trim();
  if (!payload) {
    throw new Error("Active workbook identity helper returned no payload.");
  }

  let parsed;
  try {
    parsed = JSON.parse(payload);
  } catch {
    throw new Error("Active workbook identity helper returned invalid JSON.");
  }

  const parseResult = parseActiveWorkbookIdentitySnapshot(parsed);
  if (!parseResult.snapshot) {
    const issueSummary = parseResult.issues.map((issue) => `${issue.path}:${issue.code}`).join(", ");
    throw new Error(`Active workbook identity helper returned invalid snapshot${issueSummary ? `: ${issueSummary}` : "."}`);
  }

  const observedAtMs = Date.parse(parseResult.snapshot.observedAt);
  if (
    observedAtMs < startedAtMs - OBSERVED_AT_TOLERANCE_MS ||
    observedAtMs > completedAtMs + OBSERVED_AT_TOLERANCE_MS
  ) {
    throw new Error("Active workbook identity helper returned a stale observedAt value.");
  }

  assertExpectedSnapshot(parseResult.snapshot, options);

  console.log(
    [
      "Active workbook identity smoke passed:",
      `state=${parseResult.snapshot.state}`,
      "reason" in parseResult.snapshot ? `reason=${parseResult.snapshot.reason}` : undefined
    ]
      .filter(Boolean)
      .join(" ")
  );
}

export function parseSmokeOptions(argv) {
  const options = {
    expectFullName: undefined,
    expectProtectedSourceName: undefined,
    expectProtectedSourcePath: undefined,
    expectReason: undefined,
    expectState: undefined,
    helperPath: DEFAULT_HELPER_PATH
  };

  for (let index = 0; index < argv.length; index += 1) {
    const flag = argv[index];

    switch (flag) {
      case "--expect-full-name":
        options.expectFullName = readOptionValue(argv, index, flag);
        options.expectState ??= "available";
        index += 1;
        break;
      case "--expect-reason":
        options.expectReason = readOptionValue(argv, index, flag);
        index += 1;
        break;
      case "--expect-protected-source-name":
        options.expectProtectedSourceName = readOptionValue(argv, index, flag);
        options.expectState ??= "protected-view";
        index += 1;
        break;
      case "--expect-protected-source-path":
        options.expectProtectedSourcePath = readOptionValue(argv, index, flag);
        options.expectState ??= "protected-view";
        index += 1;
        break;
      case "--helper-path":
        options.helperPath = path.resolve(readOptionValue(argv, index, flag));
        index += 1;
        break;
      case "--expect-state": {
        const state = readOptionValue(argv, index, flag);
        if (!VALID_STATES.has(state)) {
          throw new Error(`Unsupported --expect-state value: ${state}`);
        }
        options.expectState = state;
        index += 1;
        break;
      }
      default:
        throw new Error(`Unknown option: ${flag}`);
    }
  }

  if (options.expectFullName && options.expectState !== "available") {
    throw new Error("--expect-full-name requires --expect-state available.");
  }
  if (
    (options.expectProtectedSourceName || options.expectProtectedSourcePath) &&
    options.expectState !== "protected-view"
  ) {
    throw new Error("--expect-protected-source-name and --expect-protected-source-path require --expect-state protected-view.");
  }

  return options;
}

export function assertExpectedSnapshot(snapshot, options) {
  if (options.expectState && snapshot.state !== options.expectState) {
    throw new Error(`Expected active workbook identity state=${options.expectState}, got state=${snapshot.state}.`);
  }

  if (options.expectReason) {
    if (!("reason" in snapshot)) {
      throw new Error("Expected active workbook identity reason, but snapshot has no reason.");
    }

    if (snapshot.reason !== options.expectReason) {
      throw new Error(`Expected active workbook identity reason=${options.expectReason}, got reason=${snapshot.reason}.`);
    }
  }

  if (options.expectFullName) {
    if (snapshot.state !== "available") {
      throw new Error("Expected available active workbook identity with matching fullName.");
    }

    if (
      normalizeWorkbookFullNameForComparison(snapshot.identity.fullName) !==
      normalizeWorkbookFullNameForComparison(options.expectFullName)
    ) {
      throw new Error("Active workbook identity fullName did not match --expect-full-name.");
    }
  }

  if (options.expectProtectedSourceName || options.expectProtectedSourcePath) {
    if (snapshot.state !== "protected-view") {
      throw new Error("Expected protected-view active workbook identity with matching source metadata.");
    }

    if (options.expectProtectedSourceName && snapshot.protectedView?.sourceName !== options.expectProtectedSourceName) {
      throw new Error("Active workbook identity protectedView.sourceName did not match --expect-protected-source-name.");
    }

    if (options.expectProtectedSourcePath && snapshot.protectedView?.sourcePath !== options.expectProtectedSourcePath) {
      throw new Error("Active workbook identity protectedView.sourcePath did not match --expect-protected-source-path.");
    }
  }
}

function readOptionValue(argv, index, flag) {
  const value = argv[index + 1];

  if (!value || value.startsWith("--")) {
    throw new Error(`Missing value for ${flag}.`);
  }

  return value;
}

function runHelper(scriptPath) {
  return new Promise((resolve, reject) => {
    const child = spawn("cscript.exe", ["//nologo", "//U", scriptPath], {
      cwd: path.dirname(scriptPath),
      windowsHide: true
    });
    const stdout = [];
    const stderr = [];
    let timedOut = false;

    const timeout = setTimeout(() => {
      timedOut = true;
      child.kill();
    }, HELPER_TIMEOUT_MS);

    child.stdout.on("data", (chunk) => stdout.push(chunk));
    child.stderr.on("data", (chunk) => stderr.push(chunk));
    child.on("error", (error) => {
      clearTimeout(timeout);
      reject(error);
    });
    child.on("close", (exitCode, signal) => {
      clearTimeout(timeout);
      resolve({
        exitCode,
        signal,
        stderr: Buffer.concat(stderr).toString("utf16le"),
        stdout: Buffer.concat(stdout).toString("utf16le"),
        timedOut
      });
    });
  });
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(String(error instanceof Error ? error.message : error));
    process.exitCode = 1;
  });
}
