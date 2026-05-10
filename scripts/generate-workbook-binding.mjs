#!/usr/bin/env node

import { mkdir, stat, writeFile } from "node:fs/promises";
import path from "node:path";

const addInWorkbookExtensions = new Set([".xla", ".xlam"]);
const manifestDirectoryName = ".vba";
const manifestFileName = "workbook-binding.json";

async function main(argv) {
  const { bundleRoot, workbookPath } = parseArguments(argv);
  const resolvedWorkbookPath = path.resolve(workbookPath);
  const resolvedBundleRoot = path.resolve(bundleRoot);
  const outputPath = buildWorkbookBindingManifestPath(resolvedBundleRoot);

  await assertWorkbookFile(resolvedWorkbookPath);

  if (isAddInWorkbookPath(resolvedWorkbookPath)) {
    throw new Error(".xla / .xlam workbook は add-in として扱うため workbook binding manifest を生成できません");
  }

  if (path.resolve(outputPath) === resolvedWorkbookPath) {
    throw new Error("--workbook-path には出力先 workbook-binding.json と別のパスを指定してください");
  }

  await mkdir(path.dirname(outputPath), { recursive: true });
  await writeFile(outputPath, `${JSON.stringify(createWorkbookBindingManifest(resolvedWorkbookPath), null, 2)}\n`, "utf8");
}

function buildWorkbookBindingManifestPath(bundleRoot) {
  return path.join(bundleRoot, manifestDirectoryName, manifestFileName);
}

function createWorkbookBindingManifest(workbookPath) {
  return {
    version: 1,
    artifact: "workbook-binding-manifest",
    bindingKind: "active-workbook-fullname",
    workbook: {
      fullName: workbookPath,
      name: path.basename(workbookPath),
      path: path.dirname(workbookPath),
      isAddIn: false,
      sourceKind: "openxml-package",
    },
  };
}

async function assertWorkbookFile(workbookPath) {
  let workbookStat;

  try {
    workbookStat = await stat(workbookPath);
  } catch {
    throw new Error("--workbook-path には存在する workbook file を指定してください");
  }

  if (!workbookStat.isFile()) {
    throw new Error("--workbook-path には workbook file を指定してください");
  }
}

function isAddInWorkbookPath(workbookPath) {
  return addInWorkbookExtensions.has(path.extname(workbookPath).toLowerCase());
}

function parseArguments(argv) {
  const argumentsToParse = [...argv];
  let bundleRoot;
  let workbookPath;

  while (argumentsToParse.length > 0) {
    const argument = argumentsToParse.shift();

    if (argument === "--help" || argument === "-h") {
      printUsage();
      process.exit(0);
    }

    if (argument === "--bundle-root") {
      if (bundleRoot) {
        throw new Error("--bundle-root は 1 回だけ指定してください");
      }

      bundleRoot = readOptionValue(argumentsToParse, "--bundle-root", "bundle root パス");

      continue;
    }

    if (argument === "--workbook-path") {
      if (workbookPath) {
        throw new Error("--workbook-path は 1 回だけ指定してください");
      }

      workbookPath = readOptionValue(argumentsToParse, "--workbook-path", "workbook パス");

      continue;
    }

    throw new Error(`未対応の引数です: ${argument}`);
  }

  if (!workbookPath || !bundleRoot) {
    printUsage();
    throw new Error("--workbook-path と --bundle-root が必要です");
  }

  return {
    bundleRoot,
    workbookPath,
  };
}

function readOptionValue(argumentsToParse, optionName, valueDescription) {
  const value = argumentsToParse.shift();

  if (!value || value.startsWith("-")) {
    throw new Error(`${optionName} の後に ${valueDescription}が必要です`);
  }

  return value;
}

function printUsage() {
  process.stdout.write(
    "使い方: node scripts/generate-workbook-binding.mjs --workbook-path <workbook-path> --bundle-root <bundle-root>\n",
  );
}

main(process.argv.slice(2)).catch((error) => {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
});
