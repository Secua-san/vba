#!/usr/bin/env node

import path from "node:path";
import { mkdir, writeFile } from "node:fs/promises";

import { extractWorksheetControlMetadataFromWorkbookFile } from "./lib/workbookControlMetadata.mjs";
import {
  buildWorksheetControlMetadataSidecarPath,
  convertWorksheetControlMetadataProbeToSidecar,
} from "./lib/worksheetControlMetadataSidecar.mjs";

async function main(argv) {
  const { bundleRoot, format, inputPath, outputPath } = parseArguments(argv);
  const resolvedOutputPath = outputPath ?? (bundleRoot ? buildWorksheetControlMetadataSidecarPath(bundleRoot) : undefined);

  if (resolvedOutputPath && path.resolve(resolvedOutputPath) === path.resolve(inputPath)) {
    throw new Error("--out には入力ファイルと別のパスを指定してください");
  }

  const probeMetadata = await extractWorksheetControlMetadataFromWorkbookFile(inputPath);
  const metadata =
    format === "sidecar" ? convertWorksheetControlMetadataProbeToSidecar(probeMetadata) : probeMetadata;
  const output = `${JSON.stringify(metadata, null, 2)}\n`;

  if (resolvedOutputPath) {
    await mkdir(path.dirname(resolvedOutputPath), { recursive: true });
    await writeFile(resolvedOutputPath, output, "utf8");
    return;
  }

  process.stdout.write(output);
}

function parseArguments(argv) {
  const argumentsToParse = [...argv];
  let bundleRoot;
  let format = "probe";
  let inputPath;
  let outputPath;

  while (argumentsToParse.length > 0) {
    const argument = argumentsToParse.shift();

    if (argument === "--help" || argument === "-h") {
      printUsage();
      process.exit(0);
    }

    if (argument === "--out") {
      outputPath = argumentsToParse.shift();

      if (!outputPath) {
        throw new Error("--out の後に出力先パスが必要です");
      }

      continue;
    }

    if (argument === "--bundle-root") {
      bundleRoot = argumentsToParse.shift();

      if (!bundleRoot) {
        throw new Error("--bundle-root の後に bundle root パスが必要です");
      }

      continue;
    }

    if (argument === "--format") {
      format = argumentsToParse.shift() ?? "";

      if (format !== "probe" && format !== "sidecar") {
        throw new Error("--format には probe または sidecar を指定してください");
      }

      continue;
    }

    if (!inputPath) {
      inputPath = argument;
      continue;
    }

    throw new Error(`未対応の引数です: ${argument}`);
  }

  if (!inputPath) {
    printUsage();
    throw new Error("workbook package のパスが必要です");
  }

  if (bundleRoot && format !== "sidecar") {
    throw new Error("--bundle-root は --format sidecar と一緒に指定してください");
  }

  if (bundleRoot && outputPath) {
    throw new Error("--bundle-root と --out は同時に指定できません");
  }

  return {
    bundleRoot,
    format,
    inputPath,
    outputPath,
  };
}

function printUsage() {
  process.stdout.write(
    "使い方: node scripts/probe-workbook-control-metadata.mjs <workbook-path> [--format probe|sidecar] [--out <json-path>] [--bundle-root <dir>]\n",
  );
}

main(process.argv.slice(2)).catch((error) => {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
});
