#!/usr/bin/env node

import { writeFile } from "node:fs/promises";

import { extractWorksheetControlMetadataFromWorkbookFile } from "./lib/workbookControlMetadata.mjs";

async function main(argv) {
  const { inputPath, outputPath } = parseArguments(argv);
  const metadata = await extractWorksheetControlMetadataFromWorkbookFile(inputPath);
  const output = `${JSON.stringify(metadata, null, 2)}\n`;

  if (outputPath) {
    await writeFile(outputPath, output, "utf8");
    return;
  }

  process.stdout.write(output);
}

function parseArguments(argv) {
  const argumentsToParse = [...argv];
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

  return {
    inputPath,
    outputPath,
  };
}

function printUsage() {
  process.stdout.write(
    "使い方: node scripts/probe-workbook-control-metadata.mjs <workbook-path> [--out <json-path>]\n",
  );
}

main(process.argv.slice(2)).catch((error) => {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
});
