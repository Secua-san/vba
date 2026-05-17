import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { verifyLocalVsix } from "../verify-local-vsix.mjs";

test("package は local VSIX verifier の npm script を公開する", async () => {
  const packageManifest = JSON.parse(await readFile(path.resolve("package.json"), "utf8"));

  assert.equal(packageManifest.scripts["verify:vsix"], "node scripts/verify-local-vsix.mjs");
});

test("local VSIX verifier は first release に必要な contribution と files を受理する", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-vsix-verify-"));

  try {
    const vsixPath = await writeSyntheticVsix(temporaryDirectory);
    const failures = await verifyLocalVsix(vsixPath);

    assert.deepEqual(failures, []);
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

test("local VSIX verifier は不足 command と missing asset を報告する", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-vsix-verify-"));

  try {
    const vsixPath = await writeSyntheticVsix(temporaryDirectory, {
      omitFiles: ["extension/resources/vbac/vbac.wsf"],
      transformManifest: (manifest) => ({
        ...manifest,
        contributes: {
          ...manifest.contributes,
          commands: manifest.contributes.commands.filter((command) => command.command !== "vba.combine")
        }
      })
    });
    const failures = await verifyLocalVsix(vsixPath);

    assert.deepEqual(failures, [
      "Missing extension/resources/vbac/vbac.wsf",
      "Missing command contribution vba.combine"
    ]);
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

test("local VSIX verifier は invalid helper と missing setting を報告する", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-vsix-verify-"));

  try {
    const vsixPath = await writeSyntheticVsix(temporaryDirectory, {
      fileOverrides: {
        "extension/resources/vbac/vbac.wsf": "<job />"
      },
      transformManifest: (manifest) => ({
        ...manifest,
        contributes: {
          ...manifest.contributes,
          configuration: {
            ...manifest.contributes.configuration,
            properties: {
              "vba.analysis.debounceMs": manifest.contributes.configuration.properties["vba.analysis.debounceMs"]
            }
          }
        }
      })
    });
    const failures = await verifyLocalVsix(vsixPath);

    assert.deepEqual(failures, [
      "Invalid vbac helper script extension/resources/vbac/vbac.wsf",
      "Missing configuration setting vba.analysis.logPerformance"
    ]);
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

async function writeSyntheticVsix(temporaryDirectory, options = {}) {
  const zip = new JSZip();
  const omitFiles = new Set(options.omitFiles ?? []);
  const manifest = options.transformManifest?.(createManifest()) ?? createManifest();
  const files = {
    "extension/dist/extension.js": "extension bundle",
    "extension/dist/server/index.js": "server bundle",
    "extension/language-configuration.json": "{}",
    "extension/package.json": JSON.stringify(manifest),
    "extension/resources/reference/mslearn-vba-reference.json": "{}",
    "extension/resources/vbac/vbac.wsf": "Usage: cscript vbac.wsf\nCommands:\n  decombine",
    "extension/snippets/vba.code-snippets": "{}",
    "extension/syntaxes/vba.tmLanguage.json": "{}",
    ...(options.fileOverrides ?? {})
  };

  for (const [filePath, content] of Object.entries(files)) {
    if (!omitFiles.has(filePath)) {
      zip.file(filePath, content);
    }
  }

  const vsixPath = path.join(temporaryDirectory, "vba-extension.vsix");

  await writeFile(vsixPath, await zip.generateAsync({ type: "nodebuffer" }));
  return vsixPath;
}

function createManifest() {
  return {
    main: "dist/extension.js",
    contributes: {
      commands: [
        {
          command: "vba.extract",
          title: "VBA: Extract Source with vbac"
        },
        {
          command: "vba.combine",
          title: "VBA: Combine Source with vbac"
        }
      ],
      grammars: [
        {
          language: "vba",
          path: "./syntaxes/vba.tmLanguage.json"
        }
      ],
      languages: [
        {
          configuration: "./language-configuration.json",
          extensions: [".bas", ".cls", ".frm"],
          id: "vba"
        }
      ],
      snippets: [
        {
          language: "vba",
          path: "./snippets/vba.code-snippets"
        }
      ],
      configuration: {
        properties: {
          "vba.analysis.debounceMs": {
            default: 300,
            type: "number"
          },
          "vba.analysis.logPerformance": {
            default: false,
            type: "boolean"
          }
        }
      }
    }
  };
}
