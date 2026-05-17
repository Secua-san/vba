import { readFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import JSZip from "jszip";

const REQUIRED_COMMANDS = ["vba.extract", "vba.combine"];
const REQUIRED_LANGUAGE_EXTENSIONS = [".bas", ".cls", ".frm"];
const REQUIRED_SETTINGS = new Map([
  ["vba.analysis.debounceMs", { defaultValue: 300, type: "number" }],
  ["vba.analysis.logPerformance", { defaultValue: false, type: "boolean" }]
]);
const REQUIRED_FIXED_FILES = [
  "extension/dist/server/index.js",
  "extension/language-configuration.json",
  "extension/resources/reference/mslearn-vba-reference.json",
  "extension/resources/vbac/vbac.wsf",
  "extension/snippets/vba.code-snippets",
  "extension/syntaxes/vba.tmLanguage.json"
];

export async function verifyLocalVsix(vsixPath) {
  const zip = await JSZip.loadAsync(await readFile(path.resolve(vsixPath)));
  const failures = [];
  const manifest = await readJsonFile(zip, "extension/package.json", failures);

  for (const filePath of REQUIRED_FIXED_FILES) {
    requireFile(zip, filePath, failures);
  }

  await verifyVbacScript(zip, failures);

  if (!manifest) {
    return failures;
  }

  requireManifestFile(zip, manifest.main, "main", failures);
  verifyConfiguration(manifest, failures);
  verifyCommands(manifest, failures);
  verifyVbaLanguage(zip, manifest, failures);
  verifyGrammarFiles(zip, manifest, failures);
  verifySnippetFiles(zip, manifest, failures);

  return failures;
}

function requireFile(zip, filePath, failures) {
  if (!zip.file(filePath)) {
    failures.push(`Missing ${filePath}`);
  }
}

function requireManifestFile(zip, manifestPath, label, failures) {
  if (typeof manifestPath !== "string" || manifestPath.length === 0) {
    failures.push(`Missing package.json ${label}`);
    return;
  }

  requireFile(zip, toExtensionZipPath(manifestPath), failures);
}

async function readJsonFile(zip, filePath, failures) {
  const file = zip.file(filePath);

  if (!file) {
    failures.push(`Missing ${filePath}`);
    return undefined;
  }

  try {
    return JSON.parse(await file.async("string"));
  } catch (error) {
    failures.push(`Invalid JSON in ${filePath}: ${String(error)}`);
    return undefined;
  }
}

async function readTextFile(zip, filePath) {
  const file = zip.file(filePath);
  return file ? file.async("string") : undefined;
}

async function verifyVbacScript(zip, failures) {
  const content = await readTextFile(zip, "extension/resources/vbac/vbac.wsf");

  if (content === undefined) {
    return;
  }

  if (!content.includes("Usage: cscript vbac.wsf") || !content.includes("decombine")) {
    failures.push("Invalid vbac helper script extension/resources/vbac/vbac.wsf");
  }
}

function verifyConfiguration(manifest, failures) {
  const properties = manifest.contributes?.configuration?.properties;

  if (!properties || typeof properties !== "object") {
    failures.push("Missing configuration contribution properties");
    return;
  }

  for (const [settingName, expected] of REQUIRED_SETTINGS) {
    const setting = properties[settingName];

    if (!setting) {
      failures.push(`Missing configuration setting ${settingName}`);
      continue;
    }

    if (setting.type !== expected.type) {
      failures.push(`Unexpected type for configuration setting ${settingName}`);
    }

    if (setting.default !== expected.defaultValue) {
      failures.push(`Unexpected default for configuration setting ${settingName}`);
    }
  }
}

function verifyCommands(manifest, failures) {
  const commandIds = new Set((manifest.contributes?.commands ?? []).map((command) => command.command));

  for (const commandId of REQUIRED_COMMANDS) {
    if (!commandIds.has(commandId)) {
      failures.push(`Missing command contribution ${commandId}`);
    }
  }
}

function verifyVbaLanguage(zip, manifest, failures) {
  const language = (manifest.contributes?.languages ?? []).find((item) => item.id === "vba");

  if (!language) {
    failures.push("Missing vba language contribution");
    return;
  }

  for (const extension of REQUIRED_LANGUAGE_EXTENSIONS) {
    if (!language.extensions?.includes(extension)) {
      failures.push(`Missing vba language extension ${extension}`);
    }
  }

  requireManifestFile(zip, language.configuration, "vba language configuration", failures);
}

function verifyGrammarFiles(zip, manifest, failures) {
  const grammars = manifest.contributes?.grammars ?? [];
  const vbaGrammar = grammars.find((grammar) => grammar.language === "vba");

  if (!vbaGrammar) {
    failures.push("Missing vba grammar contribution");
    return;
  }

  requireManifestFile(zip, vbaGrammar.path, "vba grammar path", failures);
}

function verifySnippetFiles(zip, manifest, failures) {
  const snippets = manifest.contributes?.snippets ?? [];
  const vbaSnippet = snippets.find((snippet) => snippet.language === "vba");

  if (!vbaSnippet) {
    failures.push("Missing vba snippet contribution");
    return;
  }

  requireManifestFile(zip, vbaSnippet.path, "vba snippet path", failures);
}

function toExtensionZipPath(manifestPath) {
  return `extension/${manifestPath.replace(/^\.?\//u, "")}`;
}

async function main(argv) {
  const vsixPath = argv[2];

  if (!vsixPath) {
    throw new Error("Usage: node scripts/verify-local-vsix.mjs <path-to-vsix>");
  }

  const failures = await verifyLocalVsix(vsixPath);

  if (failures.length > 0) {
    throw new Error(`VSIX verification failed:\n${failures.map((failure) => `- ${failure}`).join("\n")}`);
  }

  console.log(`VSIX verification passed: ${path.resolve(vsixPath)}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main(process.argv).catch((error) => {
    console.error(String(error instanceof Error ? error.message : error));
    process.exitCode = 1;
  });
}
