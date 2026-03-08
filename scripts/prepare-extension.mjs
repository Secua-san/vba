import { cp, mkdir, rm } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const extensionDistDir = path.join(rootDir, "packages", "extension", "dist");
const extensionResourcesDir = path.join(rootDir, "packages", "extension", "resources", "vbac");
const extensionReferenceDir = path.join(rootDir, "packages", "extension", "resources", "reference");
const serverBundleSource = path.join(rootDir, "packages", "server", "dist", "index.js");
const serverBundleTargetDir = path.join(extensionDistDir, "server");
const vbacSource = path.join(rootDir, "resources", "vbac", "vbac.wsf");
const vbacTarget = path.join(extensionResourcesDir, "vbac.wsf");
const referenceSource = path.join(rootDir, "resources", "reference", "mslearn-vba-reference.json");
const referenceTarget = path.join(extensionReferenceDir, "mslearn-vba-reference.json");

await mkdir(serverBundleTargetDir, { recursive: true });
await mkdir(extensionResourcesDir, { recursive: true });
await mkdir(extensionReferenceDir, { recursive: true });
await rm(path.join(serverBundleTargetDir, "index.js"), { force: true });
await cp(serverBundleSource, path.join(serverBundleTargetDir, "index.js"));
await cp(vbacSource, vbacTarget);
await cp(referenceSource, referenceTarget);
await mkdir(path.join(rootDir, "dist"), { recursive: true });
