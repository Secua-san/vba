import assert from "node:assert/strict";
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import {
  buildWorkbookBindingManifestPath,
  findNearestWorkbookBindingManifest,
  parseWorkbookBindingManifest
} from "../dist/index.js";

test("findNearestWorkbookBindingManifest は nearest ancestor を採用する", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-binding-core-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "samples", "book1");
  const moduleDirectory = path.join(bundleRoot, "modules");

  mkdirSync(moduleDirectory, { recursive: true });
  mkdirSync(path.join(workspaceRoot, ".vba"), { recursive: true });
  mkdirSync(path.join(bundleRoot, ".vba"), { recursive: true });
  writeFileSync(buildWorkbookBindingManifestPath(workspaceRoot), "{}\n");
  writeFileSync(buildWorkbookBindingManifestPath(bundleRoot), "{}\n");

  try {
    const location = findNearestWorkbookBindingManifest(path.join(moduleDirectory, "Module1.bas"), {
      workspaceRoots: [workspaceRoot]
    });

    assert.equal(location?.bundleRoot, bundleRoot);
    assert.equal(location?.manifestPath, buildWorkbookBindingManifestPath(bundleRoot));
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("parseWorkbookBindingManifest は valid manifest を受理する", () => {
  const result = parseWorkbookBindingManifest(`{
    "version": 1,
    "artifact": "workbook-binding-manifest",
    "bindingKind": "active-workbook-fullname",
    "workbook": {
      "fullName": "C:\\\\Work\\\\Book1.xlsm",
      "name": "Book1.xlsm",
      "path": "C:\\\\Work",
      "isAddIn": false,
      "sourceKind": "openxml-package"
    }
  }`);

  assert.equal(result.issues.length, 0);
  assert.equal(result.manifest?.workbook.fullName, "C:\\Work\\Book1.xlsm");
  assert.equal(result.manifest?.workbook.isAddIn, false);
});

test("parseWorkbookBindingManifest は invalid top-level を reject する", () => {
  const result = parseWorkbookBindingManifest(`{
    "version": 2,
    "artifact": "wrong-artifact",
    "bindingKind": "wrong-binding",
    "workbook": {}
  }`);

  assert.equal(result.manifest, undefined);
  assert.equal(result.issues.some((issue) => issue.code === "invalid-version"), true);
  assert.equal(result.issues.some((issue) => issue.code === "invalid-artifact"), true);
  assert.equal(result.issues.some((issue) => issue.code === "invalid-binding-kind"), true);
});

test("parseWorkbookBindingManifest は unsaved / add-in workbook を loaded 扱いにしない", () => {
  const result = parseWorkbookBindingManifest(`{
    "version": 1,
    "artifact": "workbook-binding-manifest",
    "bindingKind": "active-workbook-fullname",
    "workbook": {
      "fullName": "Addin.xlam",
      "name": "Addin.xlam",
      "path": "",
      "isAddIn": true,
      "sourceKind": "openxml-package"
    }
  }`);

  assert.equal(result.manifest, undefined);
  assert.equal(result.issues.some((issue) => issue.path === "$.workbook.path"), true);
  assert.equal(result.issues.some((issue) => issue.path === "$.workbook.isAddIn"), true);
});
