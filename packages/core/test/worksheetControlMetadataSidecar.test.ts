import assert from "node:assert/strict";
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import {
  buildWorksheetControlMetadataSidecarPath,
  findNearestWorksheetControlMetadataSidecar,
  getSupportedWorksheetControlMetadataOwners,
  parseWorksheetControlMetadataSidecar
} from "../dist/index.js";

test("findNearestWorksheetControlMetadataSidecar は nearest ancestor を採用する", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-sidecar-core-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "samples", "book1");
  const moduleDirectory = path.join(bundleRoot, "modules");

  mkdirSync(moduleDirectory, { recursive: true });
  mkdirSync(path.join(workspaceRoot, ".vba"), { recursive: true });
  mkdirSync(path.join(bundleRoot, ".vba"), { recursive: true });
  writeFileSync(buildWorksheetControlMetadataSidecarPath(workspaceRoot), "{}\n");
  writeFileSync(buildWorksheetControlMetadataSidecarPath(bundleRoot), "{}\n");

  try {
    const location = findNearestWorksheetControlMetadataSidecar(path.join(moduleDirectory, "Module1.bas"), {
      workspaceRoots: [workspaceRoot]
    });

    assert.equal(location?.bundleRoot, bundleRoot);
    assert.equal(location?.sidecarPath, buildWorksheetControlMetadataSidecarPath(bundleRoot));
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("findNearestWorksheetControlMetadataSidecar は workspace root を越えない", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-sidecar-core-"));
  const outsideRoot = path.join(temporaryDirectory, "outside");
  const workspaceRoot = path.join(outsideRoot, "workspace");
  const moduleDirectory = path.join(workspaceRoot, "src");

  mkdirSync(moduleDirectory, { recursive: true });
  mkdirSync(path.join(outsideRoot, ".vba"), { recursive: true });
  writeFileSync(buildWorksheetControlMetadataSidecarPath(outsideRoot), "{}\n");

  try {
    const location = findNearestWorksheetControlMetadataSidecar(path.join(moduleDirectory, "Module1.bas"), {
      workspaceRoots: [workspaceRoot]
    });

    assert.equal(location, undefined);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("findNearestWorksheetControlMetadataSidecar は workspace root 未確定時に lookup しない", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-sidecar-core-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const moduleDirectory = path.join(workspaceRoot, "src");

  mkdirSync(moduleDirectory, { recursive: true });
  mkdirSync(path.join(workspaceRoot, ".vba"), { recursive: true });
  writeFileSync(buildWorksheetControlMetadataSidecarPath(workspaceRoot), "{}\n");

  try {
    const location = findNearestWorksheetControlMetadataSidecar(path.join(moduleDirectory, "Module1.bas"));

    assert.equal(location, undefined);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("parseWorksheetControlMetadataSidecar は unsupported owner を残しつつ supported owner を切り出せる", () => {
  const result = parseWorksheetControlMetadataSidecar(`{
    "version": 1,
    "artifact": "worksheet-control-metadata-sidecar",
    "workbook": {
      "name": "Book1.xlsm",
      "sourceKind": "openxml-package"
    },
    "owners": [
      {
        "ownerKind": "worksheet",
        "sheetName": "Sheet1",
        "sheetCodeName": "Sheet1",
        "status": "supported",
        "controls": [
          {
            "shapeName": "CheckBox1",
            "codeName": "chkFinished",
            "shapeId": 3,
            "controlType": "CheckBox",
            "progId": "Forms.CheckBox.1"
          }
        ]
      },
      {
        "ownerKind": "chartsheet",
        "sheetName": "Chart1",
        "sheetCodeName": "Chart1",
        "status": "unsupported",
        "reason": "chart-sheet-metadata-unproven"
      }
    ]
  }`);

  assert.equal(result.issues.length, 0);
  assert.equal(result.sidecar?.owners.length, 2);
  assert.deepEqual(
    getSupportedWorksheetControlMetadataOwners(result.sidecar ?? assert.fail("sidecar must be parsed")),
    [
      {
        controls: [
          {
            codeName: "chkFinished",
            controlType: "CheckBox",
            progId: "Forms.CheckBox.1",
            shapeId: 3,
            shapeName: "CheckBox1"
          }
        ],
        ownerKind: "worksheet",
        sheetCodeName: "Sheet1",
        sheetName: "Sheet1",
        status: "supported"
      }
    ]
  );
});

test("parseWorksheetControlMetadataSidecar は invalid owner/control を issue として無視する", () => {
  const result = parseWorksheetControlMetadataSidecar(`{
    "version": 1,
    "artifact": "worksheet-control-metadata-sidecar",
    "workbook": {
      "name": "Book1.xlsm",
      "sourceKind": "openxml-package"
    },
    "owners": [
      {
        "ownerKind": "worksheet",
        "sheetName": "Sheet1",
        "sheetCodeName": "Sheet1",
        "status": "supported",
        "controls": [
          {
            "shapeName": "CheckBox1",
            "shapeId": 3,
            "controlType": "CheckBox"
          },
          {
            "shapeName": "CheckBox2",
            "codeName": "chkOk",
            "shapeId": 4,
            "controlType": "CheckBox"
          }
        ]
      },
      {
        "ownerKind": "worksheet",
        "sheetName": "Broken",
        "status": "supported",
        "controls": []
      }
    ]
  }`);

  assert.equal(result.sidecar?.owners.length, 1);
  assert.equal(getSupportedWorksheetControlMetadataOwners(result.sidecar ?? assert.fail("sidecar must be parsed"))[0]?.controls.length, 1);
  assert.equal(result.issues.some((issue: { path: string }) => issue.path === "$.owners[0].controls[0].codeName"), true);
  assert.equal(result.issues.some((issue: { path: string }) => issue.path === "$.owners[1].sheetCodeName"), true);
});

test("parseWorksheetControlMetadataSidecar は top-level 不正を reject する", () => {
  const result = parseWorksheetControlMetadataSidecar(`{
    "version": 2,
    "artifact": "wrong-artifact",
    "owners": []
  }`);

  assert.equal(result.sidecar, undefined);
  assert.equal(result.issues.some((issue: { code: string }) => issue.code === "invalid-version"), true);
  assert.equal(result.issues.some((issue: { code: string }) => issue.code === "invalid-artifact"), true);
});
