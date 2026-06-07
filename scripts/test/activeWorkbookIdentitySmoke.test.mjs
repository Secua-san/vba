import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import {
  assertExpectedSnapshot,
  main,
  parseSmokeOptions
} from "../smoke-active-workbook-identity.mjs";

test("active workbook smoke options parse explicit state and reason", () => {
  assert.deepEqual(omitHelperPath(parseSmokeOptions(["--expect-state", "unavailable", "--expect-reason", "host-unreachable"])), {
    expectFullName: undefined,
    expectProtectedSourceName: undefined,
    expectProtectedSourcePath: undefined,
    expectReason: "host-unreachable",
    expectState: "unavailable"
  });
});

test("active workbook smoke --expect-full-name implies available state", () => {
  assert.deepEqual(omitHelperPath(parseSmokeOptions(["--expect-full-name", "C:\\Work\\Book1.xlsm"])), {
    expectFullName: "C:\\Work\\Book1.xlsm",
    expectProtectedSourceName: undefined,
    expectProtectedSourcePath: undefined,
    expectReason: undefined,
    expectState: "available"
  });
});

test("active workbook smoke protected source expectations imply protected-view state", () => {
  assert.deepEqual(
    omitHelperPath(
      parseSmokeOptions([
        "--expect-protected-source-name",
        "Book1.xlsm",
        "--expect-protected-source-path",
        "C:\\Downloads"
      ])
    ),
    {
      expectFullName: undefined,
      expectProtectedSourceName: "Book1.xlsm",
      expectProtectedSourcePath: "C:\\Downloads",
      expectReason: undefined,
      expectState: "protected-view"
    }
  );
});

test("active workbook smoke options parse helper path", () => {
  const options = parseSmokeOptions(["--helper-path", "tmp\\helper.js"]);

  assert.equal(options.helperPath, path.resolve("tmp\\helper.js"));
});

test("active workbook smoke rejects invalid expectation options", () => {
  assert.throws(() => parseSmokeOptions(["--expect-state", "stale"]), /Unsupported --expect-state value/);
  assert.throws(() => parseSmokeOptions(["--expect-full-name"]), /Missing value for --expect-full-name/);
  assert.throws(
    () => parseSmokeOptions(["--expect-full-name", "C:\\Work\\Book1.xlsm", "--expect-state", "unsupported"]),
    /--expect-full-name requires --expect-state available/
  );
  assert.throws(
    () => parseSmokeOptions(["--expect-protected-source-name", "Book1.xlsm", "--expect-state", "available"]),
    /--expect-protected-source-name and --expect-protected-source-path require --expect-state protected-view/
  );
});

test("active workbook smoke validates expected available fullName with workbook normalization", () => {
  assert.doesNotThrow(() =>
    assertExpectedSnapshot(createAvailableSnapshot(), {
      expectFullName: "c:/work/book1.xlsm",
      expectProtectedSourceName: undefined,
      expectProtectedSourcePath: undefined,
      expectReason: undefined,
      expectState: "available"
    })
  );
});

test("active workbook smoke validates expected protected-view source metadata", () => {
  assert.doesNotThrow(() =>
    assertExpectedSnapshot(createProtectedViewSnapshot(), {
      expectFullName: undefined,
      expectProtectedSourceName: "Book1.xlsm",
      expectProtectedSourcePath: "C:\\Downloads",
      expectReason: undefined,
      expectState: "protected-view"
    })
  );
});

test("active workbook smoke rejects mismatched expectation", () => {
  assert.throws(
    () =>
      assertExpectedSnapshot(createAvailableSnapshot(), {
        expectFullName: "C:\\Other\\Book1.xlsm",
        expectProtectedSourceName: undefined,
        expectProtectedSourcePath: undefined,
        expectReason: undefined,
        expectState: "available"
      }),
    /fullName did not match/
  );
  assert.throws(
    () =>
      assertExpectedSnapshot(createUnavailableSnapshot(), {
        expectFullName: undefined,
        expectProtectedSourceName: undefined,
        expectProtectedSourcePath: undefined,
        expectReason: "host-error",
        expectState: "unavailable"
      }),
    /Expected active workbook identity reason=host-error/
  );
  assert.throws(
    () =>
      assertExpectedSnapshot(createProtectedViewSnapshot(), {
        expectFullName: undefined,
        expectProtectedSourceName: "OtherBook.xlsm",
        expectProtectedSourcePath: undefined,
        expectReason: undefined,
        expectState: "protected-view"
      }),
    /protectedView\.sourceName did not match/
  );
});

test("active workbook smoke main accepts synthetic available helper", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-active-workbook-smoke-"));

  try {
    const helperPath = path.join(temporaryDirectory, "activeWorkbookIdentity.js");
    await writeFile(helperPath, createSyntheticHelper(createAvailableSnapshotExpression()), "utf8");

    await main([
      "--helper-path",
      helperPath,
      "--expect-state",
      "available",
      "--expect-full-name",
      "c:/work/book1.xlsm"
    ]);
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

test("active workbook smoke main accepts synthetic protected-view helper with metadata", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-active-workbook-smoke-"));

  try {
    const helperPath = path.join(temporaryDirectory, "activeWorkbookIdentity.js");
    await writeFile(helperPath, createSyntheticHelper(createProtectedViewSnapshotExpression()), "utf8");

    await main(["--helper-path", helperPath, "--expect-state", "protected-view"]);
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

test("active workbook smoke main accepts synthetic protected-view helper with expected metadata", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-active-workbook-smoke-"));

  try {
    const helperPath = path.join(temporaryDirectory, "activeWorkbookIdentity.js");
    await writeFile(helperPath, createSyntheticHelper(createProtectedViewSnapshotExpression()), "utf8");

    await main([
      "--helper-path",
      helperPath,
      "--expect-protected-source-name",
      "Book1.xlsm",
      "--expect-protected-source-path",
      "C:\\Downloads"
    ]);
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

test("active workbook smoke main rejects synthetic stale helper", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-active-workbook-smoke-"));

  try {
    const helperPath = path.join(temporaryDirectory, "activeWorkbookIdentity.js");
    await writeFile(helperPath, createSyntheticHelper(createUnavailableSnapshotExpression("2000-01-01T00:00:00.000Z")), "utf8");

    await assert.rejects(() => main(["--helper-path", helperPath]), /stale observedAt/);
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

function omitHelperPath(options) {
  const { helperPath: _helperPath, ...rest } = options;
  return rest;
}

function createAvailableSnapshot() {
  return {
    identity: {
      fullName: "C:\\Work\\Book1.xlsm",
      isAddin: false,
      name: "Book1.xlsm",
      path: "C:\\Work"
    },
    observedAt: "2026-06-07T00:00:00.000Z",
    providerKind: "excel-active-workbook",
    state: "available",
    version: 1
  };
}

function createProtectedViewSnapshot() {
  return {
    observedAt: "2026-06-07T00:00:00.000Z",
    protectedView: {
      sourceName: "Book1.xlsm",
      sourcePath: "C:\\Downloads"
    },
    providerKind: "excel-active-workbook",
    state: "protected-view",
    version: 1
  };
}

function createSyntheticHelper(snapshotExpression) {
  return `function pad2(value) {
  return value < 10 ? "0" + value : String(value);
}
function pad3(value) {
  if (value < 10) {
    return "00" + value;
  }
  if (value < 100) {
    return "0" + value;
  }
  return String(value);
}
function observedAt() {
  var now = new Date();
  return now.getUTCFullYear() + "-" +
    pad2(now.getUTCMonth() + 1) + "-" +
    pad2(now.getUTCDate()) + "T" +
    pad2(now.getUTCHours()) + ":" +
    pad2(now.getUTCMinutes()) + ":" +
    pad2(now.getUTCSeconds()) + "." +
    pad3(now.getUTCMilliseconds()) + "Z";
}
WScript.StdOut.Write(${snapshotExpression});
`;
}

function createAvailableSnapshotExpression() {
  return `'{"version":1,"providerKind":"excel-active-workbook","state":"available","observedAt":"' + observedAt() + '","identity":{"fullName":"C:\\\\\\\\Work\\\\\\\\Book1.xlsm","name":"Book1.xlsm","path":"C:\\\\\\\\Work","isAddin":false}}'`;
}

function createProtectedViewSnapshotExpression() {
  return `'{"version":1,"providerKind":"excel-active-workbook","state":"protected-view","observedAt":"' + observedAt() + '","protectedView":{"sourceName":"Book1.xlsm","sourcePath":"C:\\\\\\\\Downloads"}}'`;
}

function createUnavailableSnapshotExpression(observedAt) {
  return JSON.stringify(
    JSON.stringify({
      observedAt,
      providerKind: "excel-active-workbook",
      reason: "host-unreachable",
      state: "unavailable",
      version: 1
    })
  );
}

function createUnavailableSnapshot() {
  return {
    observedAt: "2026-06-07T00:00:00.000Z",
    providerKind: "excel-active-workbook",
    reason: "host-unreachable",
    state: "unavailable",
    version: 1
  };
}
