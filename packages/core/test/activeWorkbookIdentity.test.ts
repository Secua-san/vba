import assert from "node:assert/strict";
import test from "node:test";
import {
  normalizeWorkbookFullNameForComparison,
  parseActiveWorkbookIdentitySnapshot
} from "../dist/index.js";

test("parseActiveWorkbookIdentitySnapshot は available snapshot を受理する", () => {
  const result = parseActiveWorkbookIdentitySnapshot({
    identity: {
      fullName: "C:/Work/Book1.xlsm",
      isAddin: false,
      name: "Book1.xlsm",
      path: "C:/Work"
    },
    observedAt: "2026-03-21T00:00:00.000Z",
    providerKind: "excel-active-workbook",
    state: "available",
    version: 1
  });

  assert.equal(result.issues.length, 0);
  assert.deepEqual(result.snapshot, {
    identity: {
      fullName: "C:/Work/Book1.xlsm",
      isAddin: false,
      name: "Book1.xlsm",
      path: "C:/Work"
    },
    observedAt: "2026-03-21T00:00:00.000Z",
    providerKind: "excel-active-workbook",
    state: "available",
    version: 1
  });
});

test("parseActiveWorkbookIdentitySnapshot は invalid payload を issue として reject する", () => {
  const result = parseActiveWorkbookIdentitySnapshot({
    identity: {
      fullName: "",
      isAddin: "no",
      name: "Book1.xlsm"
    },
    observedAt: "not-a-date",
    providerKind: "wrong-provider",
    state: "available",
    version: 2
  });

  assert.equal(result.snapshot, undefined);
  assert.equal(result.issues.some((issue) => issue.code === "invalid-version"), true);
  assert.equal(result.issues.some((issue) => issue.code === "invalid-provider-kind"), true);
  assert.equal(result.issues.some((issue) => issue.code === "invalid-observed-at"), true);
  assert.equal(result.issues.some((issue) => issue.path === "$.identity.fullName"), true);
  assert.equal(result.issues.some((issue) => issue.path === "$.identity.isAddin"), true);
});

test("parseActiveWorkbookIdentitySnapshot は available で unsaved / add-in identity を reject する", () => {
  const result = parseActiveWorkbookIdentitySnapshot({
    identity: {
      fullName: "C:/Work/Addin.xlam",
      isAddin: true,
      name: "Addin.xlam",
      path: ""
    },
    observedAt: "2026-03-21T00:00:00.000Z",
    providerKind: "excel-active-workbook",
    state: "available",
    version: 1
  });

  assert.equal(result.snapshot, undefined);
  assert.equal(result.issues.some((issue) => issue.path === "$.identity.isAddin"), true);
  assert.equal(result.issues.some((issue) => issue.path === "$.identity.path"), true);
});

test("normalizeWorkbookFullNameForComparison は Windows path 揺れを吸収する", () => {
  assert.equal(
    normalizeWorkbookFullNameForComparison("C:/Work/Book1.xlsm"),
    normalizeWorkbookFullNameForComparison("c:\\work\\BOOK1.xlsm")
  );
  assert.equal(
    normalizeWorkbookFullNameForComparison("\\\\SERVER\\Share\\Book1.xlsm"),
    normalizeWorkbookFullNameForComparison("\\\\server\\share\\book1.xlsm")
  );
});
