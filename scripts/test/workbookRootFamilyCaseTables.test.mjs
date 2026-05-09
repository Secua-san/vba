import assert from "node:assert/strict";
import test from "node:test";

import workbookRootFamilyCaseTablesModule from "../../test-support/workbookRootFamilyCaseTables.cjs";

const { workbookRootFamilyCaseTables } = workbookRootFamilyCaseTablesModule;

const APPLICATION_SERVER_SCOPES = [
  "server-application-ole",
  "server-application-shadowed",
  "server-application-shape"
];
const WORKSHEET_BROAD_ROOT_SERVER_SCOPES = [
  "server-worksheet-broad-root-direct",
  "server-worksheet-broad-root-item"
];

function uniqueSorted(values) {
  return [...new Set(values)].sort();
}

function scopesFor(entries) {
  return uniqueSorted(entries.flatMap((entry) => entry.scopes));
}

function reasonsFor(entries) {
  return uniqueSorted(entries.flatMap((entry) => (entry.reason ? [entry.reason] : [])));
}

function statesFor(entries) {
  return uniqueSorted(entries.flatMap((entry) => (entry.state ? [entry.state] : [])));
}

function routesFor(entries) {
  return uniqueSorted(entries.flatMap((entry) => (entry.route ? [entry.route] : [])));
}

function assertContainsAll(actualValues, expectedValues, messagePrefix) {
  for (const expectedValue of expectedValues) {
    assert.equal(actualValues.includes(expectedValue), true, `${messagePrefix} must include ${expectedValue}`);
  }
}

function hasOnlyExtensionScope(entry) {
  return entry.scopes.length === 1 && entry.scopes[0] === "extension";
}

test("application workbook root case spec satisfies the v1 mirror coverage", () => {
  const table = workbookRootFamilyCaseTables.applicationWorkbookRoot;

  for (const surface of ["completion", "hover", "signature", "semantic"]) {
    assert.ok(table[surface].positive.length > 0, `application ${surface} positive cases must not be empty`);
    assert.ok(table[surface].negative.length > 0, `application ${surface} negative cases must not be empty`);
  }

  assert.deepEqual(routesFor(table.completion.positive), ["ole", "shape"]);
  assert.deepEqual(statesFor(table.completion.positive), ["matched", "static"]);
  assertContainsAll(
    scopesFor(table.completion.positive),
    ["extension", "server-application-ole", "server-application-shape"],
    "application completion positive scopes"
  );

  assertContainsAll(
    reasonsFor(table.completion.negative),
    [
      "code-name-selector",
      "dynamic-selector",
      "non-target-root",
      "numeric-selector",
      "shadowed-root",
      "snapshot-closed"
    ],
    "application completion negative reasons"
  );
  assertContainsAll(
    statesFor(table.completion.negative),
    ["closed", "matched", "shadowed", "static"],
    "application states"
  );
  assertContainsAll(
    scopesFor(table.completion.negative),
    APPLICATION_SERVER_SCOPES,
    "application completion negative scopes"
  );
  assert.equal(
    table.completion.negative.some(hasOnlyExtensionScope),
    false,
    "application completion negative entries must stay mirrored by server scope"
  );

  for (const surface of ["hover", "signature"]) {
    assertContainsAll(
      scopesFor(table[surface].positive),
      ["server-application-ole", "server-application-shape"],
      `${surface} positive scopes`
    );
    assertContainsAll(reasonsFor(table[surface].negative), ["shadowed-root", "snapshot-closed"], `${surface} negative reasons`);
    assertContainsAll(scopesFor(table[surface].negative), APPLICATION_SERVER_SCOPES, `${surface} negative scopes`);
  }

  assertContainsAll(
    scopesFor(table.semantic.positive),
    ["server-application-ole", "server-application-shape"],
    "semantic positive scopes"
  );
  assert.equal(
    table.semantic.positive.every((entry) => typeof entry.identifier === "string" && entry.identifier.length > 0),
    true,
    "application semantic positive entries must declare an identifier"
  );
  assert.equal(
    table.semantic.negative.every(
      (entry) =>
        typeof entry.reason === "string" &&
        entry.reason.length > 0 &&
        typeof entry.identifier === "string" &&
        entry.identifier.length > 0
    ),
    true,
    "application semantic negative entries must declare reason and identifier"
  );
});

test("application workbook root extension-only residual stays outside completion", () => {
  const table = workbookRootFamilyCaseTables.applicationWorkbookRoot;
  const extensionOnlyResiduals = [];

  for (const surface of ["hover", "signature", "semantic"]) {
    for (const entry of table[surface].negative) {
      if (hasOnlyExtensionScope(entry)) {
        extensionOnlyResiduals.push([surface, entry.anchor]);
      }
    }
  }

  assert.deepEqual(extensionOnlyResiduals, [
    ["hover", 'Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Valu'],
    ["hover", 'Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu'],
    ["signature", 'Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Select('],
    ["signature", 'Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select('],
    ["signature", 'Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select('],
    ["semantic", 'Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value'],
    ["semantic", 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value']
  ]);
});

test("worksheet broad root case spec satisfies the v1 shared coverage", () => {
  const table = workbookRootFamilyCaseTables.worksheetBroadRoot;

  for (const surface of ["completion", "hover", "signature"]) {
    assert.ok(table[surface].positive.length > 0, `worksheet broad root ${surface} positive cases must not be empty`);
    assert.ok(table[surface].negative.length > 0, `worksheet broad root ${surface} negative cases must not be empty`);
  }

  assert.deepEqual(routesFor(table.completion.positive), ["ole", "shape"]);
  assertContainsAll(
    scopesFor(table.completion.positive),
    ["extension", ...WORKSHEET_BROAD_ROOT_SERVER_SCOPES],
    "completion positive scopes"
  );
  assertContainsAll(
    scopesFor(table.completion.negative),
    ["extension", ...WORKSHEET_BROAD_ROOT_SERVER_SCOPES],
    "completion negative scopes"
  );
  assertContainsAll(
    reasonsFor(table.completion.negative),
    ["dynamic-selector", "non-target-root", "numeric-selector"],
    "completion negative reasons"
  );

  for (const surface of ["hover", "signature"]) {
    assertContainsAll(
      scopesFor(table[surface].positive),
      ["extension", ...WORKSHEET_BROAD_ROOT_SERVER_SCOPES],
      `${surface} positive scopes`
    );
    assert.deepEqual(scopesFor(table[surface].negative), ["extension"]);
    assertContainsAll(
      reasonsFor(table[surface].negative),
      ["dynamic-selector", "non-target-root", "numeric-selector"],
      `${surface} negative reasons`
    );
  }
});
