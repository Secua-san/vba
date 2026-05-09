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
const APPLICATION_EXTENSION_ONLY_RESIDUALS = [
  ["hover", 'Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Valu'],
  ["hover", 'Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu'],
  ["signature", 'Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Select('],
  ["signature", 'Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select('],
  ["signature", 'Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select('],
  ["semantic", 'Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value'],
  ["semantic", 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value']
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

function assertEntriesHaveAnyScope(entries, expectedScopes, messagePrefix) {
  for (const entry of entries) {
    assert.equal(
      expectedScopes.some((scope) => entry.scopes.includes(scope)),
      true,
      `${messagePrefix}: ${entry.anchor} must include one of ${expectedScopes.join(", ")}`
    );
  }
}

function hasOnlyExtensionScope(entry) {
  return entry.scopes.length === 1 && entry.scopes[0] === "extension";
}

function residualKey([surface, anchor]) {
  return `${surface}\0${anchor}`;
}

function sortResiduals(residuals) {
  return [...residuals].sort((left, right) => residualKey(left).localeCompare(residualKey(right)));
}

test("application workbook root case spec satisfies the v1 mirror coverage", () => {
  const table = workbookRootFamilyCaseTables.applicationWorkbookRoot;
  const extensionOnlyResidualKeys = new Set(APPLICATION_EXTENSION_ONLY_RESIDUALS.map(residualKey));

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
  for (const entry of table.completion.positive) {
    const expectedScope = entry.route === "ole" ? "server-application-ole" : "server-application-shape";
    assert.equal(entry.scopes.includes(expectedScope), true, `${entry.anchor} must include ${expectedScope}`);
  }

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
  assertEntriesHaveAnyScope(
    table.completion.negative,
    APPLICATION_SERVER_SCOPES,
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
    assertEntriesHaveAnyScope(
      table[surface].positive,
      ["server-application-ole", "server-application-shape"],
      `${surface} positive entries must stay mirrored by server scope`
    );
    assertEntriesHaveAnyScope(
      table[surface].negative.filter((entry) => !extensionOnlyResidualKeys.has(residualKey([surface, entry.anchor]))),
      APPLICATION_SERVER_SCOPES,
      `${surface} negative non-residual entries must stay mirrored by server scope`
    );
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
  assertEntriesHaveAnyScope(
    table.semantic.negative.filter((entry) => !extensionOnlyResidualKeys.has(residualKey(["semantic", entry.anchor]))),
    ["server-application-ole", "server-application-shape"],
    "semantic negative non-residual entries must stay mirrored by server scope"
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

  assert.deepEqual(sortResiduals(extensionOnlyResiduals), sortResiduals(APPLICATION_EXTENSION_ONLY_RESIDUALS));
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
  assertEntriesHaveAnyScope(
    table.completion.positive,
    WORKSHEET_BROAD_ROOT_SERVER_SCOPES,
    "worksheet broad root completion positive entries must stay mirrored by server scope"
  );
  assertEntriesHaveAnyScope(
    table.completion.negative,
    WORKSHEET_BROAD_ROOT_SERVER_SCOPES,
    "worksheet broad root completion negative entries must stay mirrored by server scope"
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
