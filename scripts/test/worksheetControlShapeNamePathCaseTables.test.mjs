import assert from "node:assert/strict";
import test from "node:test";

import worksheetControlShapeNamePathCaseTablesModule from "../../test-support/worksheetControlShapeNamePathCaseTables.cjs";

const { worksheetControlShapeNamePathCaseTables } = worksheetControlShapeNamePathCaseTablesModule;

const OLE_FIXTURE = "packages/extension/test/fixtures/OleObjectBuiltIn.bas";
const SHAPE_FIXTURE = "packages/extension/test/fixtures/ShapesBuiltIn.bas";
const SERVER_OLE_SCOPE = "server-worksheet-control-shape-name-path-ole";
const SERVER_SHAPE_SCOPE = "server-worksheet-control-shape-name-path-shape";

test("worksheet control shapeName path completion case spec satisfies the v1 minimum coverage", () => {
  const positiveEntries = worksheetControlShapeNamePathCaseTables.worksheetControlShapeNamePath.completion.positive;
  const negativeEntries = worksheetControlShapeNamePathCaseTables.worksheetControlShapeNamePath.completion.negative;

  const positiveRouteKinds = new Set(positiveEntries.map((entry) => entry.routeKind));
  const positiveRootKinds = new Set(positiveEntries.map((entry) => entry.rootKind));
  const negativeRouteKinds = new Set(negativeEntries.map((entry) => entry.routeKind));
  const negativeReasons = new Set(negativeEntries.flatMap((entry) => (entry.reason ? [entry.reason] : [])));
  const fixtures = new Set([...positiveEntries, ...negativeEntries].map((entry) => entry.fixture));
  const scopes = new Set([...positiveEntries, ...negativeEntries].flatMap((entry) => entry.scopes));

  assert.deepEqual([...positiveRouteKinds].sort(), ["ole-object", "shape-oleformat"]);
  assert.equal(positiveRootKinds.has("document-module"), true);
  assert.equal(positiveRootKinds.has("workbook-qualified-static"), true);
  assert.equal(positiveRootKinds.has("workbook-qualified-matched"), true);
  assert.equal(negativeRouteKinds.has("ole-object"), true);
  assert.equal(negativeRouteKinds.has("shape-oleformat"), true);

  for (const reason of [
    "chartsheet-root",
    "closed-workbook",
    "code-name-selector",
    "dynamic-selector",
    "non-target-root",
    "numeric-selector",
    "plain-shape"
  ]) {
    assert.equal(negativeReasons.has(reason), true, `negative reason '${reason}' must exist`);
  }

  assert.equal(
    negativeEntries.every((entry) => typeof entry.reason === "string" && entry.reason.length > 0),
    true,
    "all negative worksheet control shapeName path completion entries must declare a reason"
  );

  assert.equal(
    negativeEntries.some((entry) => entry.routeKind === "ole-object" && entry.rootKind === "workbook-qualified-closed"),
    true
  );
  assert.equal(
    negativeEntries.some((entry) => entry.routeKind === "shape-oleformat" && entry.rootKind === "workbook-qualified-closed"),
    true
  );

  assert.deepEqual([...fixtures].sort(), [OLE_FIXTURE, SHAPE_FIXTURE]);
  assert.equal(scopes.has("extension"), true);
  assert.equal(scopes.has(SERVER_OLE_SCOPE), true);
  assert.equal(scopes.has(SERVER_SHAPE_SCOPE), true);
});

for (const interactionKind of ["hover", "signature"]) {
  test(`worksheet control shapeName path ${interactionKind} case spec satisfies the v1 minimum coverage`, () => {
    const positiveEntries = worksheetControlShapeNamePathCaseTables.worksheetControlShapeNamePath[interactionKind].positive;
    const negativeEntries = worksheetControlShapeNamePathCaseTables.worksheetControlShapeNamePath[interactionKind].negative;

    const positiveRouteKinds = new Set(positiveEntries.map((entry) => entry.routeKind));
    const positiveRootKinds = new Set(positiveEntries.map((entry) => entry.rootKind));
    const negativeRouteKinds = new Set(negativeEntries.map((entry) => entry.routeKind));
    const negativeReasons = new Set(negativeEntries.map((entry) => entry.reason));
    const fixtures = new Set([...positiveEntries, ...negativeEntries].map((entry) => entry.fixture));
    const scopes = new Set([...positiveEntries, ...negativeEntries].flatMap((entry) => entry.scopes));

    assert.deepEqual([...positiveRouteKinds].sort(), ["ole-object", "shape-oleformat"]);
    assert.equal(positiveRootKinds.has("document-module"), true);
    assert.equal(positiveRootKinds.has("workbook-qualified-static"), true);
    assert.equal(positiveRootKinds.has("workbook-qualified-matched"), true);
    assert.equal(negativeRouteKinds.has("ole-object"), true);
    assert.equal(negativeRouteKinds.has("shape-oleformat"), true);

    for (const reason of [
      "chartsheet-root",
      "closed-workbook",
      "code-name-selector",
      "dynamic-selector",
      "non-target-root",
      "numeric-selector",
      "plain-shape"
    ]) {
      assert.equal(negativeReasons.has(reason), true, `${interactionKind} negative reason '${reason}' must exist`);
    }

    assert.equal(
      negativeEntries.every((entry) => typeof entry.reason === "string" && entry.reason.length > 0),
      true,
      `all negative worksheet control shapeName path ${interactionKind} entries must declare a reason`
    );
    assert.equal(
      negativeEntries.some((entry) => entry.routeKind === "ole-object" && entry.rootKind === "workbook-qualified-closed"),
      true,
      `${interactionKind} negative entries must include a closed workbook ole-object case`
    );
    assert.equal(
      negativeEntries.some(
        (entry) => entry.routeKind === "shape-oleformat" && entry.rootKind === "workbook-qualified-closed"
      ),
      true,
      `${interactionKind} negative entries must include a closed workbook shape-oleformat case`
    );

    assert.deepEqual([...fixtures].sort(), [OLE_FIXTURE, SHAPE_FIXTURE]);
    assert.equal(scopes.has("extension"), true);
    assert.equal(scopes.has(SERVER_OLE_SCOPE), true);
    assert.equal(scopes.has(SERVER_SHAPE_SCOPE), true);
  });
}
