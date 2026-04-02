import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import * as vscode from "vscode";
import {
  TEST_GET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND,
  TEST_SET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND
} from "../../src/testCommands";

type WorkbookRootFamilyScope =
  | "extension"
  | "server-application-ole"
  | "server-application-shadowed"
  | "server-application-shape"
  | "server-worksheet-broad-root-direct"
  | "server-worksheet-broad-root-item";
type WorkbookRootFamilyState = "closed" | "matched" | "shadowed" | "static";
type WorkbookRootFamilyRoute = "ole" | "shape";
type WorkbookRootFamilyRootKind = "ActiveWorkbook" | "ThisWorkbook";
type WorkbookRootFamilySemanticFamilyName = "applicationWorkbookRoot";
type WorkbookRootFamilyReason =
  | "code-name-selector"
  | "dynamic-selector"
  | "non-target-root"
  | "numeric-selector"
  | "shadowed-root"
  | "snapshot-closed";

type WorkbookRootFamilyCaseEntryBase = {
  anchor: string;
  occurrenceIndex?: number;
  reason?: WorkbookRootFamilyReason;
  rootKind?: WorkbookRootFamilyRootKind;
  scopes: readonly WorkbookRootFamilyScope[];
  state?: WorkbookRootFamilyState;
};

type WorkbookRootFamilyPositiveCompletionEntry = WorkbookRootFamilyCaseEntryBase & {
  route: WorkbookRootFamilyRoute;
};

type WorkbookRootFamilyNegativeCompletionEntry = WorkbookRootFamilyCaseEntryBase;
type WorkbookRootFamilyInteractionEntry = WorkbookRootFamilyCaseEntryBase;

type WorkbookRootFamilySemanticEntry = {
  anchor: string;
  identifier: string;
  occurrenceIndex?: number;
  reason?: WorkbookRootFamilyReason;
  scopes: readonly WorkbookRootFamilyScope[];
  state: WorkbookRootFamilyState;
  tokenKind: "method" | "property";
};

type WorkbookRootFamilyCaseTables = {
  applicationWorkbookRoot: {
    completion: {
      negative: readonly WorkbookRootFamilyNegativeCompletionEntry[];
      positive: readonly WorkbookRootFamilyPositiveCompletionEntry[];
    };
    hover: {
      negative: readonly WorkbookRootFamilyInteractionEntry[];
      positive: readonly WorkbookRootFamilyInteractionEntry[];
    };
    semantic: {
      negative: readonly WorkbookRootFamilySemanticEntry[];
      positive: readonly WorkbookRootFamilySemanticEntry[];
    };
    signature: {
      negative: readonly WorkbookRootFamilyInteractionEntry[];
      positive: readonly WorkbookRootFamilyInteractionEntry[];
    };
  };
  worksheetBroadRoot: {
    completion: {
      negative: readonly WorkbookRootFamilyNegativeCompletionEntry[];
      positive: readonly WorkbookRootFamilyPositiveCompletionEntry[];
    };
    hover: {
      negative: readonly WorkbookRootFamilyInteractionEntry[];
      positive: readonly WorkbookRootFamilyInteractionEntry[];
    };
    signature: {
      negative: readonly WorkbookRootFamilyInteractionEntry[];
      positive: readonly WorkbookRootFamilyInteractionEntry[];
    };
  };
};

type WorksheetControlShapeNamePathScope =
  | "extension"
  | "server-worksheet-control-shape-name-path-ole"
  | "server-worksheet-control-shape-name-path-shape";
type WorksheetControlShapeNamePathRootKind =
  | "document-module"
  | "workbook-qualified-closed"
  | "workbook-qualified-matched"
  | "workbook-qualified-static";
type WorksheetControlShapeNamePathRouteKind = "ole-object" | "shape-oleformat";
type WorksheetControlShapeNamePathInteractionKind = "hover" | "signature";
type WorksheetControlShapeNamePathSemanticTokenKind = "method" | "property";
type WorksheetControlShapeNamePathReason =
  | "chartsheet-root"
  | "closed-workbook"
  | "code-name-selector"
  | "dynamic-selector"
  | "non-target-root"
  | "numeric-selector"
  | "plain-shape";
type WorksheetControlShapeNamePathFixture =
  | "packages/extension/test/fixtures/OleObjectBuiltIn.bas"
  | "packages/extension/test/fixtures/ShapesBuiltIn.bas";
type WorksheetControlShapeNamePathCaseEntryBase = {
  anchor: string;
  fixture: WorksheetControlShapeNamePathFixture;
  occurrenceIndex?: number;
  rootKind: WorksheetControlShapeNamePathRootKind;
  routeKind: WorksheetControlShapeNamePathRouteKind;
  scopes: readonly WorksheetControlShapeNamePathScope[];
};
type WorksheetControlShapeNamePathPositiveEntry = WorksheetControlShapeNamePathCaseEntryBase;
type WorksheetControlShapeNamePathNegativeEntry = WorksheetControlShapeNamePathCaseEntryBase & {
  reason: WorksheetControlShapeNamePathReason;
};
type WorksheetControlShapeNamePathSemanticEntry = WorksheetControlShapeNamePathCaseEntryBase & {
  identifier: string;
  reason?: WorksheetControlShapeNamePathReason;
  tokenKind: WorksheetControlShapeNamePathSemanticTokenKind;
};
type WorksheetControlShapeNamePathCaseTables = {
  worksheetControlShapeNamePath: {
    completion: {
      negative: readonly WorksheetControlShapeNamePathNegativeEntry[];
      positive: readonly WorksheetControlShapeNamePathPositiveEntry[];
    };
    hover: {
      negative: readonly WorksheetControlShapeNamePathNegativeEntry[];
      positive: readonly WorksheetControlShapeNamePathPositiveEntry[];
    };
    signature: {
      negative: readonly WorksheetControlShapeNamePathNegativeEntry[];
      positive: readonly WorksheetControlShapeNamePathPositiveEntry[];
    };
    semantic: {
      negative: readonly WorksheetControlShapeNamePathSemanticEntry[];
      positive: readonly WorksheetControlShapeNamePathSemanticEntry[];
    };
  };
};

const requireFromHere = createRequire(__filename);
const workbookRootFamilyCaseTables = loadWorkbookRootFamilyCaseTables();
const worksheetControlShapeNamePathCaseTables = loadWorksheetControlShapeNamePathCaseTables();

const ACTIVE_WORKBOOK_AVAILABLE_SNAPSHOT = {
  identity: {
    fullName: "c:/fixtures/FIXTURES.xlsm",
    isAddin: false,
    name: "fixtures.xlsm",
    path: "c:/fixtures"
  },
  observedAt: "2026-03-21T00:00:00.000Z",
  providerKind: "excel-active-workbook",
  state: "available",
  version: 1
} as const;
const ACTIVE_WORKBOOK_MISMATCHED_SNAPSHOT = {
  identity: {
    fullName: "c:/fixtures/OTHER-BOOK.xlsm",
    isAddin: false,
    name: "other-book.xlsm",
    path: "c:/fixtures"
  },
  observedAt: "2026-03-21T00:01:00.000Z",
  providerKind: "excel-active-workbook",
  state: "available",
  version: 1
} as const;
const ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT = {
  observedAt: "2026-03-21T00:01:00.000Z",
  providerKind: "excel-active-workbook",
  reason: "no-active-workbook",
  state: "unavailable",
  version: 1
} as const;
const NEGATIVE_LANGUAGE_FEATURE_RETRY_COUNT = 2;
const NEGATIVE_LANGUAGE_FEATURE_RETRY_DELAY_MS = 100;

export async function run(): Promise<void> {
  const extension = vscode.extensions.getExtension("tagi0.vba-extension");
  assert.ok(extension, "extension must be discoverable");

  await extension.activate();
  await vscode.workspace.getConfiguration("editor").update("snippetSuggestions", "top", vscode.ConfigurationTarget.Global);
  await vscode.workspace.getConfiguration("editor").update("insertSpaces", true, vscode.ConfigurationTarget.Global);
  await vscode.workspace.getConfiguration("editor").update("tabSize", 4, vscode.ConfigurationTarget.Global);

  const extensionRoot = path.resolve(__dirname, "..", "..", "..");
  const fixturesPath = path.resolve(extensionRoot, "test", "fixtures");
  const sampleDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "sample.bas"));
  await vscode.window.showTextDocument(sampleDocument);

  assert.equal(sampleDocument.languageId, "vba");

  const symbols = await waitForSymbols(sampleDocument);
  assert.ok(symbols.length > 0, "document symbols should be available");

  const libraryDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "PublicApi.bas"));
  await vscode.window.showTextDocument(libraryDocument);
  const numberDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "NumberApi.bas"));
  await vscode.window.showTextDocument(numberDocument);
  const formatterDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "FormatterApi.bas"));
  await vscode.window.showTextDocument(formatterDocument);

  const consumerDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "Consumer.bas"));
  await vscode.window.showTextDocument(consumerDocument);

  const completionItems = await waitForCompletions(
    consumerDocument,
    new vscode.Position(5, 4),
    (items) => items.some((item) => item.label === "PublicMessage")
  );
  const publicMessageCompletion = completionItems.find((item) => item.label === "PublicMessage");

  assert.ok(
    completionItems.some((item) => item.label === "PublicMessage"),
    "cross-file completion should include exported workspace symbols"
  );
  assert.ok(publicMessageCompletion?.detail?.includes("String"), "completion detail should include inferred type information");

  const builtInCompletionDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "BuiltInCompletion.bas"));
  await vscode.window.showTextDocument(builtInCompletionDocument);

  const applicationCompletionItems = await waitForCompletions(
    builtInCompletionDocument,
    new vscode.Position(4, 7),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Application")
  );
  const excelConstantCompletionItems = await waitForCompletions(
    builtInCompletionDocument,
    new vscode.Position(5, 7),
    (items) => items.some((item) => getCompletionItemLabel(item) === "xlAll")
  );
  const applicationCompletion = applicationCompletionItems.find((item) => getCompletionItemLabel(item) === "Application");
  const excelConstantCompletion = excelConstantCompletionItems.find((item) => getCompletionItemLabel(item) === "xlAll");

  assert.ok(applicationCompletion, "built-in completion should include Application");
  assert.ok(applicationCompletion.detail?.includes("Excel"), "built-in completion should include source detail");
  assert.ok(excelConstantCompletion, "built-in completion should include Excel constants");
  assert.ok(excelConstantCompletion.detail?.includes("Excel"), "built-in constant should include source detail");

  const thisWorkbookDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "ThisWorkbook.cls"));
  await vscode.window.showTextDocument(thisWorkbookDocument);
  const sheet1Document = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "Sheet1.cls"));
  await vscode.window.showTextDocument(sheet1Document);
  const chart1Document = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "Chart1.cls"));
  await vscode.window.showTextDocument(chart1Document);

  const builtInMemberCompletionDocument = await vscode.workspace.openTextDocument(
    path.resolve(fixturesPath, "BuiltInMemberCompletion.bas")
  );
  await vscode.window.showTextDocument(builtInMemberCompletionDocument);
  const thisWorkbookDefinitions = await waitForDefinitions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "ThisWorkbook", -1),
    (locations) => locations.some((location) => location.uri.toString() === thisWorkbookDocument.uri.toString())
  );
  const sheet1Definitions = await waitForDefinitions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "Sheet1", -1),
    (locations) => locations.some((location) => location.uri.toString() === sheet1Document.uri.toString())
  );
  const chart1Definitions = await waitForDefinitions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "Chart1", -1),
    (locations) => locations.some((location) => location.uri.toString() === chart1Document.uri.toString())
  );

  const applicationMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "Application."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "WorksheetFunction")
  );
  const worksheetFunctionMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "WorksheetFunction.Su"),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Sum")
  );
  const chainedWorksheetFunctionMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "Application.WorksheetFunction.Su"),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Sum")
  );
  const activeWorkbookMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "ActiveWorkbook."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "SaveAs")
  );
  const thisWorkbookMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "ThisWorkbook."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "SaveAs")
  );
  const sheet1MemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "Sheet1."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Evaluate")
  );
  const chart1MemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "Chart1."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "ChartArea")
  );
  const workbookWorksheetsMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "ActiveWorkbook.Worksheets."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const indexedWorksheetMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "Worksheets(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Evaluate")
  );
  const indexedWorksheetStringMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, 'Worksheets("A(1)").'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Evaluate")
  );
  const indexedWorksheetExpressionMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "Worksheets(i + 1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "SaveAs")
  );
  const groupedWorksheetMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, 'Worksheets(Array("Sheet1", "Sheet2")).'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const chainedIndexedWorksheetMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "ActiveWorkbook.Worksheets(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "ExportAsFixedFormat")
  );
  const chainedIndexedWorksheetFunctionMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "ActiveWorkbook.Worksheets(GetIndex())."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const applicationActiveCellMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    findPositionAfterToken(builtInMemberCompletionDocument, "Application.ActiveCell."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Address")
  );
  const worksheetFunctionPropertyCompletion = applicationMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "WorksheetFunction"
  );
  const worksheetFunctionSumCompletion = worksheetFunctionMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Sum"
  );
  const activeWorkbookSaveAsCompletion = activeWorkbookMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SaveAs"
  );
  const activeWorkbookWorksheetsCompletion = activeWorkbookMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Worksheets"
  );
  const thisWorkbookSaveAsCompletion = thisWorkbookMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SaveAs"
  );
  const sheet1EvaluateCompletion = sheet1MemberCompletionItems.find((item) => getCompletionItemLabel(item) === "Evaluate");
  const sheet1SaveAsCompletion = sheet1MemberCompletionItems.find((item) => getCompletionItemLabel(item) === "SaveAs");
  const chart1ChartAreaCompletion = chart1MemberCompletionItems.find((item) => getCompletionItemLabel(item) === "ChartArea");
  const chart1SetSourceDataCompletion = chart1MemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SetSourceData"
  );
  const workbookWorksheetsCountCompletion = workbookWorksheetsMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const indexedWorksheetEvaluateCompletion = indexedWorksheetMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Evaluate"
  );
  const indexedWorksheetSaveAsCompletion = indexedWorksheetMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SaveAs"
  );
  const indexedWorksheetStringEvaluateCompletion = indexedWorksheetStringMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Evaluate"
  );
  const indexedWorksheetExpressionSaveAsCompletion = indexedWorksheetExpressionMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SaveAs"
  );
  const groupedWorksheetCountCompletion = groupedWorksheetMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const chainedIndexedWorksheetExportCompletion = chainedIndexedWorksheetMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "ExportAsFixedFormat"
  );
  const chainedIndexedWorksheetFunctionCountCompletion = chainedIndexedWorksheetFunctionMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const applicationActiveCellAddressCompletion = applicationActiveCellMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Address"
  );

  assert.ok(worksheetFunctionPropertyCompletion, "built-in member completion should include Application.WorksheetFunction");
  assert.ok(
    worksheetFunctionPropertyCompletion.detail?.includes("Excel Application property"),
    "built-in member completion should include owner detail"
  );
  assert.ok(
    worksheetFunctionSumCompletion?.detail?.includes("Excel WorksheetFunction method"),
    "built-in member completion should include method detail"
  );
  assert.ok(
    chainedWorksheetFunctionMemberCompletionItems.some((item) => getCompletionItemLabel(item) === "Sum"),
    "built-in chained member completion should resolve through known member types"
  );
  assert.ok(activeWorkbookSaveAsCompletion?.detail?.includes("Excel Workbook method"));
  assert.ok(activeWorkbookWorksheetsCompletion?.detail?.includes("Excel Workbook property"));
  assert.ok(
    thisWorkbookDefinitions.some((location) => location.uri.toString() === thisWorkbookDocument.uri.toString()),
    "ThisWorkbook root should resolve to the workbook document module before alias assertions"
  );
  assert.ok(thisWorkbookSaveAsCompletion?.detail?.includes("Excel Workbook method"));
  assert.ok(
    sheet1Definitions.some((location) => location.uri.toString() === sheet1Document.uri.toString()),
    "Sheet1 root should resolve to the worksheet document module before alias assertions"
  );
  assert.ok(sheet1EvaluateCompletion?.detail?.includes("Excel Worksheet method"));
  assert.ok(sheet1SaveAsCompletion?.detail?.includes("Excel Worksheet method"));
  assert.ok(
    chart1Definitions.some((location) => location.uri.toString() === chart1Document.uri.toString()),
    "Chart1 root should resolve to the chart document module before alias assertions"
  );
  assert.ok(chart1ChartAreaCompletion?.detail?.includes("Excel Chart property"));
  assert.ok(chart1SetSourceDataCompletion?.detail?.includes("Excel Chart method"));
  assert.ok(workbookWorksheetsCountCompletion?.detail?.includes("Excel Worksheets property"));
  assert.ok(indexedWorksheetEvaluateCompletion?.detail?.includes("Excel Worksheet method"));
  assert.ok(indexedWorksheetSaveAsCompletion?.detail?.includes("Excel Worksheet method"));
  assert.ok(indexedWorksheetStringEvaluateCompletion?.detail?.includes("Excel Worksheet method"));
  assert.ok(indexedWorksheetExpressionSaveAsCompletion?.detail?.includes("Excel Worksheet method"));
  assert.ok(groupedWorksheetCountCompletion?.detail?.includes("Excel Worksheets property"));
  assert.equal(
    groupedWorksheetMemberCompletionItems.some((item) => getCompletionItemLabel(item) === "Evaluate"),
    false,
    "grouped Worksheets selector should stay on the Worksheets collection"
  );
  assert.ok(chainedIndexedWorksheetExportCompletion?.detail?.includes("Excel Worksheet method"));
  assert.ok(chainedIndexedWorksheetFunctionCountCompletion?.detail?.includes("Excel Worksheets property"));
  assert.equal(
    chainedIndexedWorksheetFunctionMemberCompletionItems.some(
      (item) => getCompletionItemLabel(item) === "ExportAsFixedFormat"
    ),
    false,
    "function-based Worksheets selector should stay on the Worksheets collection"
  );
  assert.ok(applicationActiveCellAddressCompletion?.detail?.includes("Excel Range"));

  const oleObjectBuiltInDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "OleObjectBuiltIn.bas"));
  await vscode.window.showTextDocument(oleObjectBuiltInDocument);

  const sheetOleObjectsCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const indexedOleObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Activate")
  );
  const namedOleObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'Sheet1.OLEObjects("CheckBox1").'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Visible")
  );
  const expressionOleObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects(i + 1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Name")
  );
  const functionOleObjectsCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects(GetIndex())."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const itemIndexedOleObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects.Item(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Activate")
  );
  const itemNamedOleObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'Sheet1.OLEObjects.Item("CheckBox1").'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Visible")
  );
  const itemExpressionOleObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects.Item(i + 1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Name")
  );
  const itemFunctionOleObjectsCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects.Item(GetIndex())."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const chartOleObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Chart1.OLEObjects(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Name")
  );
  const chartItemOleObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Chart1.OLEObjects.Item(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Name")
  );
  const oleObjectObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects(1).Object."),
    (items) => !items.some((item) => getCompletionItemLabel(item) === "Activate")
  );
  const namedOleObjectObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'Sheet1.OLEObjects("CheckBox1").Object.'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const itemNamedOleObjectObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'Sheet1.OLEObjects.Item("CheckBox1").Object.'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Select")
  );
  const chartOleObjectObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'Chart1.OLEObjects("CheckBox1").Object.'),
    (items) => !items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const activeSheetOleObjectObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ActiveSheet.OLEObjects("CheckBox1").Object.'),
    (items) => !items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const thisWorkbookNamedOleObjectObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const thisWorkbookItemNamedOleObjectObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(
      oleObjectBuiltInDocument,
      'ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.'
    ),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Select")
  );
  const thisWorkbookCodeNameOleObjectObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.'),
    (items) => !items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const thisWorkbookIndexedOleObjectObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.'),
    (items) => !items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const activeWorkbookNamedOleObjectObjectCompletionItems = await waitForCompletions(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.'),
    (items) => !items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const activateSignatureHelp = await waitForSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects(1).Activate("),
    (help) => help.signatures.length > 0
  );
  const dynamicActivateSuppressed = await waitForNoSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects(GetIndex()).Activate(")
  );
  const itemActivateSignatureHelp = await waitForSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects.Item(1).Activate("),
    (help) => help.signatures.length > 0
  );
  const namedObjectSelectSignatureHelp = await waitForSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'Sheet1.OLEObjects("CheckBox1").Object.Select('),
    (help) => help.signatures.length > 0
  );
  const itemNamedObjectSelectSignatureHelp = await waitForSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'Sheet1.OLEObjects.Item("CheckBox1").Object.Select('),
    (help) => help.signatures.length > 0
  );
  const thisWorkbookNamedObjectSelectSignatureHelp = await waitForSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select('),
    (help) => help.signatures.length > 0
  );
  const thisWorkbookItemNamedObjectSelectSignatureHelp = await waitForSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(
      oleObjectBuiltInDocument,
      'ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Select('
    ),
    (help) => help.signatures.length > 0
  );
  const itemDynamicActivateSuppressed = await waitForNoSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects.Item(GetIndex()).Activate(")
  );
  const chartObjectSelectSuppressed = await waitForNoSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'Chart1.OLEObjects("CheckBox1").Object.Select(')
  );
  const thisWorkbookIndexedObjectSelectSuppressed = await waitForNoSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.Select(')
  );
  const activeWorkbookNamedObjectSelectSuppressed = await waitForNoSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(')
  );
  const thisWorkbookCodeNameObjectSelectSuppressed = await waitForNoSignatureHelp(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Select(')
  );
  const oleObjectHover = await waitForHover(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects(1).Nam"),
    (hovers) => hovers.length > 0
  );
  const itemOleObjectHover = await waitForHover(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, "Sheet1.OLEObjects.Item(1).Nam"),
    (hovers) => hovers.length > 0
  );
  const namedObjectHover = await waitForHover(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'Sheet1.OLEObjects("CheckBox1").Object.Valu'),
    (hovers) => hovers.length > 0
  );
  const itemNamedObjectHover = await waitForHover(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'Sheet1.OLEObjects.Item("CheckBox1").Object.Valu'),
    (hovers) => hovers.length > 0
  );
  const thisWorkbookNamedObjectHover = await waitForHover(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu'),
    (hovers) => hovers.length > 0
  );
  const thisWorkbookItemNamedObjectHover = await waitForHover(
    oleObjectBuiltInDocument,
    findPositionAfterToken(
      oleObjectBuiltInDocument,
      'ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu'
    ),
    (hovers) => hovers.length > 0
  );
  const thisWorkbookIndexedObjectHoverSuppressed = await waitForNoHover(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.Valu')
  );
  const activeWorkbookNamedObjectHoverSuppressed = await waitForNoHover(
    oleObjectBuiltInDocument,
    findPositionAfterToken(
      oleObjectBuiltInDocument,
      'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu'
    )
  );
  const thisWorkbookCodeNameObjectHoverSuppressed = await waitForNoHover(
    oleObjectBuiltInDocument,
    findPositionAfterToken(oleObjectBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Valu')
  );
  const oleObjectLegend = await waitForSemanticTokensLegend(
    oleObjectBuiltInDocument,
    (legend) => legend.tokenTypes.includes("variable") && legend.tokenTypes.includes("function")
  );
  const oleObjectTokens = await waitForSemanticTokens(
    oleObjectBuiltInDocument,
    (tokens) => tokens.data.length > 0
  );
  const sheetOleObjectsCountCompletion = sheetOleObjectsCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const indexedOleObjectActivateCompletion = indexedOleObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Activate"
  );
  const indexedOleObjectNameCompletion = indexedOleObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Name"
  );
  const namedOleObjectVisibleCompletion = namedOleObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Visible"
  );
  const expressionOleObjectNameCompletion = expressionOleObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Name"
  );
  const functionOleObjectsCountCompletion = functionOleObjectsCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const itemIndexedOleObjectActivateCompletion = itemIndexedOleObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Activate"
  );
  const itemNamedOleObjectVisibleCompletion = itemNamedOleObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Visible"
  );
  const itemExpressionOleObjectNameCompletion = itemExpressionOleObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Name"
  );
  const itemFunctionOleObjectsCountCompletion = itemFunctionOleObjectsCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const chartOleObjectNameCompletion = chartOleObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Name"
  );
  const chartItemOleObjectNameCompletion = chartItemOleObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Name"
  );
  const namedOleObjectObjectValueCompletion = namedOleObjectObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const namedOleObjectObjectSelectCompletion = namedOleObjectObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Select"
  );
  const itemNamedOleObjectObjectValueCompletion = itemNamedOleObjectObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const thisWorkbookNamedOleObjectObjectValueCompletion = thisWorkbookNamedOleObjectObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const thisWorkbookNamedOleObjectObjectSelectCompletion = thisWorkbookNamedOleObjectObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Select"
  );
  const thisWorkbookItemNamedOleObjectObjectValueCompletion = thisWorkbookItemNamedOleObjectObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const thisWorkbookItemNamedOleObjectObjectSelectCompletion = thisWorkbookItemNamedOleObjectObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Select"
  );
  const oleObjectHoverText = getHoverContentsText(oleObjectHover[0]);
  const itemOleObjectHoverText = getHoverContentsText(itemOleObjectHover[0]);
  const namedObjectHoverText = getHoverContentsText(namedObjectHover[0]);
  const itemNamedObjectHoverText = getHoverContentsText(itemNamedObjectHover[0]);
  const thisWorkbookNamedObjectHoverText = getHoverContentsText(thisWorkbookNamedObjectHover[0]);
  const thisWorkbookItemNamedObjectHoverText = getHoverContentsText(thisWorkbookItemNamedObjectHover[0]);
  const decodedOleObjectTokens = decodeSemanticTokens(oleObjectTokens, oleObjectLegend);

  assert.ok(sheetOleObjectsCountCompletion?.detail?.includes("Excel OLEObjects property"));
  assert.ok(indexedOleObjectActivateCompletion?.detail?.includes("Excel OLEObject method"));
  assert.ok(indexedOleObjectNameCompletion?.detail?.includes("Excel OLEObject property"));
  assert.ok(namedOleObjectVisibleCompletion?.detail?.includes("Excel OLEObject property"));
  assert.ok(expressionOleObjectNameCompletion?.detail?.includes("Excel OLEObject property"));
  assert.ok(functionOleObjectsCountCompletion?.detail?.includes("Excel OLEObjects property"));
  assert.ok(itemIndexedOleObjectActivateCompletion?.detail?.includes("Excel OLEObject method"));
  assert.ok(itemNamedOleObjectVisibleCompletion?.detail?.includes("Excel OLEObject property"));
  assert.ok(itemExpressionOleObjectNameCompletion?.detail?.includes("Excel OLEObject property"));
  assert.ok(itemFunctionOleObjectsCountCompletion?.detail?.includes("Excel OLEObjects property"));
  assert.equal(
    functionOleObjectsCompletionItems.some((item) => getCompletionItemLabel(item) === "Activate"),
    false,
    "function-based OLEObjects selector should stay on the OLEObjects collection"
  );
  assert.equal(
    itemFunctionOleObjectsCompletionItems.some((item) => getCompletionItemLabel(item) === "Activate"),
    false,
    "function-based OLEObjects.Item selector should stay on the OLEObjects collection"
  );
  assert.equal(
    oleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Activate"),
    false,
    "OLEObject.Object の先では OLEObject method を出さない"
  );
  assert.ok(namedOleObjectObjectValueCompletion?.detail?.includes("CheckBox property"));
  assert.ok(namedOleObjectObjectSelectCompletion?.detail?.includes("CheckBox method"));
  assert.equal(
    namedOleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Activate"),
    false,
    "named worksheet selector の OLEObject.Object は control owner へ解決し、OLEObject method を出さない"
  );
  assert.ok(itemNamedOleObjectObjectValueCompletion?.detail?.includes("CheckBox property"));
  assert.equal(
    itemNamedOleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Activate"),
    false,
    "named worksheet Item selector の OLEObject.Object は control owner へ解決し、OLEObject method を出さない"
  );
  assert.ok(thisWorkbookNamedOleObjectObjectValueCompletion?.detail?.includes("CheckBox property"));
  assert.ok(thisWorkbookNamedOleObjectObjectSelectCompletion?.detail?.includes("CheckBox method"));
  assert.equal(
    thisWorkbookNamedOleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Activate"),
    false,
    "ThisWorkbook.Worksheets(\"Sheet One\") の OLEObject.Object も control owner へ解決する"
  );
  assert.ok(thisWorkbookItemNamedOleObjectObjectValueCompletion?.detail?.includes("CheckBox property"));
  assert.ok(thisWorkbookItemNamedOleObjectObjectSelectCompletion?.detail?.includes("CheckBox method"));
  assert.equal(
    thisWorkbookItemNamedOleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Activate"),
    false,
    "ThisWorkbook.Worksheets(\"Sheet One\").OLEObjects.Item(\"CheckBox1\") も control owner へ解決する"
  );
  assert.equal(
    thisWorkbookCodeNameOleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Value"),
    false,
    "ThisWorkbook.Worksheets(\"Sheet1\") は codeName 指定なので control owner に昇格しない"
  );
  assert.equal(
    thisWorkbookIndexedOleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Value"),
    false,
    "ThisWorkbook.Worksheets(1) は index 指定なので control owner に昇格しない"
  );
  assert.equal(
    activeWorkbookNamedOleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Value"),
    false,
    "ActiveWorkbook.Worksheets(\"Sheet One\") は broad root として未解決を維持する"
  );
  assert.equal(
    chartOleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Value"),
    false,
    "unsupported chartsheet owner は OLEObject.Object を未解決のまま維持する"
  );
  assert.equal(
    activeSheetOleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Value"),
    false,
    "ActiveSheet root は worksheet sidecar owner 解決へ昇格しない"
  );
  assert.ok(chartOleObjectNameCompletion?.detail?.includes("Excel OLEObject property"));
  assert.ok(chartItemOleObjectNameCompletion?.detail?.includes("Excel OLEObject property"));
  assert.equal(activateSignatureHelp.signatures[0]?.label.includes("Activate()"), true);
  assert.equal(dynamicActivateSuppressed, true);
  assert.equal(itemActivateSignatureHelp.signatures[0]?.label.includes("Activate()"), true);
  assert.equal(itemDynamicActivateSuppressed, true);
  assert.equal(namedObjectSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(itemNamedObjectSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(chartObjectSelectSuppressed, true);
  assert.equal(oleObjectHoverText.includes("OLEObject.Name"), true);
  assert.equal(oleObjectHoverText.includes("excel.oleobject.name"), true);
  assert.equal(itemOleObjectHoverText.includes("OLEObject.Name"), true);
  assert.equal(itemOleObjectHoverText.includes("excel.oleobject.name"), true);
  assert.equal(namedObjectHoverText.includes("CheckBox.Value"), true);
  assert.equal(namedObjectHoverText.includes("microsoft.office.interop.excel.checkbox.value"), true);
  assert.equal(itemNamedObjectHoverText.includes("CheckBox.Value"), true);
  assert.equal(thisWorkbookNamedObjectHoverText.includes("CheckBox.Value"), true);
  assert.equal(thisWorkbookItemNamedObjectHoverText.includes("CheckBox.Value"), true);
  assertDecodedSemanticToken(oleObjectBuiltInDocument.getText(), decodedOleObjectTokens, 21, "Select", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(oleObjectBuiltInDocument.getText(), decodedOleObjectTokens, 42, "Value", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(oleObjectBuiltInDocument.getText(), decodedOleObjectTokens, 44, "Select", {
    modifiers: [],
    type: "function"
  });
  assertNoDecodedSemanticToken(oleObjectBuiltInDocument.getText(), decodedOleObjectTokens, 23, "Select");
  assert.equal(thisWorkbookNamedObjectSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(thisWorkbookItemNamedObjectSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(thisWorkbookIndexedObjectHoverSuppressed, true);
  assert.equal(activeWorkbookNamedObjectHoverSuppressed, true);
  assert.equal(thisWorkbookCodeNameObjectHoverSuppressed, true);
  assert.equal(thisWorkbookIndexedObjectSelectSuppressed, true);
  assert.equal(activeWorkbookNamedObjectSelectSuppressed, true);
  assert.equal(thisWorkbookCodeNameObjectSelectSuppressed, true);

  const worksheetControlShapeNamePathOlePositiveEntries = getWorksheetControlShapeNamePathCompletionEntries("positive", {
    fixture: "packages/extension/test/fixtures/OleObjectBuiltIn.bas",
    scope: "extension",
    routeKind: "ole-object"
  });
  const worksheetControlShapeNamePathOleAlwaysAvailablePositiveEntries =
    worksheetControlShapeNamePathOlePositiveEntries.filter((entry) => entry.rootKind !== "workbook-qualified-matched");
  const worksheetControlShapeNamePathOleNegativeEntries = getWorksheetControlShapeNamePathCompletionEntries("negative", {
    fixture: "packages/extension/test/fixtures/OleObjectBuiltIn.bas",
    scope: "extension",
    routeKind: "ole-object"
  });
  const worksheetControlShapeNamePathOleClosedEntries = worksheetControlShapeNamePathOleNegativeEntries.filter(
    (entry) => entry.rootKind === "workbook-qualified-closed"
  );
  const worksheetControlShapeNamePathOleReasonEntries = worksheetControlShapeNamePathOleNegativeEntries.filter(
    (entry) => entry.rootKind !== "workbook-qualified-closed"
  );

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT);
  await assertWorkbookRootCompletionCases(
    oleObjectBuiltInDocument,
    mapExtensionWorksheetControlShapeNamePathPositiveCompletionCases(
      worksheetControlShapeNamePathOleAlwaysAvailablePositiveEntries,
      (entry) => `${entry.anchor} は ${entry.rootKind} root なので snapshot なしでも control owner へ解決する`
    )
  );
  await assertWorkbookRootClosedCompletionCases(
    oleObjectBuiltInDocument,
    mapExtensionWorksheetControlShapeNamePathNoCompletionCases(
      [...worksheetControlShapeNamePathOleReasonEntries, ...worksheetControlShapeNamePathOleClosedEntries],
      (entry) =>
        entry.rootKind === "workbook-qualified-closed"
          ? `${entry.anchor} は active workbook が閉じている間は control owner に昇格しない`
          : `${entry.anchor} は ${entry.reason} のため control owner に昇格しない`
    )
  );

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_AVAILABLE_SNAPSHOT);
  try {
    await assertWorkbookRootCompletionCases(
      oleObjectBuiltInDocument,
      mapExtensionWorksheetControlShapeNamePathPositiveCompletionCases(
        worksheetControlShapeNamePathOlePositiveEntries,
        (entry) =>
          entry.rootKind === "workbook-qualified-matched"
            ? `${entry.anchor} は active workbook match 時に control owner へ解決する`
            : `${entry.anchor} は ${entry.rootKind} root として control owner へ解決する`
      )
    );
    await assertWorkbookRootClosedCompletionCases(
      oleObjectBuiltInDocument,
      mapExtensionWorksheetControlShapeNamePathNoCompletionCases(
        worksheetControlShapeNamePathOleReasonEntries,
        (entry) => `${entry.anchor} は ${entry.reason} のため match 中でも control owner に昇格しない`
      )
    );
  } finally {
    await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT);
  }

  await runExtensionWorksheetControlShapeNamePathInteractionSharedCases({
    document: oleObjectBuiltInDocument,
    fixture: "packages/extension/test/fixtures/OleObjectBuiltIn.bas",
    routeKind: "ole-object"
  });

  const workbookQualifiedRootItemOleStaticCompletionChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.',
      'CheckBox property',
      "Activate",
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object は control owner へ解決する'
    ],
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.',
      'CheckBox property',
      "Activate",
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object は control owner へ解決する'
    ]
  ] as const;
  const workbookQualifiedRootItemOleStaticHoverChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Valu',
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1") の hover は control owner へ解決する'
    ],
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu',
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1") の hover は control owner へ解決する'
    ]
  ] as const;
  const workbookQualifiedRootItemOleStaticSignatureChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(',
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1") の signature help は control owner へ解決する'
    ],
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1") の signature help は control owner へ解決する'
    ]
  ] as const;
  const workbookQualifiedRootItemOleNonTargetHoverChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Valu',
      'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'
    ],
    [
      'ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu',
      'ThisWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Valu',
      'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu',
      'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'
    ]
  ] as const;
  const workbookQualifiedRootItemOleNonTargetCompletionChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.',
      'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'
    ],
    [
      'ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.',
      'ThisWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.',
      'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.',
      'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'
    ]
  ] as const;
  const workbookQualifiedRootItemOleNonTargetSignatureChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Select(',
      'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので signature help を出さない'
    ],
    [
      'ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(',
      'ThisWorkbook.Worksheets.Item(1) は numeric selector なので signature help を出さない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Select(',
      'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので signature help を出さない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(',
      'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので signature help を出さない'
    ]
  ] as const;
  const workbookQualifiedRootItemOleClosedCompletionChecks = [
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.',
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.',
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ]
  ] as const;
  const workbookQualifiedRootItemOleClosedHoverChecks = [
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Valu',
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu',
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ]
  ] as const;
  const workbookQualifiedRootItemOleClosedSignatureChecks = [
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(',
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ]
  ] as const;

  for (const [token, detailFragment, blockedLabel, message] of workbookQualifiedRootItemOleStaticCompletionChecks) {
    const items = await waitForCompletionLabelStateAtToken(oleObjectBuiltInDocument, token, "Value", true);
    const completion = items.find((item) => getCompletionItemLabel(item) === "Value");

    assert.ok(completion?.detail?.includes(detailFragment), message);
    assert.equal(hasCompletionItemLabel(items, blockedLabel), false, message);
  }
  for (const [token, message] of workbookQualifiedRootItemOleStaticHoverChecks) {
    const hovers = await waitForHoverAtToken(oleObjectBuiltInDocument, token, (items) => items.length > 0);
    assert.equal(getHoverContentsText(hovers[0]).includes("CheckBox.Value"), true, message);
  }
  for (const [token, message] of workbookQualifiedRootItemOleStaticSignatureChecks) {
    const signatureHelp = await waitForSignatureHelpAtToken(
      oleObjectBuiltInDocument,
      token,
      (help) => help.signatures.length > 0
    );
    assert.equal(signatureHelp.signatures[0]?.label, "Select(Replace) As Object", message);
  }
  for (const [token, message] of workbookQualifiedRootItemOleNonTargetCompletionChecks) {
    const items = await waitForCompletionLabelStateAtToken(oleObjectBuiltInDocument, token, "Value", false);
    assert.equal(hasCompletionItemLabel(items, "Value"), false, message);
  }
  for (const [token, message] of workbookQualifiedRootItemOleNonTargetHoverChecks) {
    assert.equal(await waitForNoHoverAtToken(oleObjectBuiltInDocument, token), true, message);
  }
  for (const [token, message] of workbookQualifiedRootItemOleNonTargetSignatureChecks) {
    assert.equal(await waitForNoSignatureHelpAtToken(oleObjectBuiltInDocument, token), true, message);
  }
  for (const [token, message] of workbookQualifiedRootItemOleClosedCompletionChecks) {
    const items = await waitForCompletionLabelStateAtToken(oleObjectBuiltInDocument, token, "Value", false);
    assert.equal(hasCompletionItemLabel(items, "Value"), false, message);
  }
  for (const [token, message] of workbookQualifiedRootItemOleClosedHoverChecks) {
    assert.equal(await waitForNoHoverAtToken(oleObjectBuiltInDocument, token), true, message);
  }
  for (const [token, message] of workbookQualifiedRootItemOleClosedSignatureChecks) {
    assert.equal(await waitForNoSignatureHelpAtToken(oleObjectBuiltInDocument, token), true, message);
  }
  assertWorkbookRootSemanticCases(oleObjectBuiltInDocument.getText(), decodedOleObjectTokens, [
    [
      'Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Value',
      "Value",
      { modifiers: [], type: "variable" },
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object は semantic token を出す'
    ],
    [
      'Call ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(',
      "Select",
      { modifiers: [], type: "function" },
      'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object は semantic token を出す'
    ]
  ] as const);
  assertWorkbookRootNoSemanticCases(oleObjectBuiltInDocument.getText(), decodedOleObjectTokens, [
    [
      'Debug.Print ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Value',
      "Value",
      'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので semantic token を出さない'
    ],
    [
      'Debug.Print ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value',
      "Value",
      'ThisWorkbook.Worksheets.Item(1) は numeric selector なので semantic token を出さない'
    ],
    [
      'Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Value',
      "Value",
      'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1") は snapshot 未一致の間は semantic token を出さない'
    ]
  ] as const);

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_AVAILABLE_SNAPSHOT);
  try {
    const activeWorkbookBoundOleObjectObjectCompletionItems = await waitForCompletions(
      oleObjectBuiltInDocument,
      findPositionAfterToken(oleObjectBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.'),
      (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
    );
    const activeWorkbookBoundOleObjectHover = await waitForHover(
      oleObjectBuiltInDocument,
      findPositionAfterToken(oleObjectBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu'),
      (hovers) => hovers.length > 0
    );
    const activeWorkbookBoundOleObjectSelectSignatureHelp = await waitForSignatureHelp(
      oleObjectBuiltInDocument,
      findPositionAfterToken(oleObjectBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select('),
      (help) => help.signatures.length > 0
    );
    const activeWorkbookBoundOleObjectValueCompletion = activeWorkbookBoundOleObjectObjectCompletionItems.find(
      (item) => getCompletionItemLabel(item) === "Value"
    );
    const activeWorkbookBoundOleObjectSelectCompletion = activeWorkbookBoundOleObjectObjectCompletionItems.find(
      (item) => getCompletionItemLabel(item) === "Select"
    );
    const activeWorkbookBoundOleObjectHoverText = getHoverContentsText(activeWorkbookBoundOleObjectHover[0]);

    assert.ok(activeWorkbookBoundOleObjectValueCompletion?.detail?.includes("CheckBox property"));
    assert.ok(activeWorkbookBoundOleObjectSelectCompletion?.detail?.includes("CheckBox method"));
    assert.equal(
      activeWorkbookBoundOleObjectObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Activate"),
      false,
      "match 済み active workbook の OLEObject.Object は control owner へ解決し、OLEObject method を出さない"
    );
    assert.equal(activeWorkbookBoundOleObjectHoverText.includes("CheckBox.Value"), true);
    assert.equal(activeWorkbookBoundOleObjectSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");

    const workbookQualifiedRootItemOleMatchedCompletionChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.',
        'CheckBox property',
        "Activate",
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object は control owner へ解決する'
      ],
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.',
        'CheckBox property',
        "Activate",
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object は control owner へ解決する'
      ]
    ] as const;
    const workbookQualifiedRootItemOleMatchedHoverChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Valu',
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1") の hover は control owner へ解決する'
      ],
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu',
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1") の hover は control owner へ解決する'
      ]
    ] as const;
    const workbookQualifiedRootItemOleMatchedNonTargetCompletionChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.',
        'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'
      ],
      [
        'ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.',
        'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'
      ]
    ] as const;
    const workbookQualifiedRootItemOleMatchedNonTargetHoverChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Valu',
        'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので hover を出さない'
      ],
      [
        'ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu',
        'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので hover を出さない'
      ]
    ] as const;
    const workbookQualifiedRootItemOleMatchedNonTargetSignatureChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Select(',
        'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので signature help を出さない'
      ],
      [
        'ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(',
        'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので signature help を出さない'
      ]
    ] as const;
    const workbookQualifiedRootItemOleMatchedSignatureChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(',
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1") の signature help は control owner へ解決する'
      ],
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
        'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1") の signature help は control owner へ解決する'
      ]
    ] as const;

    for (const [token, detailFragment, blockedLabel, message] of workbookQualifiedRootItemOleMatchedCompletionChecks) {
      const items = await waitForCompletionLabelStateAtToken(oleObjectBuiltInDocument, token, "Value", true);
      const completion = items.find((item) => getCompletionItemLabel(item) === "Value");

      assert.ok(completion?.detail?.includes(detailFragment), message);
      assert.equal(hasCompletionItemLabel(items, blockedLabel), false, message);
    }
    for (const [token, message] of workbookQualifiedRootItemOleMatchedHoverChecks) {
      const hovers = await waitForHoverAtToken(oleObjectBuiltInDocument, token, (items) => items.length > 0);
      assert.equal(getHoverContentsText(hovers[0]).includes("CheckBox.Value"), true, message);
    }
    for (const [token, message] of workbookQualifiedRootItemOleMatchedNonTargetCompletionChecks) {
      const items = await waitForCompletionLabelStateAtToken(oleObjectBuiltInDocument, token, "Value", false);
      assert.equal(hasCompletionItemLabel(items, "Value"), false, message);
    }
    for (const [token, message] of workbookQualifiedRootItemOleMatchedNonTargetHoverChecks) {
      assert.equal(await waitForNoHoverAtToken(oleObjectBuiltInDocument, token), true, message);
    }
    for (const [token, message] of workbookQualifiedRootItemOleMatchedNonTargetSignatureChecks) {
      assert.equal(await waitForNoSignatureHelpAtToken(oleObjectBuiltInDocument, token), true, message);
    }
    for (const [token, message] of workbookQualifiedRootItemOleMatchedSignatureChecks) {
      const signatureHelp = await waitForSignatureHelpAtToken(
        oleObjectBuiltInDocument,
        token,
        (help) => help.signatures.length > 0
      );
      assert.equal(signatureHelp.signatures[0]?.label, "Select(Replace) As Object", message);
    }
  } finally {
    await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT);
  }

  const shapesBuiltInDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "ShapesBuiltIn.bas"));
  await vscode.window.showTextDocument(shapesBuiltInDocument);

  await runExtensionWorksheetControlShapeNamePathInteractionSharedCases({
    document: shapesBuiltInDocument,
    fixture: "packages/extension/test/fixtures/ShapesBuiltIn.bas",
    routeKind: "shape-oleformat"
  });

  const shapesCollectionCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Sheet1.Shapes."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const indexedShapeCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Sheet1.Shapes(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Name")
  );
  const functionShapeCollectionCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Sheet1.Shapes(GetIndex())."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const itemShapeCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Sheet1.Shapes.Item(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Name")
  );
  const itemFunctionShapeCollectionCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Sheet1.Shapes.Item(GetIndex())."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const chartShapeCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Chart1.Shapes(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Name")
  );
  const chartItemShapeCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Chart1.Shapes.Item(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Name")
  );
  const oleFormatCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes("CheckBox1").OLEFormat.'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "progID")
  );
  const shapeObjectCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes("CheckBox1").OLEFormat.Object.'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const itemShapeObjectCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Select")
  );
  const shapeNameHover = await waitForHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes("CheckBox1").Nam'),
    (hovers) => hovers.length > 0
  );
  const shapeObjectValueHover = await waitForHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes("CheckBox1").OLEFormat.Object.Valu'),
    (hovers) => hovers.length > 0
  );
  const itemShapeObjectValueHover = await waitForHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.Valu'),
    (hovers) => hovers.length > 0
  );
  const shapeObjectSelectSignatureHelp = await waitForSignatureHelp(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes("CheckBox1").OLEFormat.Object.Select('),
    (help) => help.signatures.length > 0
  );
  const itemShapeObjectSelectSignatureHelp = await waitForSignatureHelp(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.Select('),
    (help) => help.signatures.length > 0
  );
  const indexedShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Sheet1.Shapes(1).OLEFormat.Object.Valu")
  );
  const itemIndexedShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Sheet1.Shapes.Item(1).OLEFormat.Object.Valu")
  );
  const chartShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Chart1.Shapes("CheckBox1").OLEFormat.Object.Valu')
  );
  const chartItemShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Chart1.Shapes.Item("CheckBox1").OLEFormat.Object.Valu')
  );
  const groupedShapeRangeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes.Range(Array("CheckBox1")).OLEFormat.Object.Valu')
  );
  const unmatchedShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes("PlainShape").OLEFormat.Object.Valu')
  );
  const itemUnmatchedShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes.Item("PlainShape").OLEFormat.Object.Valu')
  );
  const worksheetNameRootShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu')
  );
  const worksheetNameRootItemShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Worksheets("Sheet1").Shapes.Item("CheckBox1").OLEFormat.Object.Valu')
  );
  const thisWorkbookWorksheetShapeObjectCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const thisWorkbookWorksheetItemShapeObjectCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Select")
  );
  const thisWorkbookWorksheetCodeNameShapeObjectCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.'),
    (items) => !items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const thisWorkbookWorksheetIndexedShapeObjectCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.'),
    (items) => !items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const activeWorkbookWorksheetShapeObjectCompletionItems = await waitForCompletions(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.'),
    (items) => !items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const thisWorkbookWorksheetShapeObjectValueHover = await waitForHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu'),
    (hovers) => hovers.length > 0
  );
  const thisWorkbookWorksheetItemShapeObjectValueHover = await waitForHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Valu'),
    (hovers) => hovers.length > 0
  );
  const thisWorkbookWorksheetShapeObjectSelectSignatureHelp = await waitForSignatureHelp(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select('),
    (help) => help.signatures.length > 0
  );
  const thisWorkbookWorksheetItemShapeObjectSelectSignatureHelp = await waitForSignatureHelp(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select('),
    (help) => help.signatures.length > 0
  );
  const thisWorkbookWorksheetIndexedShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.Valu')
  );
  const activeWorkbookWorksheetShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu')
  );
  const thisWorkbookWorksheetCodeNameShapeObjectValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu')
  );
  const thisWorkbookWorksheetIndexedShapeObjectSelectSuppressed = await waitForNoSignatureHelp(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.Select(')
  );
  const activeWorkbookWorksheetShapeObjectSelectSuppressed = await waitForNoSignatureHelp(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(')
  );
  const thisWorkbookWorksheetCodeNameShapeObjectSelectSuppressed = await waitForNoSignatureHelp(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(')
  );
  const indexedObjectCallValueHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, 'Sheet1.Shapes("CheckBox1").OLEFormat.Object(1).Valu')
  );
  const functionShapeNameHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Sheet1.Shapes(GetIndex()).Nam")
  );
  const itemFunctionShapeNameHoverSuppressed = await waitForNoHover(
    shapesBuiltInDocument,
    findPositionAfterToken(shapesBuiltInDocument, "Sheet1.Shapes.Item(GetIndex()).Nam")
  );
  const shapesLegend = await waitForSemanticTokensLegend(
    shapesBuiltInDocument,
    (legend) => legend.tokenTypes.includes("variable") && legend.tokenTypes.includes("function")
  );
  const shapesTokens = await waitForSemanticTokens(shapesBuiltInDocument, (tokens) => tokens.data.length > 0);
  const indexedShapeNameCompletion = indexedShapeCompletionItems.find((item) => getCompletionItemLabel(item) === "Name");
  const itemShapeNameCompletion = itemShapeCompletionItems.find((item) => getCompletionItemLabel(item) === "Name");
  const chartShapeNameCompletion = chartShapeCompletionItems.find((item) => getCompletionItemLabel(item) === "Name");
  const chartItemShapeNameCompletion = chartItemShapeCompletionItems.find((item) => getCompletionItemLabel(item) === "Name");
  const oleFormatProgIdCompletion = oleFormatCompletionItems.find((item) => getCompletionItemLabel(item) === "progID");
  const shapeObjectValueCompletion = shapeObjectCompletionItems.find((item) => getCompletionItemLabel(item) === "Value");
  const shapeObjectSelectCompletion = shapeObjectCompletionItems.find((item) => getCompletionItemLabel(item) === "Select");
  const itemShapeObjectValueCompletion = itemShapeObjectCompletionItems.find((item) => getCompletionItemLabel(item) === "Value");
  const itemShapeObjectSelectCompletion = itemShapeObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Select"
  );
  const thisWorkbookWorksheetShapeObjectValueCompletion = thisWorkbookWorksheetShapeObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const thisWorkbookWorksheetShapeObjectSelectCompletion = thisWorkbookWorksheetShapeObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Select"
  );
  const thisWorkbookWorksheetItemShapeObjectValueCompletion = thisWorkbookWorksheetItemShapeObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const thisWorkbookWorksheetItemShapeObjectSelectCompletion = thisWorkbookWorksheetItemShapeObjectCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Select"
  );
  const shapeHoverText = getHoverContentsText(shapeNameHover[0]);
  const shapeObjectHoverText = getHoverContentsText(shapeObjectValueHover[0]);
  const itemShapeObjectHoverText = getHoverContentsText(itemShapeObjectValueHover[0]);
  const thisWorkbookWorksheetShapeObjectHoverText = getHoverContentsText(thisWorkbookWorksheetShapeObjectValueHover[0]);
  const thisWorkbookWorksheetItemShapeObjectHoverText = getHoverContentsText(
    thisWorkbookWorksheetItemShapeObjectValueHover[0]
  );
  const decodedShapesTokens = decodeSemanticTokens(shapesTokens, shapesLegend);

  assert.ok(shapesCollectionCompletionItems.some((item) => getCompletionItemLabel(item) === "Count"));
  assert.equal(
    shapesCollectionCompletionItems.some((item) => getCompletionItemLabel(item) === "Name"),
    false,
    "Shapes collection は indexed access なしでは Shape owner に降りない"
  );
  assert.ok(indexedShapeNameCompletion?.detail?.includes("Excel Shape property"));
  assert.ok(itemShapeNameCompletion?.detail?.includes("Excel Shape property"));
  assert.ok(chartShapeNameCompletion?.detail?.includes("Excel Shape property"));
  assert.ok(chartItemShapeNameCompletion?.detail?.includes("Excel Shape property"));
  assert.equal(
    functionShapeCollectionCompletionItems.some((item) => getCompletionItemLabel(item) === "Name"),
    false,
    "function selector の Shapes は collection のまま維持する"
  );
  assert.equal(
    itemFunctionShapeCollectionCompletionItems.some((item) => getCompletionItemLabel(item) === "Name"),
    false,
    "function selector の Shapes.Item は collection のまま維持する"
  );
  assert.ok(oleFormatProgIdCompletion?.detail?.includes("Excel OLEFormat property"));
  assert.equal(shapeHoverText.includes("Shape.Name"), true);
  assert.equal(shapeHoverText.includes("excel.shape.name"), true);
  assert.ok(shapeObjectValueCompletion?.detail?.includes("CheckBox property"));
  assert.ok(shapeObjectSelectCompletion?.detail?.includes("CheckBox method"));
  assert.equal(
    shapeObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Delete"),
    false,
    "named worksheet selector の Shape.OLEFormat.Object は control owner へ解決し、Shape 専用 method を出さない"
  );
  assert.ok(itemShapeObjectValueCompletion?.detail?.includes("CheckBox property"));
  assert.ok(itemShapeObjectSelectCompletion?.detail?.includes("CheckBox method"));
  assert.equal(
    itemShapeObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Delete"),
    false,
    "named worksheet Item selector の Shape.OLEFormat.Object は control owner へ解決し、Shape 専用 method を出さない"
  );
  assert.ok(thisWorkbookWorksheetShapeObjectValueCompletion?.detail?.includes("CheckBox property"));
  assert.ok(thisWorkbookWorksheetShapeObjectSelectCompletion?.detail?.includes("CheckBox method"));
  assert.equal(
    thisWorkbookWorksheetShapeObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Delete"),
    false,
    "ThisWorkbook.Worksheets(\"Sheet One\") の Shape.OLEFormat.Object は control owner へ解決し、Shape 専用 method を出さない"
  );
  assert.ok(thisWorkbookWorksheetItemShapeObjectValueCompletion?.detail?.includes("CheckBox property"));
  assert.ok(thisWorkbookWorksheetItemShapeObjectSelectCompletion?.detail?.includes("CheckBox method"));
  assert.equal(
    thisWorkbookWorksheetItemShapeObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Delete"),
    false,
    "ThisWorkbook.Worksheets(\"Sheet One\").Shapes.Item(\"CheckBox1\") も control owner へ解決する"
  );
  assert.equal(
    thisWorkbookWorksheetCodeNameShapeObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Value"),
    false,
    "ThisWorkbook.Worksheets(\"Sheet1\") は codeName 指定なので control owner に昇格しない"
  );
  assert.equal(
    thisWorkbookWorksheetIndexedShapeObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Value"),
    false,
    "ThisWorkbook.Worksheets(1) は index 指定なので control owner に昇格しない"
  );
  assert.equal(
    activeWorkbookWorksheetShapeObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Value"),
    false,
    "ActiveWorkbook.Worksheets(\"Sheet One\") は broad root として未解決を維持する"
  );
  assert.equal(shapeObjectHoverText.includes("CheckBox.Value"), true);
  assert.equal(shapeObjectHoverText.includes("microsoft.office.interop.excel.checkbox.value"), true);
  assert.equal(itemShapeObjectHoverText.includes("CheckBox.Value"), true);
  assert.equal(thisWorkbookWorksheetShapeObjectHoverText.includes("CheckBox.Value"), true);
  assert.equal(thisWorkbookWorksheetItemShapeObjectHoverText.includes("CheckBox.Value"), true);
  assert.equal(shapeObjectSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(itemShapeObjectSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(thisWorkbookWorksheetShapeObjectSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(thisWorkbookWorksheetItemShapeObjectSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(indexedShapeObjectValueHoverSuppressed, true);
  assert.equal(itemIndexedShapeObjectValueHoverSuppressed, true);
  assert.equal(chartShapeObjectValueHoverSuppressed, true);
  assert.equal(chartItemShapeObjectValueHoverSuppressed, true);
  assert.equal(groupedShapeRangeObjectValueHoverSuppressed, true);
  assert.equal(unmatchedShapeObjectValueHoverSuppressed, true);
  assert.equal(itemUnmatchedShapeObjectValueHoverSuppressed, true);
  assert.equal(worksheetNameRootShapeObjectValueHoverSuppressed, true);
  assert.equal(worksheetNameRootItemShapeObjectValueHoverSuppressed, true);
  assert.equal(thisWorkbookWorksheetIndexedShapeObjectValueHoverSuppressed, true);
  assert.equal(activeWorkbookWorksheetShapeObjectValueHoverSuppressed, true);
  assert.equal(thisWorkbookWorksheetCodeNameShapeObjectValueHoverSuppressed, true);
  assert.equal(thisWorkbookWorksheetIndexedShapeObjectSelectSuppressed, true);
  assert.equal(activeWorkbookWorksheetShapeObjectSelectSuppressed, true);
  assert.equal(thisWorkbookWorksheetCodeNameShapeObjectSelectSuppressed, true);
  assert.equal(indexedObjectCallValueHoverSuppressed, true);
  assert.equal(functionShapeNameHoverSuppressed, true);
  assert.equal(itemFunctionShapeNameHoverSuppressed, true);
  assertDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 15, "Name", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 17, "ProgID", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 19, "Value", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 20, "Select", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 21, "Value", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 22, "Select", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 43, "Value", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 45, "Select", {
    modifiers: [],
    type: "function"
  });
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 23, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 24, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 25, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 26, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 27, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 28, "Name");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 29, "Name");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 30, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 31, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 32, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 33, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 39, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 40, "Value");
  assertNoDecodedSemanticToken(shapesBuiltInDocument.getText(), decodedShapesTokens, 41, "Value");

  const worksheetControlShapeNamePathShapePositiveEntries = getWorksheetControlShapeNamePathCompletionEntries("positive", {
    fixture: "packages/extension/test/fixtures/ShapesBuiltIn.bas",
    scope: "extension",
    routeKind: "shape-oleformat"
  });
  const worksheetControlShapeNamePathShapeAlwaysAvailablePositiveEntries =
    worksheetControlShapeNamePathShapePositiveEntries.filter((entry) => entry.rootKind !== "workbook-qualified-matched");
  const worksheetControlShapeNamePathShapeNegativeEntries = getWorksheetControlShapeNamePathCompletionEntries("negative", {
    fixture: "packages/extension/test/fixtures/ShapesBuiltIn.bas",
    scope: "extension",
    routeKind: "shape-oleformat"
  });
  const worksheetControlShapeNamePathShapeClosedEntries = worksheetControlShapeNamePathShapeNegativeEntries.filter(
    (entry) => entry.rootKind === "workbook-qualified-closed"
  );
  const worksheetControlShapeNamePathShapeReasonEntries = worksheetControlShapeNamePathShapeNegativeEntries.filter(
    (entry) => entry.rootKind !== "workbook-qualified-closed"
  );

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT);
  await assertWorkbookRootCompletionCases(
    shapesBuiltInDocument,
    mapExtensionWorksheetControlShapeNamePathPositiveCompletionCases(
      worksheetControlShapeNamePathShapeAlwaysAvailablePositiveEntries,
      (entry) => `${entry.anchor} は ${entry.rootKind} root なので snapshot なしでも control owner へ解決する`
    )
  );
  await assertWorkbookRootClosedCompletionCases(
    shapesBuiltInDocument,
    mapExtensionWorksheetControlShapeNamePathNoCompletionCases(
      [...worksheetControlShapeNamePathShapeReasonEntries, ...worksheetControlShapeNamePathShapeClosedEntries],
      (entry) =>
        entry.rootKind === "workbook-qualified-closed"
          ? `${entry.anchor} は active workbook が閉じている間は control owner に昇格しない`
          : `${entry.anchor} は ${entry.reason} のため control owner に昇格しない`
    )
  );

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_AVAILABLE_SNAPSHOT);
  try {
    await assertWorkbookRootCompletionCases(
      shapesBuiltInDocument,
      mapExtensionWorksheetControlShapeNamePathPositiveCompletionCases(
        worksheetControlShapeNamePathShapePositiveEntries,
        (entry) =>
          entry.rootKind === "workbook-qualified-matched"
            ? `${entry.anchor} は active workbook match 時に control owner へ解決する`
            : `${entry.anchor} は ${entry.rootKind} root として control owner へ解決する`
      )
    );
    await assertWorkbookRootClosedCompletionCases(
      shapesBuiltInDocument,
      mapExtensionWorksheetControlShapeNamePathNoCompletionCases(
        worksheetControlShapeNamePathShapeReasonEntries,
        (entry) => `${entry.anchor} は ${entry.reason} のため match 中でも control owner に昇格しない`
      )
    );
  } finally {
    await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT);
  }

  const workbookQualifiedRootItemShapeStaticCompletionChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
      'CheckBox property',
      "Delete",
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object は control owner へ解決する'
    ],
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
      'CheckBox property',
      "Delete",
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object は control owner へ解決する'
    ]
  ] as const;
  const workbookQualifiedRootItemShapeStaticHoverChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu',
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1") の hover は control owner へ解決する'
    ],
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Valu',
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1") の hover は control owner へ解決する'
    ]
  ] as const;
  const workbookQualifiedRootItemShapeStaticSignatureChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(',
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1") の signature help は control owner へ解決する'
    ],
    [
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1") の signature help は control owner へ解決する'
    ]
  ] as const;
  const workbookQualifiedRootItemShapeNonTargetHoverChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu',
      'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'
    ],
    [
      'ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Valu',
      'ThisWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu',
      'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Valu',
      'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'
    ]
  ] as const;
  const workbookQualifiedRootItemShapeNonTargetCompletionChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.',
      'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'
    ],
    [
      'ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.',
      'ThisWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.',
      'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.',
      'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'
    ]
  ] as const;
  const workbookQualifiedRootItemShapeNonTargetSignatureChecks = [
    [
      'ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(',
      'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので signature help を出さない'
    ],
    [
      'ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Select(',
      'ThisWorkbook.Worksheets.Item(1) は numeric selector なので signature help を出さない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(',
      'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので signature help を出さない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Select(',
      'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので signature help を出さない'
    ]
  ] as const;
  const workbookQualifiedRootItemShapeClosedCompletionChecks = [
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ]
  ] as const;
  const workbookQualifiedRootItemShapeClosedHoverChecks = [
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu',
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Valu',
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ]
  ] as const;
  const workbookQualifiedRootItemShapeClosedSignatureChecks = [
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(',
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ],
    [
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない'
    ]
  ] as const;

  for (const [token, detailFragment, blockedLabel, message] of workbookQualifiedRootItemShapeStaticCompletionChecks) {
    const items = await waitForCompletionLabelStateAtToken(shapesBuiltInDocument, token, "Value", true);
    const completion = items.find((item) => getCompletionItemLabel(item) === "Value");

    assert.ok(completion?.detail?.includes(detailFragment), message);
    assert.equal(hasCompletionItemLabel(items, blockedLabel), false, message);
  }
  for (const [token, message] of workbookQualifiedRootItemShapeStaticHoverChecks) {
    const hovers = await waitForHoverAtToken(shapesBuiltInDocument, token, (items) => items.length > 0);
    assert.equal(getHoverContentsText(hovers[0]).includes("CheckBox.Value"), true, message);
  }
  for (const [token, message] of workbookQualifiedRootItemShapeStaticSignatureChecks) {
    const signatureHelp = await waitForSignatureHelpAtToken(
      shapesBuiltInDocument,
      token,
      (help) => help.signatures.length > 0
    );
    assert.equal(signatureHelp.signatures[0]?.label, "Select(Replace) As Object", message);
  }
  for (const [token, message] of workbookQualifiedRootItemShapeNonTargetCompletionChecks) {
    const items = await waitForCompletionLabelStateAtToken(shapesBuiltInDocument, token, "Value", false);
    assert.equal(hasCompletionItemLabel(items, "Value"), false, message);
  }
  for (const [token, message] of workbookQualifiedRootItemShapeNonTargetHoverChecks) {
    assert.equal(await waitForNoHoverAtToken(shapesBuiltInDocument, token), true, message);
  }
  for (const [token, message] of workbookQualifiedRootItemShapeNonTargetSignatureChecks) {
    assert.equal(await waitForNoSignatureHelpAtToken(shapesBuiltInDocument, token), true, message);
  }
  for (const [token, message] of workbookQualifiedRootItemShapeClosedCompletionChecks) {
    const items = await waitForCompletionLabelStateAtToken(shapesBuiltInDocument, token, "Value", false);
    assert.equal(hasCompletionItemLabel(items, "Value"), false, message);
  }
  for (const [token, message] of workbookQualifiedRootItemShapeClosedHoverChecks) {
    assert.equal(await waitForNoHoverAtToken(shapesBuiltInDocument, token), true, message);
  }
  for (const [token, message] of workbookQualifiedRootItemShapeClosedSignatureChecks) {
    assert.equal(await waitForNoSignatureHelpAtToken(shapesBuiltInDocument, token), true, message);
  }
  assertWorkbookRootSemanticCases(shapesBuiltInDocument.getText(), decodedShapesTokens, [
    [
      'Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value',
      "Value",
      { modifiers: [], type: "variable" },
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object は semantic token を出す'
    ],
    [
      'Call ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(',
      "Select",
      { modifiers: [], type: "function" },
      'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object は semantic token を出す'
    ]
  ] as const);
  assertWorkbookRootNoSemanticCases(shapesBuiltInDocument.getText(), decodedShapesTokens, [
    [
      'Debug.Print ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value',
      "Value",
      'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので semantic token を出さない'
    ],
    [
      'Debug.Print ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Value',
      "Value",
      'ThisWorkbook.Worksheets.Item(1) は numeric selector なので semantic token を出さない'
    ],
    [
      'Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value',
      "Value",
      'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1") は snapshot 未一致の間は semantic token を出さない'
    ]
  ] as const);

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_AVAILABLE_SNAPSHOT);
  try {
    const activeWorkbookBoundShapeObjectCompletionItems = await waitForCompletions(
      shapesBuiltInDocument,
      findPositionAfterToken(shapesBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.'),
      (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
    );
    const activeWorkbookBoundShapeObjectHover = await waitForHover(
      shapesBuiltInDocument,
      findPositionAfterToken(shapesBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu'),
      (hovers) => hovers.length > 0
    );
    const activeWorkbookBoundShapeObjectSelectSignatureHelp = await waitForSignatureHelp(
      shapesBuiltInDocument,
      findPositionAfterToken(shapesBuiltInDocument, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select('),
      (help) => help.signatures.length > 0
    );
    const activeWorkbookBoundShapeObjectValueCompletion = activeWorkbookBoundShapeObjectCompletionItems.find(
      (item) => getCompletionItemLabel(item) === "Value"
    );
    const activeWorkbookBoundShapeObjectSelectCompletion = activeWorkbookBoundShapeObjectCompletionItems.find(
      (item) => getCompletionItemLabel(item) === "Select"
    );
    const activeWorkbookBoundShapeObjectHoverText = getHoverContentsText(activeWorkbookBoundShapeObjectHover[0]);

    assert.ok(activeWorkbookBoundShapeObjectValueCompletion?.detail?.includes("CheckBox property"));
    assert.ok(activeWorkbookBoundShapeObjectSelectCompletion?.detail?.includes("CheckBox method"));
    assert.equal(
      activeWorkbookBoundShapeObjectCompletionItems.some((item) => getCompletionItemLabel(item) === "Delete"),
      false,
      "match 済み active workbook の Shape.OLEFormat.Object は control owner へ解決し、Shape 専用 method を出さない"
    );
    assert.equal(activeWorkbookBoundShapeObjectHoverText.includes("CheckBox.Value"), true);
    assert.equal(activeWorkbookBoundShapeObjectSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");

    const workbookQualifiedRootItemShapeMatchedCompletionChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        'CheckBox property',
        "Delete",
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object は control owner へ解決する'
      ],
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
        'CheckBox property',
        "Delete",
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object は control owner へ解決する'
      ]
    ] as const;
    const workbookQualifiedRootItemShapeMatchedHoverChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu',
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1") の hover は control owner へ解決する'
      ],
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Valu',
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1") の hover は control owner へ解決する'
      ]
    ] as const;
    const workbookQualifiedRootItemShapeMatchedNonTargetCompletionChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.',
        'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'
      ],
      [
        'ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.',
        'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'
      ]
    ] as const;
    const workbookQualifiedRootItemShapeMatchedNonTargetHoverChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu',
        'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので hover を出さない'
      ],
      [
        'ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Valu',
        'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので hover を出さない'
      ]
    ] as const;
    const workbookQualifiedRootItemShapeMatchedNonTargetSignatureChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(',
        'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので signature help を出さない'
      ],
      [
        'ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Select(',
        'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので signature help を出さない'
      ]
    ] as const;
    const workbookQualifiedRootItemShapeMatchedSignatureChecks = [
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(',
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1") の signature help は control owner へ解決する'
      ],
      [
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
        'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1") の signature help は control owner へ解決する'
      ]
    ] as const;

    for (const [token, detailFragment, blockedLabel, message] of workbookQualifiedRootItemShapeMatchedCompletionChecks) {
      const items = await waitForCompletionLabelStateAtToken(shapesBuiltInDocument, token, "Value", true);
      const completion = items.find((item) => getCompletionItemLabel(item) === "Value");

      assert.ok(completion?.detail?.includes(detailFragment), message);
      assert.equal(hasCompletionItemLabel(items, blockedLabel), false, message);
    }
    for (const [token, message] of workbookQualifiedRootItemShapeMatchedHoverChecks) {
      const hovers = await waitForHoverAtToken(shapesBuiltInDocument, token, (items) => items.length > 0);
      assert.equal(getHoverContentsText(hovers[0]).includes("CheckBox.Value"), true, message);
    }
    for (const [token, message] of workbookQualifiedRootItemShapeMatchedNonTargetCompletionChecks) {
      const items = await waitForCompletionLabelStateAtToken(shapesBuiltInDocument, token, "Value", false);
      assert.equal(hasCompletionItemLabel(items, "Value"), false, message);
    }
    for (const [token, message] of workbookQualifiedRootItemShapeMatchedNonTargetHoverChecks) {
      assert.equal(await waitForNoHoverAtToken(shapesBuiltInDocument, token), true, message);
    }
    for (const [token, message] of workbookQualifiedRootItemShapeMatchedNonTargetSignatureChecks) {
      assert.equal(await waitForNoSignatureHelpAtToken(shapesBuiltInDocument, token), true, message);
    }
    for (const [token, message] of workbookQualifiedRootItemShapeMatchedSignatureChecks) {
      const signatureHelp = await waitForSignatureHelpAtToken(
        shapesBuiltInDocument,
        token,
        (help) => help.signatures.length > 0
      );
      assert.equal(signatureHelp.signatures[0]?.label, "Select(Replace) As Object", message);
    }
  } finally {
    await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT);
  }

  const worksheetBroadRootDocument = await vscode.workspace.openTextDocument(
    path.resolve(fixturesPath, "WorksheetBroadRootBuiltIn.bas")
  );
  await vscode.window.showTextDocument(worksheetBroadRootDocument);

  const broadRootMatchedCompletionChecks = mapExtensionWorkbookRootPositiveCompletionCases(
    getSharedWorkbookRootCompletionEntries("worksheetBroadRoot", "positive", { scope: "extension" }),
    (entry) => `${entry.anchor} は control owner へ解決する`
  );
  const broadRootMatchedHoverChecks = mapExtensionWorkbookRootInteractionCases(
    getSharedWorkbookRootInteractionEntries("worksheetBroadRoot", "hover", "positive", { scope: "extension" }),
    (entry) => `${entry.anchor} の hover は control owner へ解決する`
  );
  const broadRootMatchedSignatureChecks = mapExtensionWorkbookRootInteractionCases(
    getSharedWorkbookRootInteractionEntries("worksheetBroadRoot", "signature", "positive", { scope: "extension" }),
    (entry) => `${entry.anchor} の signature help は control owner へ解決する`
  );
  const broadRootNonTargetHoverChecks = mapExtensionWorkbookRootInteractionCases(
    getSharedWorkbookRootInteractionEntries("worksheetBroadRoot", "hover", "negative", { scope: "extension" }),
    (entry) => `${entry.anchor} は broad root family の対象外を維持する`
  );
  const broadRootNonTargetSignatureChecks = mapExtensionWorkbookRootInteractionCases(
    getSharedWorkbookRootInteractionEntries("worksheetBroadRoot", "signature", "negative", { scope: "extension" }),
    (entry) => `${entry.anchor} は broad root family の対象外を維持する`
  );

  await assertWorkbookRootClosedCompletionCases(
    worksheetBroadRootDocument,
    broadRootMatchedCompletionChecks.map(([token, , , message, occurrenceIndex = 0]) => [
      token,
      `no-active-workbook では ${message}`,
      occurrenceIndex
    ] as const)
  );
  await assertWorkbookRootNoHoverCases(
    worksheetBroadRootDocument,
    broadRootMatchedHoverChecks.map(([token, message, occurrenceIndex = 0]) => [
      token,
      `no-active-workbook では ${message}`,
      occurrenceIndex
    ] as const)
  );
  await assertWorkbookRootNoSignatureCases(
    worksheetBroadRootDocument,
    broadRootMatchedSignatureChecks.map(([token, message, occurrenceIndex = 0]) => [
      token,
      `no-active-workbook では ${message}`,
      occurrenceIndex
    ] as const)
  );

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_AVAILABLE_SNAPSHOT);
  try {
    await assertWorkbookRootCompletionCases(worksheetBroadRootDocument, broadRootMatchedCompletionChecks);
    await assertWorkbookRootHoverCases(worksheetBroadRootDocument, broadRootMatchedHoverChecks);
    await assertWorkbookRootSignatureCases(worksheetBroadRootDocument, broadRootMatchedSignatureChecks);
    await assertWorkbookRootNoHoverCases(worksheetBroadRootDocument, broadRootNonTargetHoverChecks);
    await assertWorkbookRootNoSignatureCases(worksheetBroadRootDocument, broadRootNonTargetSignatureChecks);

    await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_MISMATCHED_SNAPSHOT);

    await assertWorkbookRootClosedCompletionCases(
      worksheetBroadRootDocument,
      broadRootMatchedCompletionChecks.map(([token, , , message, occurrenceIndex = 0]) => [
        token,
        `mismatch snapshot では ${message}`,
        occurrenceIndex
      ] as const)
    );
    await assertWorkbookRootNoHoverCases(
      worksheetBroadRootDocument,
      broadRootMatchedHoverChecks.map(([token, message, occurrenceIndex = 0]) => [
        token,
        `mismatch snapshot では ${message}`,
        occurrenceIndex
      ] as const)
    );
    await assertWorkbookRootNoSignatureCases(
      worksheetBroadRootDocument,
      broadRootMatchedSignatureChecks.map(([token, message, occurrenceIndex = 0]) => [
        token,
        `mismatch snapshot では ${message}`,
        occurrenceIndex
      ] as const)
    );
  } finally {
    await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT);
  }

  const worksheetControlCodeNameDocument = await vscode.workspace.openTextDocument(
    path.resolve(fixturesPath, "WorksheetControlCodeName.bas")
  );
  await vscode.window.showTextDocument(worksheetControlCodeNameDocument);

  const controlCodeNameCompletionItems = await waitForCompletions(
    worksheetControlCodeNameDocument,
    findPositionAfterToken(worksheetControlCodeNameDocument, "Sheet1.chkFinished."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const controlCodeNameValueHover = await waitForHover(
    worksheetControlCodeNameDocument,
    findPositionAfterToken(worksheetControlCodeNameDocument, "Sheet1.chkFinished.Valu"),
    (hovers) => hovers.length > 0
  );
  const controlCodeNameSelectSignatureHelp = await waitForSignatureHelp(
    worksheetControlCodeNameDocument,
    findPositionAfterToken(worksheetControlCodeNameDocument, "Sheet1.chkFinished.Select("),
    (help) => help.signatures.length > 0
  );
  const shapeNameControlHoverSuppressed = await waitForNoHover(
    worksheetControlCodeNameDocument,
    findPositionAfterToken(worksheetControlCodeNameDocument, "Sheet1.CheckBox1.Valu")
  );
  const chartControlHoverSuppressed = await waitForNoHover(
    worksheetControlCodeNameDocument,
    findPositionAfterToken(worksheetControlCodeNameDocument, "Chart1.chkFinished.Valu")
  );
  const activeSheetControlHoverSuppressed = await waitForNoHover(
    worksheetControlCodeNameDocument,
    findPositionAfterToken(worksheetControlCodeNameDocument, "ActiveSheet.chkFinished.Valu")
  );
  const shapeNameControlSignatureSuppressed = await waitForNoSignatureHelp(
    worksheetControlCodeNameDocument,
    findPositionAfterToken(worksheetControlCodeNameDocument, "Sheet1.CheckBox1.Select(")
  );
  const chartControlSignatureSuppressed = await waitForNoSignatureHelp(
    worksheetControlCodeNameDocument,
    findPositionAfterToken(worksheetControlCodeNameDocument, "Chart1.chkFinished.Select(")
  );
  const activeSheetControlSignatureSuppressed = await waitForNoSignatureHelp(
    worksheetControlCodeNameDocument,
    findPositionAfterToken(worksheetControlCodeNameDocument, "ActiveSheet.chkFinished.Select(")
  );
  const controlCodeNameLegend = await waitForSemanticTokensLegend(
    worksheetControlCodeNameDocument,
    (legend) => legend.tokenTypes.includes("variable") && legend.tokenTypes.includes("function")
  );
  const controlCodeNameTokens = await waitForSemanticTokens(
    worksheetControlCodeNameDocument,
    (tokens) => tokens.data.length > 0
  );
  const controlCodeNameValueCompletion = controlCodeNameCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const controlCodeNameSelectCompletion = controlCodeNameCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Select"
  );
  const controlCodeNameHoverText = getHoverContentsText(controlCodeNameValueHover[0]);
  const decodedControlCodeNameTokens = decodeSemanticTokens(controlCodeNameTokens, controlCodeNameLegend);

  assert.ok(controlCodeNameValueCompletion?.detail?.includes("Excel CheckBox property"));
  assert.ok(controlCodeNameSelectCompletion?.detail?.includes("Excel CheckBox method"));
  assert.equal(
    controlCodeNameCompletionItems.some((item) => getCompletionItemLabel(item) === "Activate"),
    false,
    "worksheet control code name は control owner へ解決し、OLEObject method を出さない"
  );
  assert.equal(shapeNameControlHoverSuppressed, true, "shape name は direct access の code name 解決へ昇格しない");
  assert.equal(chartControlHoverSuppressed, true, "chartsheet root は control code name 解決へ昇格しない");
  assert.equal(activeSheetControlHoverSuppressed, true, "ActiveSheet root は control code name 解決へ昇格しない");
  assert.equal(controlCodeNameSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(shapeNameControlSignatureSuppressed, true);
  assert.equal(chartControlSignatureSuppressed, true);
  assert.equal(activeSheetControlSignatureSuppressed, true);
  assert.equal(controlCodeNameHoverText.includes("CheckBox.Value"), true);
  assert.equal(controlCodeNameHoverText.includes("microsoft.office.interop.excel.checkbox.value"), true);
  assertDecodedSemanticToken(worksheetControlCodeNameDocument.getText(), decodedControlCodeNameTokens, 8, "Value", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(worksheetControlCodeNameDocument.getText(), decodedControlCodeNameTokens, 12, "Select", {
    modifiers: [],
    type: "function"
  });

  const applicationWorkbookRootDocument = await vscode.workspace.openTextDocument(
    path.resolve(fixturesPath, "ApplicationWorkbookRootBuiltIn.bas")
  );
  await vscode.window.showTextDocument(applicationWorkbookRootDocument);
  const applicationWorkbookRootText = applicationWorkbookRootDocument.getText();
  const applicationWorkbookShadowedDocument = await vscode.workspace.openTextDocument(
    path.resolve(fixturesPath, "ApplicationWorkbookRootShadowed.bas")
  );
  await vscode.window.showTextDocument(applicationWorkbookShadowedDocument);
  const applicationWorkbookShadowedText = applicationWorkbookShadowedDocument.getText();
  await vscode.window.showTextDocument(applicationWorkbookRootDocument);
  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT);
  const applicationWorkbookThisWorkbookDocument = await vscode.workspace.openTextDocument(
    path.resolve(fixturesPath, "ThisWorkbook.cls")
  );
  const applicationWorkbookThisWorkbookDefinitions = await waitForDefinitions(
    applicationWorkbookRootDocument,
    findPositionAfterToken(applicationWorkbookRootDocument, "Application.ThisWorkbook", -1),
    (locations) =>
      locations.some((location) => location.uri.toString() === applicationWorkbookThisWorkbookDocument.uri.toString())
  );

  assert.ok(
    applicationWorkbookThisWorkbookDefinitions.some(
      (location) => location.uri.toString() === applicationWorkbookThisWorkbookDocument.uri.toString()
    ),
    "Application.ThisWorkbook root should resolve to the workbook document module before workbook root matrix assertions"
  );

  const applicationWorkbookStaticPositiveCompletionEntries = getSharedWorkbookRootCompletionEntries(
    "applicationWorkbookRoot",
    "positive",
    {
      scope: "extension",
      state: "static"
    }
  );
  const applicationWorkbookMatchedPositiveCompletionEntries = getSharedWorkbookRootCompletionEntries(
    "applicationWorkbookRoot",
    "positive",
    {
      scope: "extension",
      state: "matched"
    }
  );
  const applicationWorkbookStaticNegativeCompletionEntries = getSharedWorkbookRootCompletionEntries(
    "applicationWorkbookRoot",
    "negative",
    {
      scope: "extension",
      state: "static"
    }
  );
  const applicationWorkbookMatchedNegativeCompletionEntries = getSharedWorkbookRootCompletionEntries(
    "applicationWorkbookRoot",
    "negative",
    {
      scope: "extension",
      state: "matched"
    }
  );
  const applicationWorkbookShadowedCompletionEntries = getSharedWorkbookRootCompletionEntries(
    "applicationWorkbookRoot",
    "negative",
    {
      scope: "extension",
      state: "shadowed"
    }
  );
  const applicationWorkbookStaticHoverEntries = getSharedWorkbookRootInteractionEntries(
    "applicationWorkbookRoot",
    "hover",
    "positive",
    {
      scope: "extension",
      state: "static"
    }
  );
  const applicationWorkbookStaticSignatureEntries = getSharedWorkbookRootInteractionEntries(
    "applicationWorkbookRoot",
    "signature",
    "positive",
    {
      scope: "extension",
      state: "static"
    }
  );
  const applicationWorkbookStaticNegativeHoverEntries = getSharedWorkbookRootInteractionEntries(
    "applicationWorkbookRoot",
    "hover",
    "negative",
    {
      scope: "extension",
      state: "static"
    }
  );
  const applicationWorkbookStaticNegativeSignatureEntries = getSharedWorkbookRootInteractionEntries(
    "applicationWorkbookRoot",
    "signature",
    "negative",
    {
      scope: "extension",
      state: "static"
    }
  );
  const applicationWorkbookMatchedHoverEntries = getSharedWorkbookRootInteractionEntries(
    "applicationWorkbookRoot",
    "hover",
    "positive",
    {
      scope: "extension",
      state: "matched"
    }
  );
  const applicationWorkbookMatchedSignatureEntries = getSharedWorkbookRootInteractionEntries(
    "applicationWorkbookRoot",
    "signature",
    "positive",
    {
      scope: "extension",
      state: "matched"
    }
  );
  const applicationWorkbookMatchedNegativeHoverEntries = getSharedWorkbookRootInteractionEntries(
    "applicationWorkbookRoot",
    "hover",
    "negative",
    {
      scope: "extension",
      state: "matched"
    }
  );
  const applicationWorkbookMatchedNegativeSignatureEntries = getSharedWorkbookRootInteractionEntries(
    "applicationWorkbookRoot",
    "signature",
    "negative",
    {
      scope: "extension",
      state: "matched"
    }
  );
  const applicationWorkbookShadowedHoverEntries = getSharedWorkbookRootInteractionEntries(
    "applicationWorkbookRoot",
    "hover",
    "negative",
    {
      scope: "extension",
      state: "shadowed"
    }
  );
  const applicationWorkbookShadowedSignatureEntries = getSharedWorkbookRootInteractionEntries(
    "applicationWorkbookRoot",
    "signature",
    "negative",
    {
      scope: "extension",
      state: "shadowed"
    }
  );
  const applicationWorkbookStaticCompletionChecks = mapExtensionWorkbookRootPositiveCompletionCases(
    applicationWorkbookStaticPositiveCompletionEntries,
    (entry) => `${entry.anchor} は control owner へ解決する`
  );
  const applicationWorkbookStaticHoverChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookStaticHoverEntries,
    (entry) => `${entry.anchor} の hover は control owner へ解決する`
  );
  const applicationWorkbookStaticSignatureChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookStaticSignatureEntries,
    (entry) => `${entry.anchor} の signature help は control owner へ解決する`
  );
  const applicationWorkbookStaticNonTargetCompletionChecks = mapExtensionWorkbookRootClosedCompletionCases(
    applicationWorkbookStaticNegativeCompletionEntries,
    (entry) =>
      entry.reason === "snapshot-closed"
        ? `${entry.anchor} は snapshot 未一致の間は broad root を開かない`
        : `${entry.anchor} は control owner に昇格しない`
  );
  const applicationWorkbookStaticNonTargetHoverChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookStaticNegativeHoverEntries.filter((entry) => entry.reason !== "snapshot-closed"),
    (entry) =>
      entry.reason === "non-target-root"
        ? `${entry.anchor} は workbook root family に昇格しない`
        : `${entry.anchor} は control owner に昇格しない`
  );
  const applicationWorkbookStaticNonTargetSignatureChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookStaticNegativeSignatureEntries.filter((entry) => entry.reason !== "snapshot-closed"),
    (entry) =>
      entry.reason === "non-target-root"
        ? `${entry.anchor} は workbook root family に昇格しない`
        : `${entry.anchor} は control owner に昇格しない`
  );
  const applicationWorkbookClosedCompletionChecks = mapExtensionWorkbookRootClosedCompletionCases(
    applicationWorkbookMatchedPositiveCompletionEntries,
    (entry) => `${entry.anchor} は snapshot 未一致の間は broad root を開かない`
  );
  const applicationWorkbookClosedHoverChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookStaticNegativeHoverEntries.filter((entry) => entry.reason === "snapshot-closed"),
    (entry) => `${entry.anchor} は snapshot 未一致の間は broad root を開かない`
  );
  const applicationWorkbookClosedSignatureChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookStaticNegativeSignatureEntries.filter((entry) => entry.reason === "snapshot-closed"),
    (entry) => `${entry.anchor} は snapshot 未一致の間は broad root を開かない`
  );

  await assertWorkbookRootCompletionCases(applicationWorkbookRootDocument, applicationWorkbookStaticCompletionChecks);
  await assertWorkbookRootHoverCases(applicationWorkbookRootDocument, applicationWorkbookStaticHoverChecks);
  await assertWorkbookRootSignatureCases(applicationWorkbookRootDocument, applicationWorkbookStaticSignatureChecks);
  await assertWorkbookRootClosedCompletionCases(applicationWorkbookRootDocument, applicationWorkbookStaticNonTargetCompletionChecks);
  await assertWorkbookRootNoHoverCases(applicationWorkbookRootDocument, applicationWorkbookStaticNonTargetHoverChecks);
  await assertWorkbookRootNoSignatureCases(applicationWorkbookRootDocument, applicationWorkbookStaticNonTargetSignatureChecks);
  await assertWorkbookRootClosedCompletionCases(applicationWorkbookRootDocument, applicationWorkbookClosedCompletionChecks);
  await assertWorkbookRootNoHoverCases(applicationWorkbookRootDocument, applicationWorkbookClosedHoverChecks);
  await assertWorkbookRootNoSignatureCases(applicationWorkbookRootDocument, applicationWorkbookClosedSignatureChecks);

  const applicationWorkbookLegend = await waitForSemanticTokensLegend(
    applicationWorkbookRootDocument,
    (legend) => legend.tokenTypes.includes("variable") && legend.tokenTypes.includes("function")
  );
  let applicationWorkbookTokens = await waitForSemanticTokens(
    applicationWorkbookRootDocument,
    (tokens) => tokens.data.length > 0
  );
  let decodedApplicationWorkbookTokens = decodeSemanticTokens(applicationWorkbookTokens, applicationWorkbookLegend);
  assertWorkbookRootSemanticCases(
    applicationWorkbookRootText,
    decodedApplicationWorkbookTokens,
    mapExtensionWorkbookRootSemanticCases(
      getSharedWorkbookRootSemanticEntries("applicationWorkbookRoot", "positive", {
        scope: "extension",
        state: "static"
      }),
      (entry) => `${entry.anchor} は semantic token を出す`
    )
  );
  assertWorkbookRootNoSemanticCases(
    applicationWorkbookRootText,
    decodedApplicationWorkbookTokens,
    mapExtensionWorkbookRootNoSemanticCases(
      getSharedWorkbookRootSemanticEntries("applicationWorkbookRoot", "negative", {
        scope: "extension",
        state: "static"
      }),
      (entry) =>
        entry.reason === "snapshot-closed"
          ? `${entry.anchor} は snapshot 未一致の間は semantic token を出さない`
          : `${entry.anchor} は semantic token を出さない`
    )
  );
  const applicationWorkbookShadowNoSemanticCases = mapExtensionWorkbookRootNoSemanticCases(
    getSharedWorkbookRootSemanticEntries("applicationWorkbookRoot", "negative", {
      scope: "extension",
      state: "shadowed"
    }),
    (entry) => `${entry.anchor} は shadowed Application qualifier のため semantic token を出さない`
  );
  let decodedApplicationWorkbookShadowedTokens = await waitForNoSemanticTokensByCases(
    applicationWorkbookShadowedDocument,
    applicationWorkbookLegend,
    applicationWorkbookShadowedText,
    applicationWorkbookShadowNoSemanticCases
  );

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_AVAILABLE_SNAPSHOT);

  const applicationWorkbookMatchedCompletionChecks = mapExtensionWorkbookRootPositiveCompletionCases(
    applicationWorkbookMatchedPositiveCompletionEntries,
    (entry) => `${entry.anchor} は control owner へ解決する`
  );
  const applicationWorkbookMatchedHoverChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookMatchedHoverEntries,
    (entry) => `${entry.anchor} の hover は control owner へ解決する`
  );
  const applicationWorkbookMatchedSignatureChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookMatchedSignatureEntries,
    (entry) => `${entry.anchor} の signature help は control owner へ解決する`
  );
  const applicationWorkbookMatchedNonTargetCompletionChecks = mapExtensionWorkbookRootClosedCompletionCases(
    applicationWorkbookMatchedNegativeCompletionEntries,
    (entry) => `${entry.anchor} は control owner に昇格しない`
  );
  const applicationWorkbookMatchedNonTargetHoverChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookMatchedNegativeHoverEntries,
    (entry) =>
      entry.reason === "non-target-root"
        ? `snapshot 一致後も ${entry.anchor} は workbook root family に昇格しない`
        : `${entry.anchor} は control owner に昇格しない`
  );
  const applicationWorkbookMatchedNonTargetSignatureChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookMatchedNegativeSignatureEntries,
    (entry) =>
      entry.reason === "non-target-root"
        ? `snapshot 一致後も ${entry.anchor} は workbook root family に昇格しない`
        : `${entry.anchor} は control owner に昇格しない`
  );
  const applicationWorkbookMatchedShadowCompletionChecks = mapExtensionWorkbookRootClosedCompletionCases(
    applicationWorkbookShadowedCompletionEntries,
    (entry) => `${entry.anchor} は control owner に昇格しない`
  );
  const applicationWorkbookMatchedShadowHoverChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookShadowedHoverEntries,
    (entry) =>
      entry.rootKind === "ThisWorkbook"
        ? "shadowed Application.ThisWorkbook root は hover を出さない"
        : "shadowed Application.ActiveWorkbook root は hover を出さない"
  );
  const applicationWorkbookMatchedShadowSignatureChecks = mapExtensionWorkbookRootInteractionCases(
    applicationWorkbookShadowedSignatureEntries,
    (entry) =>
      entry.rootKind === "ThisWorkbook"
        ? "shadowed Application.ThisWorkbook root は signature help を出さない"
        : "shadowed Application.ActiveWorkbook root は signature help を出さない"
  );

  await assertWorkbookRootCompletionCases(applicationWorkbookRootDocument, applicationWorkbookMatchedCompletionChecks);
  await assertWorkbookRootHoverCases(applicationWorkbookRootDocument, applicationWorkbookMatchedHoverChecks);
  await assertWorkbookRootSignatureCases(applicationWorkbookRootDocument, applicationWorkbookMatchedSignatureChecks);
  await assertWorkbookRootClosedCompletionCases(applicationWorkbookRootDocument, applicationWorkbookMatchedNonTargetCompletionChecks);
  await assertWorkbookRootNoHoverCases(applicationWorkbookRootDocument, applicationWorkbookMatchedNonTargetHoverChecks);
  await assertWorkbookRootNoSignatureCases(applicationWorkbookRootDocument, applicationWorkbookMatchedNonTargetSignatureChecks);
  await assertWorkbookRootClosedCompletionCases(applicationWorkbookShadowedDocument, applicationWorkbookMatchedShadowCompletionChecks);
  await assertWorkbookRootNoHoverCases(applicationWorkbookShadowedDocument, applicationWorkbookMatchedShadowHoverChecks);
  await assertWorkbookRootNoSignatureCases(applicationWorkbookShadowedDocument, applicationWorkbookMatchedShadowSignatureChecks);

  applicationWorkbookTokens = await waitForSemanticTokens(applicationWorkbookRootDocument, (tokens) => tokens.data.length > 0);
  decodedApplicationWorkbookTokens = decodeSemanticTokens(applicationWorkbookTokens, applicationWorkbookLegend);
  assertWorkbookRootSemanticCases(
    applicationWorkbookRootText,
    decodedApplicationWorkbookTokens,
    mapExtensionWorkbookRootSemanticCases(
      getSharedWorkbookRootSemanticEntries("applicationWorkbookRoot", "positive", {
        scope: "extension",
        state: "matched"
      }),
      (entry) => `${entry.anchor} は semantic token を出す`
    )
  );
  assertWorkbookRootNoSemanticCases(
    applicationWorkbookRootText,
    decodedApplicationWorkbookTokens,
    mapExtensionWorkbookRootNoSemanticCases(
      getSharedWorkbookRootSemanticEntries("applicationWorkbookRoot", "negative", {
        scope: "extension",
        state: "matched"
      }),
      (entry) => `${entry.anchor} は semantic token を出さない`
    )
  );
  decodedApplicationWorkbookShadowedTokens = await waitForNoSemanticTokensByCases(
    applicationWorkbookShadowedDocument,
    applicationWorkbookLegend,
    applicationWorkbookShadowedText,
    applicationWorkbookShadowNoSemanticCases
  );

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT);

  const dialogSheetBuiltInDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "DialogSheetBuiltIn.bas"));
  await vscode.window.showTextDocument(dialogSheetBuiltInDocument);

  const dialogSheetsCollectionCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const dialogSheetItemCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "SaveAs")
  );
  const dialogFrameCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1).DialogFrame."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Caption")
  );
  const namedDialogSheetCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, 'DialogSheets("Dialog1").'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Activate")
  );
  const namedDialogFrameCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, 'DialogSheets("Dialog1").DialogFrame.'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Caption")
  );
  const groupedDialogSheetsCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, 'DialogSheets(Array("Dialog1", "Dialog2")).'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const itemDialogSheetCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets.Item(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "SaveAs")
  );
  const applicationDialogSheetsCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "Application.DialogSheets."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const applicationDialogFrameCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "Application.DialogSheets(1).DialogFrame."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Text")
  );
  const activeWorkbookDialogSheetCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "ActiveWorkbook.DialogSheets(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "SaveAs")
  );
  const thisWorkbookDialogSheetCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "ThisWorkbook.DialogSheets(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "SaveAs")
  );
  const dialogSheetEvaluateSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1).Evaluate("),
    (help) => help.signatures.length > 0
  );
  const dialogFrameSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1).DialogFrame.Select("),
    (help) => help.signatures.length > 0
  );
  const dialogSheetSaveAsSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1).SaveAs("),
    (help) => help.signatures.length > 0
  );
  const dialogSheetExportSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1).ExportAsFixedFormat("),
    (help) => help.signatures.length > 0
  );
  const groupedDialogSheetSaveAsSuppressed = await waitForNoSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, 'DialogSheets(Array("Dialog1", "Dialog2")).SaveAs(')
  );
  const itemDialogSheetSaveAsSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets.Item(1).SaveAs("),
    (help) => help.signatures.length > 0
  );
  const applicationDialogSheetEvaluateSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "Application.DialogSheets(1).Evaluate("),
    (help) => help.signatures.length > 0
  );
  const activeWorkbookDialogSheetSaveAsSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "ActiveWorkbook.DialogSheets(1).SaveAs("),
    (help) => help.signatures.length > 0
  );
  const activeWorkbookDialogFrameSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "ActiveWorkbook.DialogSheets(1).DialogFrame.Select("),
    (help) => help.signatures.length > 0
  );
  const itemDialogFrameSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets.Item(1).DialogFrame.Select("),
    (help) => help.signatures.length > 0
  );
  const thisWorkbookDialogSheetSaveAsSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "ThisWorkbook.DialogSheets(1).SaveAs("),
    (help) => help.signatures.length > 0
  );
  const thisWorkbookDialogFrameSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "ThisWorkbook.DialogSheets(1).DialogFrame.Select("),
    (help) => help.signatures.length > 0
  );
  const groupedApplicationDialogSheetSaveAsSuppressed = await waitForNoSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, 'Application.DialogSheets(Array("Dialog1", "Dialog2")).SaveAs(')
  );
  const groupedDialogFrameSelectSuppressed = await waitForNoSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, 'DialogSheets(Array("Dialog1", "Dialog2")).DialogFrame.Select(')
  );
  const dialogFrameCaptionPropertySuppressed = await waitForNoSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1).DialogFrame.Caption(")
  );
  const applicationDialogFrameTextPropertySuppressed = await waitForNoSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "Application.DialogSheets(1).DialogFrame.Text(")
  );
  const dialogSheetHover = await waitForHover(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1).SaveA"),
    (hovers) => hovers.length > 0
  );
  const dialogFrameCaptionHover = await waitForHover(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1).DialogFrame.Capti"),
    (hovers) => hovers.length > 0
  );
  const applicationDialogFrameTextHover = await waitForHover(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "Application.DialogSheets(1).DialogFrame.Tex"),
    (hovers) => hovers.length > 0
  );
  const activeWorkbookDialogSheetHover = await waitForHover(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "ActiveWorkbook.DialogSheets(1).SaveA"),
    (hovers) => hovers.length > 0
  );
  const dialogSheetLegend = await waitForSemanticTokensLegend(
    dialogSheetBuiltInDocument,
    (legend) => legend.tokenTypes.length > 0
  );
  const dialogSheetTokens = await waitForSemanticTokens(
    dialogSheetBuiltInDocument,
    (tokens) => tokens.data.length > 0
  );
  const decodedDialogSheetTokens = decodeSemanticTokens(dialogSheetTokens, dialogSheetLegend);
  const dialogSheetsCountCompletion = dialogSheetsCollectionCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const dialogFramePropertyCompletion = dialogSheetItemCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "DialogFrame"
  );
  const dialogSheetSaveAsCompletion = dialogSheetItemCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SaveAs"
  );
  const dialogSheetEvaluateCompletion = dialogSheetItemCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Evaluate"
  );
  const dialogFrameCaptionCompletion = dialogFrameCompletionItems.find((item) => getCompletionItemLabel(item) === "Caption");
  const dialogFrameSelectCompletion = dialogFrameCompletionItems.find((item) => getCompletionItemLabel(item) === "Select");
  const namedDialogSheetActivateCompletion = namedDialogSheetCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Activate"
  );
  const namedDialogFrameCaptionCompletion = namedDialogFrameCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Caption"
  );
  const groupedDialogSheetsCountCompletion = groupedDialogSheetsCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const itemDialogSheetSaveAsCompletion = itemDialogSheetCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SaveAs"
  );
  const applicationDialogSheetsCountCompletion = applicationDialogSheetsCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const applicationDialogFrameTextCompletion = applicationDialogFrameCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Text"
  );
  const activeWorkbookDialogSheetSaveAsCompletion = activeWorkbookDialogSheetCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SaveAs"
  );
  const thisWorkbookDialogSheetSaveAsCompletion = thisWorkbookDialogSheetCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SaveAs"
  );
  const dialogSheetHoverText = getHoverContentsText(dialogSheetHover[0]);
  const dialogFrameCaptionHoverText = getHoverContentsText(dialogFrameCaptionHover[0]);
  const applicationDialogFrameTextHoverText = getHoverContentsText(applicationDialogFrameTextHover[0]);
  const activeWorkbookDialogSheetHoverText = getHoverContentsText(activeWorkbookDialogSheetHover[0]);

  assert.ok(dialogSheetsCountCompletion?.detail?.includes("Excel DialogSheets property"));
  assert.ok(dialogFramePropertyCompletion?.detail?.includes("Excel DialogSheet property"));
  assert.ok(dialogFramePropertyCompletion?.detail?.includes("DialogFrame"));
  assert.ok(dialogSheetSaveAsCompletion?.detail?.includes("Excel DialogSheet method"));
  assert.ok(dialogSheetEvaluateCompletion?.detail?.includes("Excel DialogSheet method"));
  assert.ok(dialogFrameCaptionCompletion?.detail?.includes("Excel DialogFrame property"));
  assert.ok(dialogFrameSelectCompletion?.detail?.includes("Excel DialogFrame method"));
  assert.ok(namedDialogSheetActivateCompletion?.detail?.includes("Excel DialogSheet method"));
  assert.ok(namedDialogFrameCaptionCompletion?.detail?.includes("Excel DialogFrame property"));
  assert.ok(groupedDialogSheetsCountCompletion?.detail?.includes("Excel DialogSheets property"));
  assert.ok(itemDialogSheetSaveAsCompletion?.detail?.includes("Excel DialogSheet method"));
  assert.ok(applicationDialogSheetsCountCompletion?.detail?.includes("Excel DialogSheets property"));
  assert.ok(applicationDialogFrameTextCompletion?.detail?.includes("Excel DialogFrame property"));
  assert.ok(activeWorkbookDialogSheetSaveAsCompletion?.detail?.includes("Excel DialogSheet method"));
  assert.ok(thisWorkbookDialogSheetSaveAsCompletion?.detail?.includes("Excel DialogSheet method"));
  assert.equal(
    groupedDialogSheetsCompletionItems.some((item) => getCompletionItemLabel(item) === "SaveAs"),
    false,
    "grouped DialogSheets selector should stay on the DialogSheets collection"
  );
  assert.equal(
    groupedDialogSheetsCompletionItems.some((item) => getCompletionItemLabel(item) === "DialogFrame"),
    false,
    "grouped DialogSheets selector should not expose DialogFrame"
  );
  assert.equal(dialogSheetEvaluateSignatureHelp.signatures[0]?.label, "Evaluate(Name) As Object");
  assert.equal(dialogFrameSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(dialogFrameSelectSignatureHelp.signatures[0]?.parameters.length, 1);
  assert.ok(
    getSignatureDocumentation(dialogFrameSelectSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "想定型: Object"
    )
  );
  assert.ok(
    getSignatureDocumentation(dialogFrameSelectSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "省略可能"
    )
  );
  assert.equal(dialogSheetSaveAsSignatureHelp.signatures[0]?.label, "SaveAs(Filename, FileFormat, Password, ..., Local)");
  assert.equal(dialogSheetSaveAsSignatureHelp.signatures[0]?.parameters.length, 10);
  assert.ok(
    getSignatureDocumentation(dialogSheetSaveAsSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes("必須引数")
  );
  assert.ok(
    getSignatureDocumentation(dialogSheetSaveAsSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("省略可能")
  );
  assert.equal(
    itemDialogSheetSaveAsSignatureHelp.signatures[0]?.label,
    "SaveAs(Filename, FileFormat, Password, ..., Local)"
  );
  assert.equal(itemDialogSheetSaveAsSignatureHelp.signatures[0]?.parameters.length, 10);
  assert.equal(applicationDialogSheetEvaluateSignatureHelp.signatures[0]?.label, "Evaluate(Name) As Object");
  assert.equal(
    activeWorkbookDialogSheetSaveAsSignatureHelp.signatures[0]?.label,
    "SaveAs(Filename, FileFormat, Password, ..., Local)"
  );
  assert.equal(activeWorkbookDialogFrameSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(itemDialogFrameSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(
    thisWorkbookDialogSheetSaveAsSignatureHelp.signatures[0]?.label,
    "SaveAs(Filename, FileFormat, Password, ..., Local)"
  );
  assert.equal(thisWorkbookDialogFrameSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(
    dialogSheetExportSignatureHelp.signatures[0]?.label,
    "ExportAsFixedFormat(Type, Filename, Quality, ..., FixedFormatExtClassPtr)"
  );
  assert.equal(groupedDialogSheetSaveAsSuppressed, true);
  assert.equal(groupedApplicationDialogSheetSaveAsSuppressed, true);
  assert.equal(groupedDialogFrameSelectSuppressed, true);
  assert.equal(dialogFrameCaptionPropertySuppressed, true);
  assert.equal(applicationDialogFrameTextPropertySuppressed, true);
  assert.ok(dialogSheetHoverText.includes("SaveAs(Filename, FileFormat, Password, ..., Local)"));
  assert.ok(dialogSheetHoverText.includes("microsoft.office.interop.excel.dialogsheet.saveas"));
  assert.ok(dialogFrameCaptionHoverText.includes("DialogFrame.Caption"));
  assert.ok(dialogFrameCaptionHoverText.includes("microsoft.office.interop.excel.dialogframe.caption"));
  assert.ok(applicationDialogFrameTextHoverText.includes("DialogFrame.Text"));
  assert.ok(applicationDialogFrameTextHoverText.includes("microsoft.office.interop.excel.dialogframe.text"));
  assert.ok(activeWorkbookDialogSheetHoverText.includes("microsoft.office.interop.excel.dialogsheet.saveas"));
  assertDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 4, "DialogSheets", {
    modifiers: [],
    type: "type"
  });
  assertDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 8, "Evaluate", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 9, "SaveAs", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 13, "DialogSheets", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 15, "SaveAs", {
    modifiers: [],
    type: "function"
  });
  assertNoDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 16, "SaveAs");
  assertDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 18, "DialogFrame", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 19, "Caption", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 20, "Select", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 22, "Text", {
    modifiers: [],
    type: "variable"
  });
  assertNoDecodedSemanticToken(dialogSheetBuiltInDocument.getText(), decodedDialogSheetTokens, 25, "Select");

  const dialogSheetControlCollectionDocument = await vscode.workspace.openTextDocument(
    path.resolve(fixturesPath, "DialogSheetControlCollection.bas")
  );
  await vscode.window.showTextDocument(dialogSheetControlCollectionDocument);

  const buttonsCollectionCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const indexedButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Caption")
  );
  const hexButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(&H1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Caption")
  );
  const octalButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(&O7)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Caption")
  );
  const suffixButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(1#)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Caption")
  );
  const exponentButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(1E+2)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Caption")
  );
  const namedButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).Buttons("Button 1").'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Caption")
  );
  const dynamicButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(index)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const groupedButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(Array(1, 2))."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const itemButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons.Item(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Caption")
  );
  const namedItemButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).Buttons.Item("Button 1").'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Caption")
  );
  const dynamicItemButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons.Item(index)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Count")
  );
  const indexedCheckBoxCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).CheckBoxes(1)."),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const namedItemCheckBoxCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).CheckBoxes.Item("Check 1").'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const namedOptionButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).OptionButtons("Option 1").'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const namedItemOptionButtonCompletionItems = await waitForCompletions(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).OptionButtons.Item("Option 1").'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Value")
  );
  const buttonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(1).Select("),
    (help) => help.signatures.length > 0
  );
  const hexButtonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(&H1).Select("),
    (help) => help.signatures.length > 0
  );
  const octalButtonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(&O7).Select("),
    (help) => help.signatures.length > 0
  );
  const suffixButtonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(1#).Select("),
    (help) => help.signatures.length > 0
  );
  const exponentButtonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(1E+2).Select("),
    (help) => help.signatures.length > 0
  );
  const itemButtonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons.Item(1).Select("),
    (help) => help.signatures.length > 0
  );
  const namedItemButtonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).Buttons.Item("Button 1").Select('),
    (help) => help.signatures.length > 0
  );
  const checkBoxSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).CheckBoxes(1).Select("),
    (help) => help.signatures.length > 0
  );
  const itemCheckBoxSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).CheckBoxes.Item(1).Select("),
    (help) => help.signatures.length > 0
  );
  const namedItemCheckBoxSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).CheckBoxes.Item("Check 1").Select('),
    (help) => help.signatures.length > 0
  );
  const optionButtonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).OptionButtons("Option 1").Select('),
    (help) => help.signatures.length > 0
  );
  const itemOptionButtonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).OptionButtons.Item(1).Select("),
    (help) => help.signatures.length > 0
  );
  const namedItemOptionButtonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).OptionButtons.Item("Option 1").Select('),
    (help) => help.signatures.length > 0
  );
  const applicationButtonSelectSignatureHelp = await waitForSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "Application.DialogSheets(1).Buttons(1).Select("),
    (help) => help.signatures.length > 0
  );
  const dynamicButtonSelectSuppressed = await waitForNoSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(index).Select(")
  );
  const dynamicItemButtonSelectSuppressed = await waitForNoSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons.Item(index).Select(")
  );
  const groupedButtonSelectSuppressed = await waitForNoSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(Array(1, 2)).Select(")
  );
  const checkBoxValueSuppressed = await waitForNoSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).CheckBoxes(1).Value(")
  );
  const optionButtonValueSuppressed = await waitForNoSignatureHelp(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).OptionButtons("Option 1").Value(')
  );
  const buttonCaptionHover = await waitForHover(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).Buttons(1).Capti"),
    (hovers) => hovers.length > 0
  );
  const checkBoxValueHover = await waitForHover(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, "DialogSheets(1).CheckBoxes(1).Valu"),
    (hovers) => hovers.length > 0
  );
  const optionButtonValueHover = await waitForHover(
    dialogSheetControlCollectionDocument,
    findPositionAfterToken(dialogSheetControlCollectionDocument, 'DialogSheets(1).OptionButtons("Option 1").Valu'),
    (hovers) => hovers.length > 0
  );
  const controlCollectionLegend = await waitForSemanticTokensLegend(
    dialogSheetControlCollectionDocument,
    (legend) => legend.tokenTypes.length > 0
  );
  const controlCollectionTokens = await waitForSemanticTokens(
    dialogSheetControlCollectionDocument,
    (tokens) => tokens.data.length > 0
  );
  const decodedControlCollectionTokens = decodeSemanticTokens(controlCollectionTokens, controlCollectionLegend);
  const buttonsCountCompletion = buttonsCollectionCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const buttonsItemCompletion = buttonsCollectionCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Item"
  );
  const indexedButtonCaptionCompletion = indexedButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Caption"
  );
  const hexButtonCaptionCompletion = hexButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Caption"
  );
  const octalButtonCaptionCompletion = octalButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Caption"
  );
  const suffixButtonCaptionCompletion = suffixButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Caption"
  );
  const exponentButtonCaptionCompletion = exponentButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Caption"
  );
  const indexedButtonSelectCompletion = indexedButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Select"
  );
  const namedButtonCaptionCompletion = namedButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Caption"
  );
  const dynamicButtonsCountCompletion = dynamicButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const groupedButtonsCountCompletion = groupedButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const itemButtonCaptionCompletion = itemButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Caption"
  );
  const namedItemButtonCaptionCompletion = namedItemButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Caption"
  );
  const dynamicItemButtonsCountCompletion = dynamicItemButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const checkBoxValueCompletion = indexedCheckBoxCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const namedItemCheckBoxValueCompletion = namedItemCheckBoxCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const optionButtonValueCompletion = namedOptionButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const namedItemOptionButtonValueCompletion = namedItemOptionButtonCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Value"
  );
  const buttonCaptionHoverText = getHoverContentsText(buttonCaptionHover[0]);
  const checkBoxValueHoverText = getHoverContentsText(checkBoxValueHover[0]);
  const optionButtonValueHoverText = getHoverContentsText(optionButtonValueHover[0]);

  assert.ok(buttonsCountCompletion?.detail?.includes("Excel Buttons property"));
  assert.ok(buttonsItemCompletion?.detail?.includes("Excel Buttons method"));
  assert.ok(indexedButtonCaptionCompletion?.detail?.includes("Excel Button property"));
  assert.ok(hexButtonCaptionCompletion?.detail?.includes("Excel Button property"));
  assert.ok(octalButtonCaptionCompletion?.detail?.includes("Excel Button property"));
  assert.ok(suffixButtonCaptionCompletion?.detail?.includes("Excel Button property"));
  assert.ok(exponentButtonCaptionCompletion?.detail?.includes("Excel Button property"));
  assert.ok(indexedButtonSelectCompletion?.detail?.includes("Excel Button method"));
  assert.ok(namedButtonCaptionCompletion?.detail?.includes("Excel Button property"));
  assert.ok(dynamicButtonsCountCompletion?.detail?.includes("Excel Buttons property"));
  assert.ok(groupedButtonsCountCompletion?.detail?.includes("Excel Buttons property"));
  assert.ok(itemButtonCaptionCompletion?.detail?.includes("Excel Button property"));
  assert.ok(namedItemButtonCaptionCompletion?.detail?.includes("Excel Button property"));
  assert.ok(dynamicItemButtonsCountCompletion?.detail?.includes("Excel Buttons property"));
  assert.ok(checkBoxValueCompletion?.detail?.includes("Excel CheckBox property"));
  assert.ok(namedItemCheckBoxValueCompletion?.detail?.includes("Excel CheckBox property"));
  assert.ok(optionButtonValueCompletion?.detail?.includes("Excel OptionButton property"));
  assert.ok(namedItemOptionButtonValueCompletion?.detail?.includes("Excel OptionButton property"));
  assert.equal(
    dynamicButtonCompletionItems.some((item) => getCompletionItemLabel(item) === "Caption"),
    false,
    "expression selector Buttons should stay on the Buttons collection"
  );
  assert.equal(
    groupedButtonCompletionItems.some((item) => getCompletionItemLabel(item) === "Caption"),
    false,
    "grouped selector Buttons should stay on the Buttons collection"
  );
  assert.equal(
    dynamicItemButtonCompletionItems.some((item) => getCompletionItemLabel(item) === "Caption"),
    false,
    "expression selector Buttons.Item should stay on the Buttons collection"
  );
  assert.equal(buttonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(hexButtonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(octalButtonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(suffixButtonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(exponentButtonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(itemButtonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(namedItemButtonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(checkBoxSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(itemCheckBoxSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(namedItemCheckBoxSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(optionButtonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(itemOptionButtonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(namedItemOptionButtonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(applicationButtonSelectSignatureHelp.signatures[0]?.label, "Select(Replace) As Object");
  assert.equal(dynamicButtonSelectSuppressed, true);
  assert.equal(dynamicItemButtonSelectSuppressed, true);
  assert.equal(groupedButtonSelectSuppressed, true);
  assert.equal(checkBoxValueSuppressed, true);
  assert.equal(optionButtonValueSuppressed, true);
  assert.ok(buttonCaptionHoverText.includes("Button.Caption"));
  assert.ok(buttonCaptionHoverText.includes("microsoft.office.interop.excel.button.caption"));
  assert.ok(checkBoxValueHoverText.includes("CheckBox.Value"));
  assert.ok(checkBoxValueHoverText.includes("microsoft.office.interop.excel.checkbox.value"));
  assert.ok(optionButtonValueHoverText.includes("OptionButton.Value"));
  assert.ok(optionButtonValueHoverText.includes("microsoft.office.interop.excel.optionbutton.value"));
  assertDecodedSemanticToken(dialogSheetControlCollectionDocument.getText(), decodedControlCollectionTokens, 5, "Buttons", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(dialogSheetControlCollectionDocument.getText(), decodedControlCollectionTokens, 18, "Select", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(dialogSheetControlCollectionDocument.getText(), decodedControlCollectionTokens, 25, "Caption", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(dialogSheetControlCollectionDocument.getText(), decodedControlCollectionTokens, 26, "Value", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(dialogSheetControlCollectionDocument.getText(), decodedControlCollectionTokens, 27, "Value", {
    modifiers: [],
    type: "variable"
  });
  assertNoDecodedSemanticToken(dialogSheetControlCollectionDocument.getText(), decodedControlCollectionTokens, 30, "Select");
  assertNoDecodedSemanticToken(dialogSheetControlCollectionDocument.getText(), decodedControlCollectionTokens, 31, "Select");
  assertNoDecodedSemanticToken(dialogSheetControlCollectionDocument.getText(), decodedControlCollectionTokens, 32, "Select");

  const definitions = await waitForDefinitions(
    consumerDocument,
    new vscode.Position(5, 18),
    (locations) => locations.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "PublicApi.bas")))
  );
  assert.ok(definitions.length > 0, "definition should resolve across files");
  assert.ok(definitions.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "PublicApi.bas"))));

  const references = await waitForReferences(
    consumerDocument,
    new vscode.Position(5, 18),
    (locations) =>
      locations.length >= 2 &&
      locations.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "Consumer.bas"))) &&
      locations.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "PublicApi.bas")))
  );
  assert.ok(references.length >= 2, "references should include declaration and usage");
  assert.ok(references.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "Consumer.bas"))));
  assert.ok(references.some((location) => location.uri.fsPath.endsWith(path.join("fixtures", "PublicApi.bas"))));

  const consumerCompletionDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "ConsumerCompletion.bas"));
  await vscode.window.showTextDocument(consumerCompletionDocument);

  const narrowedCompletionItems = await waitForCompletions(
    consumerCompletionDocument,
    new vscode.Position(5, 17),
    (items) => items.some((item) => item.label === "PublicMessage")
  );
  assert.ok(
    narrowedCompletionItems.some((item) => item.label === "PublicMessage"),
    "type-aware completion should keep compatible candidates"
  );
  assert.equal(
    narrowedCompletionItems.some((item) => item.label === "PublicNumber"),
    false,
    "type-aware completion should hide incompatible candidates"
  );

  const consumerSignatureDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "ConsumerSignature.bas"));
  await vscode.window.showTextDocument(consumerSignatureDocument);

  const signatureHelp = await waitForSignatureHelp(
    consumerSignatureDocument,
    new vscode.Position(5, 38),
    (help) => help.signatures.length > 0 && help.activeParameter === 1
  );
  assert.ok(signatureHelp.signatures.length > 0, "signature help should be available across files");
  assert.equal(signatureHelp.activeParameter, 1);
  assert.equal(signatureHelp.signatures[0]?.label, "FormatMessage(ByVal value As String, ByVal count As Long) As String");
  assert.equal(getSignatureDocumentation(signatureHelp.signatures[0]?.documentation), "FormatterApi モジュール");
  assert.ok(
    getSignatureDocumentation(signatureHelp.signatures[0]?.parameters[1]?.documentation).includes("現在の引数型: Long"),
    "signature help should include inferred argument type information"
  );

  const builtInSignatureDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "BuiltInMemberSignature.bas"));
  await vscode.window.showTextDocument(builtInSignatureDocument);

  const builtInSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Sum(1, 2"),
    (help) => help.signatures.length > 0
  );
  const builtInChainedSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Application.WorksheetFunction.Sum("),
    (help) => help.signatures.length > 0
  );
  const builtInPowerSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Application.WorksheetFunction.Power("),
    (help) => help.signatures.length > 0
  );
  const builtInAverageSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Average("),
    (help) => help.signatures.length > 0
  );
  const builtInMaxSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Max("),
    (help) => help.signatures.length > 0
  );
  const builtInMinSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Min("),
    (help) => help.signatures.length > 0
  );
  const builtInEdateSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.EDate("),
    (help) => help.signatures.length > 0
  );
  const builtInEomonthSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.EoMonth("),
    (help) => help.signatures.length > 0
  );
  const builtInFindSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Find("),
    (help) => help.signatures.length > 0
  );
  const builtInSearchSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Search("),
    (help) => help.signatures.length > 0
  );
  const builtInAndSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.And("),
    (help) => help.signatures.length > 0
  );
  const builtInOrSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Or("),
    (help) => help.signatures.length > 0
  );
  const builtInXorSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Xor("),
    (help) => help.signatures.length > 0
  );
  const builtInCountASignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.CountA("),
    (help) => help.signatures.length > 0
  );
  const builtInCountBlankSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.CountBlank("),
    (help) => help.signatures.length > 0
  );
  const builtInTextSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Text("),
    (help) => help.signatures.length > 0
  );
  const builtInVlookupSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.VLookup("),
    (help) => help.signatures.length > 0
  );
  const builtInMatchSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Match("),
    (help) => help.signatures.length > 0
  );
  const builtInIndexSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Index("),
    (help) => help.signatures.length > 0
  );
  const builtInLookupSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Lookup("),
    (help) => help.signatures.length > 0
  );
  const builtInHlookupSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.HLookup("),
    (help) => help.signatures.length > 0
  );
  const builtInChooseSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Choose("),
    (help) => help.signatures.length > 0
  );
  const builtInTransposeSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "WorksheetFunction.Transpose("),
    (help) => help.signatures.length > 0
  );
  const builtInAddressSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "ActiveCell.Address("),
    (help) => help.signatures.length > 0
  );
  const builtInChainedAddressSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Application.ActiveCell.Address("),
    (help) => help.signatures.length > 0
  );
  const builtInAddressLocalSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Cells.AddressLocal("),
    (help) => help.signatures.length > 0
  );
  const builtInWorksheetEvaluateSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Worksheets(1).Evaluate("),
    (help) => help.signatures.length > 0
  );
  const builtInWorksheetStringEvaluateSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, 'Worksheets("A(1)").Evaluate('),
    (help) => help.signatures.length > 0
  );
  const builtInGroupedWorksheetEvaluateSignatureSuppressed = await waitForNoSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, 'Worksheets(Array("Sheet1", "Sheet2")).Evaluate(')
  );
  const builtInWorksheetSaveAsSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Worksheets(1).SaveAs("),
    (help) => help.signatures.length > 0
  );
  const builtInWorksheetExpressionSaveAsSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Worksheets(i + 1).SaveAs("),
    (help) => help.signatures.length > 0
  );
  const builtInWorksheetExportSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "ActiveWorkbook.Worksheets(1).ExportAsFixedFormat("),
    (help) => help.signatures.length > 0
  );
  const builtInWorksheetFunctionExportSignatureSuppressed = await waitForNoSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "ActiveWorkbook.Worksheets(GetIndex()).ExportAsFixedFormat(")
  );
  const builtInWorkbookSaveAsSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "ThisWorkbook.SaveAs("),
    (help) => help.signatures.length > 0
  );
  const builtInSheet1EvaluateSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Sheet1.Evaluate("),
    (help) => help.signatures.length > 0
  );
  const builtInSheet1SaveAsSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Sheet1.SaveAs("),
    (help) => help.signatures.length > 0
  );
  const builtInChart1SetSourceDataSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Chart1.SetSourceData("),
    (help) => help.signatures.length > 0
  );
  const builtInWorkbookCloseSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "ActiveWorkbook.Close("),
    (help) => help.signatures.length > 0
  );
  const builtInWorkbookExportSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "ActiveWorkbook.ExportAsFixedFormat("),
    (help) => help.signatures.length > 0
  );
  const builtInExtractedZeroArgSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Application.CalculateFull("),
    (help) => help.signatures.length > 0
  );
  const builtInFallbackSignatureHelp = await waitForSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Application.OnTime("),
    (help) => help.signatures.length > 0
  );
  const builtInPropertyFallbackSuppressed = await waitForNoSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Application.WorksheetFunction(")
  );
  const builtInEventFallbackSuppressed = await waitForNoSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Application.AfterCalculate(")
  );
  const builtInPropertyFallbackSuppressed2 = await waitForNoSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Application.ActiveCell(")
  );
  const builtInEventFallbackSuppressed2 = await waitForNoSignatureHelp(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Application.NewWorkbook(")
  );
  const builtInHover = await waitForHover(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Debug.Print Application.Calcu"),
    (hovers) => hovers.length > 0
  );
  const builtInWorkbookHover = await waitForHover(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Debug.Print ThisWorkbook.Save"),
    (hovers) => hovers.length > 0
  );
  const builtInChartHover = await waitForHover(
    builtInSignatureDocument,
    findPositionAfterToken(builtInSignatureDocument, "Debug.Print Chart1.ChartA"),
    (hovers) => hovers.length > 0
  );
  const builtInHoverText = getHoverContentsText(builtInHover[0]);
  const builtInWorkbookHoverText = getHoverContentsText(builtInWorkbookHover[0]);
  const builtInChartHoverText = getHoverContentsText(builtInChartHover[0]);

  assert.equal(
    builtInSignatureHelp.signatures[0]?.label,
    "Sum(Arg1, Arg2, Arg3, ..., Arg30) As Double",
    "built-in member signature should be available for WorksheetFunction.Sum"
  );
  assert.equal(
    builtInSignatureHelp.signatures[0]?.parameters.length,
    30,
    "built-in member signature should include expanded argument metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes("必須引数"),
    "built-in member first argument should be required"
  );
  assert.ok(
    getSignatureDocumentation(builtInSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("想定型: Variant"),
    "built-in member parameter should include expected type"
  );
  assert.ok(
    getSignatureDocumentation(builtInSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("省略可能"),
    "built-in member second argument should be optional"
  );
  assert.ok(
    getSignatureDocumentation(builtInSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("現在の引数型: Long"),
    "built-in member parameter should include inferred argument type information"
  );
  assert.equal(
    builtInChainedSignatureHelp.signatures[0]?.label,
    "Sum(Arg1, Arg2, Arg3, ..., Arg30) As Double",
    "built-in member signature should resolve through Application.WorksheetFunction"
  );
  assert.ok(
    builtInPowerSignatureHelp.signatures[0]?.label.includes("Power("),
    "built-in member signature should be available for WorksheetFunction.Power"
  );
  assert.equal(
    builtInPowerSignatureHelp.signatures[0]?.parameters.length,
    2,
    "built-in member signature should expose fixed parameter metadata"
  );
  assert.equal(
    builtInAverageSignatureHelp.signatures[0]?.label,
    "Average(Arg1, Arg2, Arg3, ..., Arg30) As Double",
    "built-in member signature should expand WorksheetFunction.Average with variadic metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInAverageSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("省略可能"),
    "built-in member average second argument should be optional"
  );
  assertVariadicWorksheetFunctionSignature(
    builtInMaxSignatureHelp,
    "Max(Arg1, Arg2, Arg3, ..., Arg30) As Double",
    "Max"
  );
  assertVariadicWorksheetFunctionSignature(
    builtInMinSignatureHelp,
    "Min(Arg1, Arg2, Arg3, ..., Arg30) As Double",
    "Min"
  );
  assert.equal(
    builtInEdateSignatureHelp.signatures[0]?.label,
    "EDate(Arg1, Arg2) As Double",
    "built-in member signature should be available for WorksheetFunction.EDate"
  );
  assert.equal(
    builtInEdateSignatureHelp.signatures[0]?.parameters.length,
    2,
    "built-in member EDate signature should keep fixed parameter metadata"
  );
  assert.equal(
    builtInEomonthSignatureHelp.signatures[0]?.label,
    "EoMonth(Arg1, Arg2) As Double",
    "built-in member signature should be available for WorksheetFunction.EoMonth"
  );
  assert.equal(
    builtInEomonthSignatureHelp.signatures[0]?.parameters.length,
    2,
    "built-in member EoMonth signature should keep fixed parameter metadata"
  );
  assert.equal(
    builtInFindSignatureHelp.signatures[0]?.label,
    "Find(Arg1, Arg2, Arg3) As Double",
    "built-in member signature should be available for WorksheetFunction.Find"
  );
  assert.equal(
    builtInFindSignatureHelp.signatures[0]?.parameters.length,
    3,
    "built-in member Find signature should expose optional third argument"
  );
  assert.ok(
    getSignatureDocumentation(builtInFindSignatureHelp.signatures[0]?.parameters[2]?.documentation).includes("省略可能"),
    "built-in member Find third argument should be optional"
  );
  assert.equal(
    builtInSearchSignatureHelp.signatures[0]?.label,
    "Search(Arg1, Arg2, Arg3) As Double",
    "built-in member signature should be available for WorksheetFunction.Search"
  );
  assert.equal(
    builtInSearchSignatureHelp.signatures[0]?.parameters.length,
    3,
    "built-in member Search signature should expose optional third argument"
  );
  assert.ok(
    getSignatureDocumentation(builtInSearchSignatureHelp.signatures[0]?.parameters[2]?.documentation).includes("省略可能"),
    "built-in member Search third argument should be optional"
  );
  assert.equal(
    builtInAndSignatureHelp.signatures[0]?.label,
    "And(Arg1, Arg2, Arg3, ..., Arg30) As Boolean",
    "built-in member signature should be available for WorksheetFunction.And"
  );
  assert.equal(
    builtInAndSignatureHelp.signatures[0]?.parameters.length,
    30,
    "built-in member And signature should expose variadic parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInAndSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("省略可能"),
    "built-in member And second argument should be optional"
  );
  assert.equal(
    builtInOrSignatureHelp.signatures[0]?.label,
    "Or(Arg1, Arg2, Arg3, ..., Arg30) As Boolean",
    "built-in member signature should be available for WorksheetFunction.Or"
  );
  assert.equal(
    builtInOrSignatureHelp.signatures[0]?.parameters.length,
    30,
    "built-in member Or signature should expose variadic parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInOrSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("想定型: Variant"),
    "built-in member Or second argument should include expected type"
  );
  assert.ok(
    getSignatureDocumentation(builtInOrSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("省略可能"),
    "built-in member Or second argument should be optional"
  );
  assert.equal(
    builtInXorSignatureHelp.signatures[0]?.label,
    "Xor(Arg1, Arg2, Arg3, ..., Arg30) As Boolean",
    "built-in member signature should be available for WorksheetFunction.Xor"
  );
  assert.equal(
    builtInXorSignatureHelp.signatures[0]?.parameters.length,
    30,
    "built-in member Xor signature should expose variadic parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInXorSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("想定型: Variant"),
    "built-in member Xor second argument should include expected type"
  );
  assert.ok(
    getSignatureDocumentation(builtInXorSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("省略可能"),
    "built-in member Xor second argument should be optional"
  );
  assert.equal(
    builtInCountASignatureHelp.signatures[0]?.label,
    "CountA(Arg1, Arg2, Arg3, ..., Arg30) As Double",
    "built-in member signature should be available for WorksheetFunction.CountA"
  );
  assert.equal(
    builtInCountASignatureHelp.signatures[0]?.parameters.length,
    30,
    "built-in member CountA signature should expose variadic parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInCountASignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("省略可能"),
    "built-in member CountA second argument should be optional"
  );
  assert.equal(
    builtInCountBlankSignatureHelp.signatures[0]?.label,
    "CountBlank(Arg1) As Double",
    "built-in member signature should be available for WorksheetFunction.CountBlank"
  );
  assert.equal(
    builtInCountBlankSignatureHelp.signatures[0]?.parameters.length,
    1,
    "built-in member CountBlank signature should expose single-argument metadata"
  );
  assert.equal(
    builtInTextSignatureHelp.signatures[0]?.label,
    "Text(Arg1, Arg2) As String",
    "built-in member signature should be available for WorksheetFunction.Text"
  );
  assert.equal(
    builtInTextSignatureHelp.signatures[0]?.parameters.length,
    2,
    "built-in member Text signature should keep fixed parameter metadata"
  );
  assert.equal(
    builtInVlookupSignatureHelp.signatures[0]?.label,
    "VLookup(Arg1, Arg2, Arg3, Arg4) As Variant",
    "built-in member signature should be available for WorksheetFunction.VLookup"
  );
  assert.ok(
    getSignatureDocumentation(builtInVlookupSignatureHelp.signatures[0]?.parameters[3]?.documentation).includes("省略可能"),
    "built-in member VLookup fourth argument should be optional"
  );
  assert.equal(
    builtInMatchSignatureHelp.signatures[0]?.label,
    "Match(Arg1, Arg2, Arg3) As Double",
    "built-in member signature should be available for WorksheetFunction.Match"
  );
  assert.ok(
    getSignatureDocumentation(builtInMatchSignatureHelp.signatures[0]?.parameters[2]?.documentation).includes("省略可能"),
    "built-in member Match third argument should be optional"
  );
  assert.equal(
    builtInIndexSignatureHelp.signatures[0]?.label,
    "Index(Arg1, Arg2, Arg3, Arg4) As Variant",
    "built-in member signature should be available for WorksheetFunction.Index"
  );
  assert.ok(
    getSignatureDocumentation(builtInIndexSignatureHelp.signatures[0]?.parameters[2]?.documentation).includes("省略可能"),
    "built-in member Index third argument should be optional"
  );
  assert.ok(
    getSignatureDocumentation(builtInIndexSignatureHelp.signatures[0]?.parameters[3]?.documentation).includes("省略可能"),
    "built-in member Index fourth argument should be optional"
  );
  assert.equal(
    builtInLookupSignatureHelp.signatures[0]?.label,
    "Lookup(Arg1, Arg2, Arg3) As Variant",
    "built-in member signature should be available for WorksheetFunction.Lookup"
  );
  assert.ok(
    getSignatureDocumentation(builtInLookupSignatureHelp.signatures[0]?.parameters[2]?.documentation).includes("省略可能"),
    "built-in member Lookup third argument should be optional"
  );
  assert.equal(
    builtInHlookupSignatureHelp.signatures[0]?.label,
    "HLookup(Arg1, Arg2, Arg3, Arg4) As Variant",
    "built-in member signature should be available for WorksheetFunction.HLookup"
  );
  assert.ok(
    getSignatureDocumentation(builtInHlookupSignatureHelp.signatures[0]?.parameters[3]?.documentation).includes("省略可能"),
    "built-in member HLookup fourth argument should be optional"
  );
  assert.equal(
    builtInChooseSignatureHelp.signatures[0]?.label,
    "Choose(Arg1, Arg2, Arg3, ..., Arg30) As Variant",
    "built-in member signature should be available for WorksheetFunction.Choose"
  );
  assert.equal(
    builtInChooseSignatureHelp.signatures[0]?.parameters.length,
    30,
    "built-in member Choose signature should expose expanded parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInChooseSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("想定型: Variant"),
    "built-in member Choose second argument should include expected type"
  );
  assert.equal(
    getSignatureDocumentation(builtInChooseSignatureHelp.signatures[0]?.parameters[1]?.documentation).includes("省略可能"),
    false,
    "built-in member Choose second argument should stay required"
  );
  assert.equal(
    getSignatureDocumentation(builtInChooseSignatureHelp.signatures[0]?.parameters[29]?.documentation).includes("省略可能"),
    false,
    "built-in member Choose last argument should stay required"
  );
  assert.equal(
    builtInTransposeSignatureHelp.signatures[0]?.label,
    "Transpose(Arg1) As Variant",
    "built-in member signature should be available for WorksheetFunction.Transpose"
  );
  assert.equal(
    builtInTransposeSignatureHelp.signatures[0]?.parameters.length,
    1,
    "built-in member Transpose signature should keep single-argument metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInTransposeSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes("必須引数"),
    "built-in member Transpose argument should stay required"
  );
  assert.equal(
    builtInAddressSignatureHelp.signatures[0]?.label,
    "Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo) As String",
    "built-in member signature should be available for ActiveCell.Address"
  );
  assert.equal(
    builtInAddressSignatureHelp.signatures[0]?.parameters.length,
    5,
    "built-in Address signature should expose fixed parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInAddressSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes("省略可能"),
    "built-in Address first argument should be optional"
  );
  assert.ok(
    getSignatureDocumentation(builtInAddressSignatureHelp.signatures[0]?.parameters[2]?.documentation).includes(
      "想定型: XlReferenceStyle"
    ),
    "built-in Address reference style should include the expected enum type"
  );
  assert.ok(
    getSignatureDocumentation(builtInAddressSignatureHelp.signatures[0]?.parameters[4]?.documentation).includes("省略可能"),
    "built-in Address RelativeTo argument should be optional"
  );
  assert.equal(
    builtInChainedAddressSignatureHelp.signatures[0]?.label,
    "Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo) As String",
    "built-in member signature should resolve through Application.ActiveCell"
  );
  assert.equal(
    builtInChainedAddressSignatureHelp.signatures[0]?.parameters.length,
    5,
    "Application.ActiveCell.Address should preserve Range.Address parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInChainedAddressSignatureHelp.signatures[0]?.parameters[2]?.documentation).includes(
      "想定型: XlReferenceStyle"
    ),
    "Application.ActiveCell.Address should keep reference style type metadata"
  );
  assert.equal(
    builtInAddressLocalSignatureHelp.signatures[0]?.label,
    "AddressLocal(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo) As String",
    "built-in member signature should be available for Cells.AddressLocal"
  );
  assert.equal(
    builtInAddressLocalSignatureHelp.signatures[0]?.parameters.length,
    5,
    "built-in AddressLocal signature should expose fixed parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInAddressLocalSignatureHelp.signatures[0]?.parameters[2]?.documentation).includes(
      "想定型: XlReferenceStyle"
    ),
    "built-in AddressLocal reference style should include the expected enum type"
  );
  assert.ok(
    getSignatureDocumentation(builtInAddressLocalSignatureHelp.signatures[0]?.parameters[4]?.documentation).includes(
      "省略可能"
    ),
    "built-in AddressLocal RelativeTo argument should be optional"
  );
  assert.equal(
    builtInWorksheetEvaluateSignatureHelp.signatures[0]?.label,
    "Evaluate(Name) As Variant",
    "built-in member signature should be available for Worksheets(1).Evaluate"
  );
  assert.equal(
    builtInWorksheetEvaluateSignatureHelp.signatures[0]?.parameters.length,
    1,
    "built-in Worksheet.Evaluate signature should expose one required parameter"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorksheetEvaluateSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "必須引数"
    ),
    "built-in Worksheet.Evaluate argument should stay required"
  );
  assert.equal(
    builtInWorksheetStringEvaluateSignatureHelp.signatures[0]?.label,
    "Evaluate(Name) As Variant",
    "built-in member signature should be available for Worksheets(\"A(1)\").Evaluate"
  );
  assert.ok(
    getSignatureDocumentation(
      builtInWorksheetStringEvaluateSignatureHelp.signatures[0]?.parameters[0]?.documentation
    ).includes("必須引数"),
    "built-in Worksheet.Evaluate should stay available for string index access"
  );
  assert.equal(
    builtInGroupedWorksheetEvaluateSignatureSuppressed,
    true,
    "grouped Worksheets selector should not expose Worksheet.Evaluate signature help"
  );
  assert.equal(
    builtInWorksheetSaveAsSignatureHelp.signatures[0]?.label,
    "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local)",
    "built-in member signature should be available for Worksheets(1).SaveAs"
  );
  assert.equal(
    builtInWorksheetSaveAsSignatureHelp.signatures[0]?.parameters.length,
    10,
    "built-in Worksheet.SaveAs signature should expose worksheet-specific parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorksheetSaveAsSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "必須引数"
    ),
    "built-in Worksheet.SaveAs first argument should stay required"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorksheetSaveAsSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "想定型: String"
    ),
    "built-in Worksheet.SaveAs first argument should include string type metadata"
  );
  assert.equal(
    builtInWorksheetExpressionSaveAsSignatureHelp.signatures[0]?.label,
    "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local)",
    "built-in member signature should be available for Worksheets(i + 1).SaveAs"
  );
  assert.ok(
    getSignatureDocumentation(
      builtInWorksheetExpressionSaveAsSignatureHelp.signatures[0]?.parameters[0]?.documentation
    ).includes("想定型: String"),
    "built-in Worksheet.SaveAs should stay available for expression index access"
  );
  assert.equal(
    builtInWorksheetExportSignatureHelp.signatures[0]?.label,
    "ExportAsFixedFormat(Type, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr)",
    "built-in member signature should be available for ActiveWorkbook.Worksheets(1).ExportAsFixedFormat"
  );
  assert.equal(
    builtInWorksheetExportSignatureHelp.signatures[0]?.parameters.length,
    9,
    "built-in Worksheet.ExportAsFixedFormat signature should expose fixed parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorksheetExportSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "必須引数"
    ),
    "built-in Worksheet.ExportAsFixedFormat first argument should stay required"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorksheetExportSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "想定型: XlFixedFormatType"
    ),
    "built-in Worksheet.ExportAsFixedFormat first argument should include enum type metadata"
  );
  assert.equal(
    builtInWorksheetFunctionExportSignatureSuppressed,
    true,
    "function-based Worksheets selector should not expose Worksheet.ExportAsFixedFormat signature help"
  );
  assert.equal(
    builtInWorkbookSaveAsSignatureHelp.signatures[0]?.label,
    "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)",
    "built-in member signature should be available for ThisWorkbook.SaveAs"
  );
  assert.equal(
    builtInSheet1EvaluateSignatureHelp.signatures[0]?.label,
    "Evaluate(Name) As Variant",
    "worksheet document root should expose Worksheet.Evaluate signature"
  );
  assert.equal(
    builtInSheet1SaveAsSignatureHelp.signatures[0]?.label,
    "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local)",
    "worksheet document root should expose Worksheet.SaveAs signature"
  );
  assert.equal(
    builtInChart1SetSourceDataSignatureHelp.signatures[0]?.label,
    "Chart.SetSourceData()",
    "chart document root should expose Chart.SetSourceData signature"
  );
  assert.ok(
    getSignatureDocumentation(builtInChart1SetSourceDataSignatureHelp.signatures[0]?.documentation).includes(
      "excel.chart.setsourcedata"
    ),
    "Chart1.SetSourceData signature should include the Chart Learn URL"
  );
  assert.equal(
    builtInWorkbookSaveAsSignatureHelp.signatures[0]?.parameters.length,
    12,
    "built-in Workbook.SaveAs signature should expose all parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorkbookSaveAsSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "省略可能"
    ),
    "built-in Workbook.SaveAs first argument should stay optional"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorkbookSaveAsSignatureHelp.signatures[0]?.parameters[6]?.documentation).includes(
      "想定型: XlSaveAsAccessMode"
    ),
    "built-in Workbook.SaveAs AccessMode should include enum type metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorkbookSaveAsSignatureHelp.signatures[0]?.parameters[7]?.documentation).includes(
      "想定型: XlSaveConflictResolution"
    ),
    "built-in Workbook.SaveAs ConflictResolution should include enum type metadata"
  );
  assert.equal(
    builtInWorkbookCloseSignatureHelp.signatures[0]?.label,
    "Close(SaveChanges, FileName, RouteWorkbook)",
    "built-in member signature should be available for ActiveWorkbook.Close"
  );
  assert.equal(
    builtInWorkbookCloseSignatureHelp.signatures[0]?.parameters.length,
    3,
    "built-in Workbook.Close signature should expose optional parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorkbookCloseSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "省略可能"
    ),
    "built-in Workbook.Close first argument should stay optional"
  );
  assert.equal(
    builtInWorkbookExportSignatureHelp.signatures[0]?.label,
    "ExportAsFixedFormat(Type, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr)",
    "built-in member signature should be available for ActiveWorkbook.ExportAsFixedFormat"
  );
  assert.equal(
    builtInWorkbookExportSignatureHelp.signatures[0]?.parameters.length,
    9,
    "built-in Workbook.ExportAsFixedFormat signature should expose fixed parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorkbookExportSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "必須引数"
    ),
    "built-in Workbook.ExportAsFixedFormat first argument should stay required"
  );
  assert.ok(
    getSignatureDocumentation(builtInWorkbookExportSignatureHelp.signatures[0]?.parameters[0]?.documentation).includes(
      "想定型: XlFixedFormatType"
    ),
    "built-in Workbook.ExportAsFixedFormat first argument should include enum type metadata"
  );
  assert.equal(
    builtInExtractedZeroArgSignatureHelp.signatures[0]?.label,
    "CalculateFull()",
    "built-in zero-arg allow-listed signature should use extracted short label"
  );
  assert.equal(
    builtInExtractedZeroArgSignatureHelp.signatures[0]?.parameters.length,
    0,
    "built-in zero-arg allow-listed signature should not fabricate parameters"
  );
  assert.equal(
    builtInFallbackSignatureHelp.signatures[0]?.label,
    "Application.OnTime()",
    "built-in callable without signature metadata should expose fallback label"
  );
  assert.equal(
    builtInFallbackSignatureHelp.signatures[0]?.parameters.length,
    0,
    "built-in callable fallback should not fabricate parameter metadata"
  );
  assert.ok(
    getSignatureDocumentation(builtInFallbackSignatureHelp.signatures[0]?.documentation).includes("excel.application.ontime"),
    "built-in callable fallback should keep learn URL in documentation"
  );
  assert.equal(
    builtInPropertyFallbackSuppressed,
    true,
    "built-in property call should not fabricate fallback signature help"
  );
  assert.equal(
    builtInEventFallbackSuppressed,
    true,
    "built-in event call should not fabricate fallback signature help"
  );
  assert.equal(
    builtInPropertyFallbackSuppressed2,
    true,
    "built-in property call should stay suppressed for Application.ActiveCell()"
  );
  assert.equal(
    builtInEventFallbackSuppressed2,
    true,
    "built-in event call should stay suppressed for Application.NewWorkbook()"
  );
  assert.ok(
    builtInHoverText.includes("Calculate()"),
    "built-in hover should include callable signature label"
  );
  assert.ok(
    builtInHoverText.includes("Microsoft Learn"),
    "built-in hover should include source link"
  );
  assert.ok(
    builtInWorkbookHoverText.includes(
      "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)"
    ),
    "ThisWorkbook hover should resolve through Workbook members"
  );
  assert.ok(
    builtInWorkbookHoverText.includes("excel.workbook.saveas"),
    "ThisWorkbook hover should include the Workbook.SaveAs Learn URL"
  );
  assert.ok(
    builtInChartHoverText.includes("Chart.ChartArea"),
    "Chart1 hover should resolve through Chart members"
  );
  assert.ok(
    builtInChartHoverText.includes("excel.chart.chartarea"),
    "Chart1 hover should include the Chart.ChartArea Learn URL"
  );

  const renameDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "RenameLocal.bas"));
  await vscode.window.showTextDocument(renameDocument);

  const renameEdit = await waitForRename(
    renameDocument,
    new vscode.Position(6, 6),
    "currentCount",
    (edit) => (edit.get(renameDocument.uri)?.length ?? 0) === 4
  );
  const renameEntries = renameEdit.get(renameDocument.uri) ?? [];

  assert.equal(renameEntries.length, 4, "rename should update declaration and local references only");
  assert.ok(renameEntries.every((edit) => edit.newText === "currentCount"));
  assert.ok(renameEntries.every((edit) => edit.range.start.line >= 4 && edit.range.start.line <= 8));

  const semanticDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "SemanticTokens.bas"));
  await vscode.window.showTextDocument(semanticDocument);

  const semanticLegend = await waitForSemanticTokensLegend(
    semanticDocument,
    (legend) =>
      legend.tokenTypes.includes("variable") &&
      legend.tokenTypes.includes("parameter") &&
      legend.tokenTypes.includes("function") &&
      legend.tokenTypes.includes("type")
  );
  const semanticTokens = await waitForSemanticTokens(
    semanticDocument,
    (tokens) => tokens.data.length > 0
  );

  assert.ok(semanticLegend.tokenTypes.includes("variable"), "semantic token legend should include variable");
  assert.ok(semanticLegend.tokenTypes.includes("parameter"), "semantic token legend should include parameter");
  assert.ok(semanticLegend.tokenTypes.includes("function"), "semantic token legend should include function");
  assert.ok(semanticLegend.tokenTypes.includes("type"), "semantic token legend should include type");
  assert.ok(semanticTokens.data.length > 0, "semantic tokens should be available");

  const builtInSemanticDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "BuiltInSemantic.bas"));
  await vscode.window.showTextDocument(builtInSemanticDocument);

  const builtInSemanticLegend = await waitForSemanticTokensLegend(
    builtInSemanticDocument,
    (legend) => legend.tokenTypes.includes("keyword")
  );
  const builtInSemanticTokens = await waitForSemanticTokens(
    builtInSemanticDocument,
    (tokens) => tokens.data.length > 0
  );
  const decodedBuiltInSemanticTokens = decodeSemanticTokens(builtInSemanticTokens, builtInSemanticLegend);

  assert.ok(builtInSemanticLegend.tokenTypes.includes("keyword"), "semantic token legend should include keyword");
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 4, "Beep", {
    modifiers: [],
    type: "keyword"
  });
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 5, "MsgBox", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 5, "xlAll", {
    modifiers: ["readonly"],
    type: "variable"
  });
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 6, "Name", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 7, "WorksheetFunction", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 7, "Sum", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 8, "Evaluate", {
    modifiers: [],
    type: "function"
  });
  assertNoDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 9, "Evaluate");
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 10, "ThisWorkbook", {
    modifiers: [],
    type: "variable"
  });
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 10, "SaveAs", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 11, "Evaluate", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 12, "SetSourceData", {
    modifiers: [],
    type: "function"
  });
  assertDecodedSemanticToken(builtInSemanticDocument.getText(), decodedBuiltInSemanticTokens, 13, "Address", {
    modifiers: [],
    type: "variable"
  });

  const formatDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "FormatDocument.bas"));
  await vscode.window.showTextDocument(formatDocument);

  const formattedText = await waitForFormattedDocument(
    formatDocument,
    `Attribute VB_Name = "FormatDocument"
Option Explicit

Public Sub Demo()
    If True Then
        Debug.Print "ready"
    Else
        Select Case 1
            Case Else
                Debug.Print "fallback"
        End Select
    End If
End Sub`
  );

  assert.equal(normalizeText(formattedText), normalizeText(`Attribute VB_Name = "FormatDocument"
Option Explicit

Public Sub Demo()
    If True Then
        Debug.Print "ready"
    Else
        Select Case 1
            Case Else
                Debug.Print "fallback"
        End Select
    End If
End Sub`));

  const continuationDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "ContinuationFormatting.bas"));
  await vscode.window.showTextDocument(continuationDocument);

  const formattedContinuationText = await waitForFormattedDocument(
    continuationDocument,
    `Attribute VB_Name = "ContinuationFormatting"
Option Explicit

Public Sub Demo()
    Dim message As String
    message = _
        "prefix" & _
        "suffix"

    Debug.Print JoinValues( _
        message, _
        "tail" _
    )

    message = CreateBuilder() _
        .WithName(message) _
        .WithSuffix("!")
End Sub`
  );

  assert.equal(normalizeText(formattedContinuationText), normalizeText(`Attribute VB_Name = "ContinuationFormatting"
Option Explicit

Public Sub Demo()
    Dim message As String
    message = _
        "prefix" & _
        "suffix"

    Debug.Print JoinValues( _
        message, _
        "tail" _
    )

    message = CreateBuilder() _
        .WithName(message) _
        .WithSuffix("!")
End Sub`));

  const blockLayoutDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "BlockLayoutFormatting.bas"));
  await vscode.window.showTextDocument(blockLayoutDocument);

  const formattedBlockLayoutText = await waitForFormattedDocument(
    blockLayoutDocument,
    `Attribute VB_Name = "BlockLayoutFormatting"
Option Explicit

Public Sub Demo()
    Dim value As Long: value = 0
    If value = 0 Then
        Debug.Print "zero"
    ElseIf value = 1 Then
        Debug.Print "one"
    Else
        Debug.Print "other"
    End If
    Select Case value
        Case 0
            Debug.Print "case zero"
        Case Else
            With Application
                .StatusBar = "fallback"
            End With
    End Select
    #If VBA7 Then
        value = value + 1
    #Else
        value = value - 1
    #End If
End Sub`
  );

  assert.equal(normalizeText(formattedBlockLayoutText), normalizeText(`Attribute VB_Name = "BlockLayoutFormatting"
Option Explicit

Public Sub Demo()
    Dim value As Long: value = 0
    If value = 0 Then
        Debug.Print "zero"
    ElseIf value = 1 Then
        Debug.Print "one"
    Else
        Debug.Print "other"
    End If
    Select Case value
        Case 0
            Debug.Print "case zero"
        Case Else
            With Application
                .StatusBar = "fallback"
            End With
    End Select
    #If VBA7 Then
        value = value + 1
    #Else
        value = value - 1
    #End If
End Sub`));

  const declarationAlignmentDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "DeclarationAlignment.bas"));
  await vscode.window.showTextDocument(declarationAlignmentDocument);

  const formattedDeclarationAlignmentText = await waitForFormattedDocument(
    declarationAlignmentDocument,
    `Attribute VB_Name = "DeclarationAlignment"
Option Explicit

Private Declare PtrSafe Function GetActiveWindow  Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr

Public Sub Demo()
    Dim title   As String
    Dim count   As Long
    Dim enabled As Boolean

    Const DefaultTitle As String  = "Ready"
    Const RetryCount   As Long    = 3
    Const IsEnabled    As Boolean = True

    Debug.Print title, count, enabled
End Sub`
  );

  assert.equal(normalizeText(formattedDeclarationAlignmentText), normalizeText(`Attribute VB_Name = "DeclarationAlignment"
Option Explicit

Private Declare PtrSafe Function GetActiveWindow  Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr

Public Sub Demo()
    Dim title   As String
    Dim count   As Long
    Dim enabled As Boolean

    Const DefaultTitle As String  = "Ready"
    Const RetryCount   As Long    = 3
    Const IsEnabled    As Boolean = True

    Debug.Print title, count, enabled
End Sub`));

  const commentFormattingDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "CommentFormatting.bas"));
  await vscode.window.showTextDocument(commentFormattingDocument);

  const formattedCommentText = await waitForFormattedDocument(
    commentFormattingDocument,
    `Attribute VB_Name = "CommentFormatting"
Option Explicit

Public Sub Demo()
    ' leading comment
    Dim value As Long ' counter
    If True Then ' true branch
        ' inner comment
        value = 1 ' updated
        Rem status
        #If VBA7 Then ' requires vba7
            ' conditional comment
        #Else ' fallback path
            Rem fallback comment
        #End If
    End If
End Sub`
  );

  assert.equal(normalizeText(formattedCommentText), normalizeText(`Attribute VB_Name = "CommentFormatting"
Option Explicit

Public Sub Demo()
    ' leading comment
    Dim value As Long ' counter
    If True Then ' true branch
        ' inner comment
        value = 1 ' updated
        Rem status
        #If VBA7 Then ' requires vba7
            ' conditional comment
        #Else ' fallback path
            Rem fallback comment
        #End If
    End If
End Sub`));

  const missingOptionDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "MissingOptionExplicit.bas"));
  await vscode.window.showTextDocument(missingOptionDocument);

  const optionActions = await waitForCodeActions(
    missingOptionDocument,
    new vscode.Range(new vscode.Position(0, 0), new vscode.Position(0, 0)),
    (actions) => hasCodeAction(actions, "Option Explicit を追加")
  );
  const optionAction = getCodeAction(optionActions, "Option Explicit を追加");

  assert.ok(optionAction?.edit, "missing Option Explicit should expose a quick fix");
  assert.equal(await vscode.workspace.applyEdit(optionAction.edit), true);
  assert.equal(
    normalizeText(missingOptionDocument.getText()),
    normalizeText(`Attribute VB_Name = "MissingOptionExplicit"
Option Compare Text
Option Explicit

Public Sub Demo()
    Debug.Print "ready"
End Sub`)
  );

  const missingFormDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "MissingOptionExplicit.frm"));
  await vscode.window.showTextDocument(missingFormDocument);

  const formOptionActions = await waitForCodeActions(
    missingFormDocument,
    new vscode.Range(new vscode.Position(0, 0), new vscode.Position(0, 0)),
    (actions) => hasCodeAction(actions, "Option Explicit を追加")
  );
  const formOptionAction = getCodeAction(formOptionActions, "Option Explicit を追加");

  assert.ok(formOptionAction?.edit, "form modules should expose a quick fix for missing Option Explicit");
  assert.equal(await vscode.workspace.applyEdit(formOptionAction.edit), true);
  assert.equal(
    normalizeText(missingFormDocument.getText()),
    normalizeText(`VERSION 5.00
Begin VB.Form MissingOptionExplicit
   Caption = "MissingOptionExplicit"
End
Attribute VB_Name = "MissingOptionExplicit"
Option Explicit

Public Sub Demo()
    Debug.Print "ready"
End Sub`)
  );

  const snippetDocument = await vscode.workspace.openTextDocument(path.resolve(fixturesPath, "SnippetCompletions.bas"));
  await vscode.window.showTextDocument(snippetDocument);

  const subSnippetItems = await waitForCompletions(
    snippetDocument,
    new vscode.Position(3, 3),
    (items) => hasSnippetCompletion(items, "sub")
  );
  const selectSnippetItems = await waitForCompletions(
    snippetDocument,
    new vscode.Position(4, 6),
    (items) => hasSnippetCompletion(items, "select")
  );

  assert.ok(
    hasSnippetCompletion(subSnippetItems, "sub"),
    "snippet completion should include the Sub Procedure template"
  );
  assert.ok(
    hasSnippetCompletion(selectSnippetItems, "select"),
    "snippet completion should include the Select Case template"
  );

  const commands = await vscode.commands.getCommands(true);
  assert.equal(commands.includes("vba.extract"), false);
  assert.equal(commands.includes("vba.combine"), false);
}

async function waitForSymbols(document: vscode.TextDocument): Promise<readonly vscode.DocumentSymbol[]> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const symbols = await vscode.commands.executeCommand<readonly vscode.DocumentSymbol[]>(
      "vscode.executeDocumentSymbolProvider",
      document.uri
    );

    if (symbols && symbols.length > 0) {
      return symbols;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return [];
}

async function waitForCodeActions(
  document: vscode.TextDocument,
  range: vscode.Range,
  predicate: (actions: readonly vscode.CodeAction[]) => boolean
): Promise<readonly vscode.CodeAction[]> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const actions = await vscode.commands.executeCommand<readonly (vscode.CodeAction | vscode.Command)[]>(
      "vscode.executeCodeActionProvider",
      document.uri,
      range,
      vscode.CodeActionKind.QuickFix.value
    );
    const codeActions = (actions ?? []).filter(isCodeAction);

    if (codeActions.length > 0 && predicate(codeActions)) {
      return codeActions;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return [];
}

async function waitForCompletions(
  document: vscode.TextDocument,
  position: vscode.Position,
  predicate: (items: readonly vscode.CompletionItem[]) => boolean
): Promise<readonly vscode.CompletionItem[]> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const completionList = await vscode.commands.executeCommand<vscode.CompletionList>(
      "vscode.executeCompletionItemProvider",
      document.uri,
      position
    );

    if (completionList?.items?.length && predicate(completionList.items)) {
      return completionList.items;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return [];
}

async function waitForCompletionsAtToken(
  document: vscode.TextDocument,
  token: string,
  predicate: (items: readonly vscode.CompletionItem[]) => boolean,
  occurrenceIndex = 0
): Promise<readonly vscode.CompletionItem[]> {
  return waitForCompletions(document, findPositionAfterToken(document, token, 0, occurrenceIndex), predicate);
}

async function waitForCompletionLabelStateAtToken(
  document: vscode.TextDocument,
  token: string,
  label: string,
  expectedPresent: boolean,
  occurrenceIndex = 0
): Promise<readonly vscode.CompletionItem[]> {
  return waitForCompletionsAtToken(
    document,
    token,
    (items) => hasCompletionItemLabel(items, label) === expectedPresent,
    occurrenceIndex
  );
}

async function waitForDefinitions(
  document: vscode.TextDocument,
  position: vscode.Position,
  predicate: (locations: readonly vscode.Location[]) => boolean
): Promise<readonly vscode.Location[]> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const definitions = await vscode.commands.executeCommand<readonly vscode.Location[]>(
      "vscode.executeDefinitionProvider",
      document.uri,
      position
    );

    if (definitions?.length && predicate(definitions)) {
      return definitions;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return [];
}

async function waitForReferences(
  document: vscode.TextDocument,
  position: vscode.Position,
  predicate: (locations: readonly vscode.Location[]) => boolean
): Promise<readonly vscode.Location[]> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const references = await vscode.commands.executeCommand<readonly vscode.Location[]>(
      "vscode.executeReferenceProvider",
      document.uri,
      position
    );

    if (references?.length && predicate(references)) {
      return references;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return [];
}

async function waitForSignatureHelp(
  document: vscode.TextDocument,
  position: vscode.Position,
  predicate: (help: vscode.SignatureHelp) => boolean
): Promise<vscode.SignatureHelp> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const signatureHelp = await vscode.commands.executeCommand<vscode.SignatureHelp>(
      "vscode.executeSignatureHelpProvider",
      document.uri,
      position
    );

    if (signatureHelp && predicate(signatureHelp)) {
      return signatureHelp;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return {
    activeParameter: 0,
    activeSignature: 0,
    signatures: []
  };
}

async function waitForNoSignatureHelp(
  document: vscode.TextDocument,
  position: vscode.Position
): Promise<boolean> {
  for (let attempt = 0; attempt < NEGATIVE_LANGUAGE_FEATURE_RETRY_COUNT; attempt += 1) {
    const signatureHelp = await vscode.commands.executeCommand<vscode.SignatureHelp>(
      "vscode.executeSignatureHelpProvider",
      document.uri,
      position
    );

    if (signatureHelp?.signatures.length) {
      return false;
    }

    await new Promise((resolve) => setTimeout(resolve, NEGATIVE_LANGUAGE_FEATURE_RETRY_DELAY_MS));
  }

  return true;
}

async function waitForSignatureHelpAtToken(
  document: vscode.TextDocument,
  token: string,
  predicate: (help: vscode.SignatureHelp) => boolean,
  occurrenceIndex = 0
): Promise<vscode.SignatureHelp> {
  return waitForSignatureHelp(document, findPositionAfterToken(document, token, 0, occurrenceIndex), predicate);
}

async function waitForNoSignatureHelpAtToken(
  document: vscode.TextDocument,
  token: string,
  occurrenceIndex = 0
): Promise<boolean> {
  return waitForNoSignatureHelp(document, findPositionAfterToken(document, token, 0, occurrenceIndex));
}

async function waitForHover(
  document: vscode.TextDocument,
  position: vscode.Position,
  predicate: (hovers: readonly vscode.Hover[]) => boolean
): Promise<readonly vscode.Hover[]> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const hovers = await vscode.commands.executeCommand<readonly vscode.Hover[]>(
      "vscode.executeHoverProvider",
      document.uri,
      position
    );

    if (hovers && predicate(hovers)) {
      return hovers;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return [];
}

async function waitForHoverAtToken(
  document: vscode.TextDocument,
  token: string,
  predicate: (hovers: readonly vscode.Hover[]) => boolean,
  occurrenceIndex = 0
): Promise<readonly vscode.Hover[]> {
  return waitForHover(document, findPositionAfterToken(document, token, 0, occurrenceIndex), predicate);
}

async function waitForNoHover(document: vscode.TextDocument, position: vscode.Position): Promise<boolean> {
  for (let attempt = 0; attempt < NEGATIVE_LANGUAGE_FEATURE_RETRY_COUNT; attempt += 1) {
    const hovers = await vscode.commands.executeCommand<readonly vscode.Hover[]>(
      "vscode.executeHoverProvider",
      document.uri,
      position
    );

    if (hovers?.length) {
      return false;
    }

    await new Promise((resolve) => setTimeout(resolve, NEGATIVE_LANGUAGE_FEATURE_RETRY_DELAY_MS));
  }

  return true;
}

async function waitForNoHoverAtToken(
  document: vscode.TextDocument,
  token: string,
  occurrenceIndex = 0
): Promise<boolean> {
  return waitForNoHover(document, findPositionAfterToken(document, token, 0, occurrenceIndex));
}

function loadWorkbookRootFamilyCaseTables(): WorkbookRootFamilyCaseTables {
  let currentDirectory = __dirname;

  for (let depth = 0; depth < 8; depth += 1) {
    const candidatePath = path.resolve(currentDirectory, "test-support", "workbookRootFamilyCaseTables.cjs");

    if (existsSync(candidatePath)) {
      const loaded = requireFromHere(candidatePath) as {
        workbookRootFamilyCaseTables: WorkbookRootFamilyCaseTables;
      };

      return loaded.workbookRootFamilyCaseTables;
    }

    currentDirectory = path.resolve(currentDirectory, "..");
  }

  throw new Error("workbook root family shared case spec が見つかりません");
}

function loadWorksheetControlShapeNamePathCaseTables(): WorksheetControlShapeNamePathCaseTables {
  let currentDirectory = __dirname;

  for (let depth = 0; depth < 8; depth += 1) {
    const candidatePath = path.resolve(currentDirectory, "test-support", "worksheetControlShapeNamePathCaseTables.cjs");

    if (existsSync(candidatePath)) {
      const loaded = requireFromHere(candidatePath) as {
        worksheetControlShapeNamePathCaseTables: WorksheetControlShapeNamePathCaseTables;
      };

      return loaded.worksheetControlShapeNamePathCaseTables;
    }

    currentDirectory = path.resolve(currentDirectory, "..");
  }

  throw new Error("worksheet control shapeName path shared case spec が見つかりません");
}

function getSharedWorkbookRootCompletionEntries(
  familyName: keyof WorkbookRootFamilyCaseTables,
  polarity: "positive",
  options?: { scope?: WorkbookRootFamilyScope; state?: WorkbookRootFamilyState }
): readonly WorkbookRootFamilyPositiveCompletionEntry[];
function getSharedWorkbookRootCompletionEntries(
  familyName: keyof WorkbookRootFamilyCaseTables,
  polarity: "negative",
  options?: { scope?: WorkbookRootFamilyScope; state?: WorkbookRootFamilyState }
): readonly WorkbookRootFamilyNegativeCompletionEntry[];
function getSharedWorkbookRootCompletionEntries(
  familyName: keyof WorkbookRootFamilyCaseTables,
  polarity: "negative" | "positive",
  options: { scope?: WorkbookRootFamilyScope; state?: WorkbookRootFamilyState } = {}
): readonly WorkbookRootFamilyPositiveCompletionEntry[] | readonly WorkbookRootFamilyNegativeCompletionEntry[] {
  const { scope, state } = options;
  return workbookRootFamilyCaseTables[familyName].completion[polarity].filter((entry) => {
    if (scope && !entry.scopes.includes(scope)) {
      return false;
    }
    if (state && entry.state !== state) {
      return false;
    }
    return true;
  });
}

function getSharedWorkbookRootSemanticEntries(
  familyName: WorkbookRootFamilySemanticFamilyName,
  polarity: "negative" | "positive",
  options: { scope?: WorkbookRootFamilyScope; state?: WorkbookRootFamilyState } = {}
): readonly WorkbookRootFamilySemanticEntry[] {
  const { scope, state } = options;
  return workbookRootFamilyCaseTables[familyName].semantic[polarity].filter((entry) => {
    if (scope && !entry.scopes.includes(scope)) {
      return false;
    }
    if (state && entry.state !== state) {
      return false;
    }
    return true;
  });
}

function getSharedWorkbookRootInteractionEntries(
  familyName: keyof WorkbookRootFamilyCaseTables,
  interactionKind: "hover" | "signature",
  polarity: "negative" | "positive",
  options: { scope?: WorkbookRootFamilyScope; state?: WorkbookRootFamilyState } = {}
): readonly WorkbookRootFamilyInteractionEntry[] {
  const { scope, state } = options;
  return workbookRootFamilyCaseTables[familyName][interactionKind][polarity].filter((entry) => {
    if (scope && !entry.scopes.includes(scope)) {
      return false;
    }
    if (state && entry.state !== state) {
      return false;
    }
    return true;
  });
}

function getWorksheetControlShapeNamePathCompletionEntries(
  polarity: "positive",
  options?: {
    fixture?: WorksheetControlShapeNamePathFixture;
    rootKind?: WorksheetControlShapeNamePathRootKind;
    routeKind?: WorksheetControlShapeNamePathRouteKind;
    scope?: WorksheetControlShapeNamePathScope;
  }
): readonly WorksheetControlShapeNamePathPositiveEntry[];
function getWorksheetControlShapeNamePathCompletionEntries(
  polarity: "negative",
  options?: {
    fixture?: WorksheetControlShapeNamePathFixture;
    rootKind?: WorksheetControlShapeNamePathRootKind;
    routeKind?: WorksheetControlShapeNamePathRouteKind;
    scope?: WorksheetControlShapeNamePathScope;
  }
): readonly WorksheetControlShapeNamePathNegativeEntry[];
function getWorksheetControlShapeNamePathCompletionEntries(
  polarity: "negative" | "positive",
  options: {
    fixture?: WorksheetControlShapeNamePathFixture;
    rootKind?: WorksheetControlShapeNamePathRootKind;
    routeKind?: WorksheetControlShapeNamePathRouteKind;
    scope?: WorksheetControlShapeNamePathScope;
  } = {}
):
  | readonly WorksheetControlShapeNamePathNegativeEntry[]
  | readonly WorksheetControlShapeNamePathPositiveEntry[] {
  const { fixture, rootKind, routeKind, scope } = options;
  return worksheetControlShapeNamePathCaseTables.worksheetControlShapeNamePath.completion[polarity].filter((entry) => {
    if (fixture && entry.fixture !== fixture) {
      return false;
    }
    if (rootKind && entry.rootKind !== rootKind) {
      return false;
    }
    if (routeKind && entry.routeKind !== routeKind) {
      return false;
    }
    if (scope && !entry.scopes.includes(scope)) {
      return false;
    }
    return true;
  });
}

function getWorksheetControlShapeNamePathInteractionEntries(
  interactionKind: WorksheetControlShapeNamePathInteractionKind,
  polarity: "positive",
  options?: {
    fixture?: WorksheetControlShapeNamePathFixture;
    rootKind?: WorksheetControlShapeNamePathRootKind;
    routeKind?: WorksheetControlShapeNamePathRouteKind;
    scope?: WorksheetControlShapeNamePathScope;
  }
): readonly WorksheetControlShapeNamePathPositiveEntry[];
function getWorksheetControlShapeNamePathInteractionEntries(
  interactionKind: WorksheetControlShapeNamePathInteractionKind,
  polarity: "negative",
  options?: {
    fixture?: WorksheetControlShapeNamePathFixture;
    rootKind?: WorksheetControlShapeNamePathRootKind;
    routeKind?: WorksheetControlShapeNamePathRouteKind;
    scope?: WorksheetControlShapeNamePathScope;
  }
): readonly WorksheetControlShapeNamePathNegativeEntry[];
function getWorksheetControlShapeNamePathInteractionEntries(
  interactionKind: WorksheetControlShapeNamePathInteractionKind,
  polarity: "negative" | "positive",
  options: {
    fixture?: WorksheetControlShapeNamePathFixture;
    rootKind?: WorksheetControlShapeNamePathRootKind;
    routeKind?: WorksheetControlShapeNamePathRouteKind;
    scope?: WorksheetControlShapeNamePathScope;
  } = {}
): readonly WorksheetControlShapeNamePathNegativeEntry[] | readonly WorksheetControlShapeNamePathPositiveEntry[] {
  const { fixture, rootKind, routeKind, scope } = options;
  return worksheetControlShapeNamePathCaseTables.worksheetControlShapeNamePath[interactionKind][polarity].filter((entry) => {
    if (fixture && entry.fixture !== fixture) {
      return false;
    }
    if (rootKind && entry.rootKind !== rootKind) {
      return false;
    }
    if (routeKind && entry.routeKind !== routeKind) {
      return false;
    }
    if (scope && !entry.scopes.includes(scope)) {
      return false;
    }
    return true;
  });
}

function getWorksheetControlShapeNamePathSemanticEntries(
  polarity: "positive",
  options?: {
    fixture?: WorksheetControlShapeNamePathFixture;
    rootKind?: WorksheetControlShapeNamePathRootKind;
    routeKind?: WorksheetControlShapeNamePathRouteKind;
    scope?: WorksheetControlShapeNamePathScope;
  }
): readonly WorksheetControlShapeNamePathSemanticEntry[];
function getWorksheetControlShapeNamePathSemanticEntries(
  polarity: "negative",
  options?: {
    fixture?: WorksheetControlShapeNamePathFixture;
    rootKind?: WorksheetControlShapeNamePathRootKind;
    routeKind?: WorksheetControlShapeNamePathRouteKind;
    scope?: WorksheetControlShapeNamePathScope;
  }
): readonly WorksheetControlShapeNamePathSemanticEntry[];
function getWorksheetControlShapeNamePathSemanticEntries(
  polarity: "negative" | "positive",
  options: {
    fixture?: WorksheetControlShapeNamePathFixture;
    rootKind?: WorksheetControlShapeNamePathRootKind;
    routeKind?: WorksheetControlShapeNamePathRouteKind;
    scope?: WorksheetControlShapeNamePathScope;
  } = {}
): readonly WorksheetControlShapeNamePathSemanticEntry[] {
  const { fixture, rootKind, routeKind, scope } = options;
  return worksheetControlShapeNamePathCaseTables.worksheetControlShapeNamePath.semantic[polarity].filter((entry) => {
    if (fixture && entry.fixture !== fixture) {
      return false;
    }
    if (rootKind && entry.rootKind !== rootKind) {
      return false;
    }
    if (routeKind && entry.routeKind !== routeKind) {
      return false;
    }
    if (scope && !entry.scopes.includes(scope)) {
      return false;
    }
    return true;
  });
}

function mapExtensionWorkbookRootPositiveCompletionCases(
  entries: readonly WorkbookRootFamilyPositiveCompletionEntry[],
  messageBuilder: (entry: WorkbookRootFamilyPositiveCompletionEntry) => string
): readonly WorkbookRootCompletionCase[] {
  assert.ok(entries.length > 0, "workbook root positive completion shared cases must not be empty");
  return entries.map((entry) => [
    entry.anchor,
    "CheckBox property",
    entry.route === "shape" ? "Delete" : "Activate",
    messageBuilder(entry),
    entry.occurrenceIndex ?? 0
  ]);
}

function mapExtensionWorkbookRootClosedCompletionCases(
  entries: readonly WorkbookRootFamilyCaseEntryBase[],
  messageBuilder: (entry: WorkbookRootFamilyCaseEntryBase) => string
): readonly WorkbookRootClosedCompletionCase[] {
  assert.ok(entries.length > 0, "workbook root closed completion shared cases must not be empty");
  return entries.map((entry) => [entry.anchor, messageBuilder(entry), entry.occurrenceIndex ?? 0]);
}

function mapExtensionWorkbookRootInteractionCases(
  entries: readonly WorkbookRootFamilyInteractionEntry[],
  messageBuilder: (entry: WorkbookRootFamilyInteractionEntry) => string
): readonly WorkbookRootHoverCase[] {
  assert.ok(entries.length > 0, "workbook root interaction shared cases must not be empty");
  return entries.map((entry) => [entry.anchor, messageBuilder(entry), entry.occurrenceIndex ?? 0]);
}

function mapExtensionWorkbookRootSemanticCases(
  entries: readonly WorkbookRootFamilySemanticEntry[],
  messageBuilder: (entry: WorkbookRootFamilySemanticEntry) => string
): readonly WorkbookRootSemanticCase[] {
  assert.ok(entries.length > 0, "workbook root semantic shared cases must not be empty");
  return entries.map((entry) => [
    entry.anchor,
    entry.identifier,
    { modifiers: [], type: entry.tokenKind === "method" ? "function" : "variable" },
    messageBuilder(entry),
    entry.occurrenceIndex ?? 0
  ]);
}

function mapExtensionWorkbookRootNoSemanticCases(
  entries: readonly WorkbookRootFamilySemanticEntry[],
  messageBuilder: (entry: WorkbookRootFamilySemanticEntry) => string
): readonly WorkbookRootNoSemanticCase[] {
  assert.ok(entries.length > 0, "workbook root negative semantic shared cases must not be empty");
  return entries.map((entry) => [entry.anchor, entry.identifier, messageBuilder(entry), entry.occurrenceIndex ?? 0]);
}

function mapExtensionWorksheetControlShapeNamePathPositiveCompletionCases(
  entries: readonly WorksheetControlShapeNamePathPositiveEntry[],
  messageBuilder: (entry: WorksheetControlShapeNamePathPositiveEntry) => string
): readonly WorkbookRootCompletionCase[] {
  assert.ok(entries.length > 0, "worksheet control shapeName path positive completion shared cases must not be empty");
  return entries.map((entry) => [
    entry.anchor,
    "CheckBox property",
    entry.routeKind === "shape-oleformat" ? "Delete" : "Activate",
    messageBuilder(entry),
    entry.occurrenceIndex ?? 0
  ]);
}

function mapExtensionWorksheetControlShapeNamePathNoCompletionCases(
  entries: readonly WorksheetControlShapeNamePathNegativeEntry[],
  messageBuilder: (entry: WorksheetControlShapeNamePathNegativeEntry) => string
): readonly WorkbookRootClosedCompletionCase[] {
  assert.ok(entries.length > 0, "worksheet control shapeName path negative completion shared cases must not be empty");
  return entries.map((entry) => [entry.anchor, messageBuilder(entry), entry.occurrenceIndex ?? 0]);
}

function mapExtensionWorksheetControlShapeNamePathInteractionCases<
  T extends WorksheetControlShapeNamePathPositiveEntry | WorksheetControlShapeNamePathNegativeEntry
>(
  entries: readonly T[],
  messageBuilder: (entry: T) => string
): readonly WorkbookRootHoverCase[] {
  assert.ok(entries.length > 0, "worksheet control shapeName path interaction shared cases must not be empty");
  return entries.map((entry) => [entry.anchor, messageBuilder(entry), entry.occurrenceIndex ?? 0]);
}

function mapExtensionWorksheetControlShapeNamePathSemanticCases(
  entries: readonly WorksheetControlShapeNamePathSemanticEntry[],
  messageBuilder: (entry: WorksheetControlShapeNamePathSemanticEntry) => string
): readonly WorkbookRootSemanticCase[] {
  assert.ok(entries.length > 0, "worksheet control shapeName path semantic shared cases must not be empty");
  return entries.map((entry) => [
    entry.anchor,
    entry.identifier,
    { modifiers: [], type: entry.tokenKind === "method" ? "function" : "variable" },
    messageBuilder(entry),
    entry.occurrenceIndex ?? 0
  ]);
}

function mapExtensionWorksheetControlShapeNamePathNoSemanticCases(
  entries: readonly WorksheetControlShapeNamePathSemanticEntry[],
  messageBuilder: (entry: WorksheetControlShapeNamePathSemanticEntry) => string
): readonly WorkbookRootNoSemanticCase[] {
  assert.ok(entries.length > 0, "worksheet control shapeName path negative semantic shared cases must not be empty");
  return entries.map((entry) => [entry.anchor, entry.identifier, messageBuilder(entry), entry.occurrenceIndex ?? 0]);
}

type WorkbookRootCompletionCase = readonly [string, string, string, string, number?];
type WorkbookRootClosedCompletionCase = readonly [string, string, number?];
type WorkbookRootHoverCase = readonly [string, string, number?];
type WorkbookRootSignatureCase = readonly [string, string, number?];
type WorkbookRootSemanticCase = readonly [string, string, { modifiers: readonly string[]; type: string }, string?, number?];
type WorkbookRootNoSemanticCase = readonly [string, string, string?, number?];

async function runExtensionWorksheetControlShapeNamePathInteractionSharedCases({
  document,
  fixture,
  routeKind
}: {
  document: vscode.TextDocument;
  fixture: WorksheetControlShapeNamePathFixture;
  routeKind: WorksheetControlShapeNamePathRouteKind;
}): Promise<void> {
  const originalSnapshot = createRestorableActiveWorkbookIdentitySnapshot(await getActiveWorkbookIdentitySnapshot());
  const positiveHoverEntries = getWorksheetControlShapeNamePathInteractionEntries("hover", "positive", {
    fixture,
    routeKind,
    scope: "extension"
  });
  const alwaysAvailablePositiveHoverEntries = positiveHoverEntries.filter((entry) => entry.rootKind !== "workbook-qualified-matched");
  const negativeHoverEntries = getWorksheetControlShapeNamePathInteractionEntries("hover", "negative", {
    fixture,
    routeKind,
    scope: "extension"
  });
  const closedHoverEntries = negativeHoverEntries.filter((entry) => entry.rootKind === "workbook-qualified-closed");
  const reasonHoverEntries = negativeHoverEntries.filter((entry) => entry.rootKind !== "workbook-qualified-closed");
  const positiveSignatureEntries = getWorksheetControlShapeNamePathInteractionEntries("signature", "positive", {
    fixture,
    routeKind,
    scope: "extension"
  });
  const alwaysAvailablePositiveSignatureEntries = positiveSignatureEntries.filter(
    (entry) => entry.rootKind !== "workbook-qualified-matched"
  );
  const negativeSignatureEntries = getWorksheetControlShapeNamePathInteractionEntries("signature", "negative", {
    fixture,
    routeKind,
    scope: "extension"
  });
  const closedSignatureEntries = negativeSignatureEntries.filter((entry) => entry.rootKind === "workbook-qualified-closed");
  const reasonSignatureEntries = negativeSignatureEntries.filter((entry) => entry.rootKind !== "workbook-qualified-closed");
  const positiveSemanticEntries = getWorksheetControlShapeNamePathSemanticEntries("positive", {
    fixture,
    routeKind,
    scope: "extension"
  });
  const alwaysAvailablePositiveSemanticEntries = positiveSemanticEntries.filter(
    (entry) => entry.rootKind !== "workbook-qualified-matched"
  );
  const negativeSemanticEntries = getWorksheetControlShapeNamePathSemanticEntries("negative", {
    fixture,
    routeKind,
    scope: "extension"
  });
  const closedSemanticEntries = negativeSemanticEntries.filter((entry) => entry.rootKind === "workbook-qualified-closed");
  const reasonSemanticEntries = negativeSemanticEntries.filter((entry) => entry.rootKind !== "workbook-qualified-closed");
  const legend = await waitForSemanticTokensLegend(document, (value) => value.tokenTypes.includes("variable") && value.tokenTypes.includes("function"));
  const text = document.getText();

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT);
  await assertWorkbookRootHoverCases(
    document,
    mapExtensionWorksheetControlShapeNamePathInteractionCases(
      alwaysAvailablePositiveHoverEntries,
      (entry) => `${entry.anchor} は ${entry.rootKind} root なので snapshot なしでも hover が control owner を指す`
    )
  );
  await assertWorkbookRootNoHoverCases(
    document,
    mapExtensionWorksheetControlShapeNamePathInteractionCases(
      [...reasonHoverEntries, ...closedHoverEntries],
      (entry) =>
        entry.rootKind === "workbook-qualified-closed"
          ? `${entry.anchor} は active workbook が閉じている間は hover を解決しない`
          : `${entry.anchor} は ${entry.reason} のため hover を解決しない`
    )
  );
  await assertWorkbookRootSignatureCases(
    document,
    mapExtensionWorksheetControlShapeNamePathInteractionCases(
      alwaysAvailablePositiveSignatureEntries,
      (entry) => `${entry.anchor} は ${entry.rootKind} root なので snapshot なしでも signature help を解決する`
    )
  );
  await assertWorkbookRootNoSignatureCases(
    document,
    mapExtensionWorksheetControlShapeNamePathInteractionCases(
      [...reasonSignatureEntries, ...closedSignatureEntries],
      (entry) =>
        entry.rootKind === "workbook-qualified-closed"
          ? `${entry.anchor} は active workbook が閉じている間は signature help を解決しない`
          : `${entry.anchor} は ${entry.reason} のため signature help を解決しない`
    )
  );
  let decodedTokens = await waitForSemanticTokensByCases(
    document,
    legend,
    text,
    mapExtensionWorksheetControlShapeNamePathSemanticCases(
      alwaysAvailablePositiveSemanticEntries,
      (entry) =>
        entry.rootKind === "workbook-qualified-matched"
          ? `${entry.anchor} は snapshot なしでは semantic token を出さない`
          : `${entry.anchor} は ${entry.rootKind} root なので snapshot なしでも semantic token を出す`
    ),
    mapExtensionWorksheetControlShapeNamePathNoSemanticCases(
      [...reasonSemanticEntries, ...closedSemanticEntries],
      (entry) =>
        entry.rootKind === "workbook-qualified-closed"
          ? `${entry.anchor} は active workbook が閉じている間は semantic token を出さない`
          : `${entry.anchor} は ${entry.reason} のため semantic token を出さない`
    )
  );
  assertWorkbookRootSemanticCases(
    text,
    decodedTokens,
    mapExtensionWorksheetControlShapeNamePathSemanticCases(
      alwaysAvailablePositiveSemanticEntries,
      (entry) => `${entry.anchor} は ${entry.rootKind} root なので snapshot なしでも semantic token を出す`
    )
  );

  await setActiveWorkbookIdentitySnapshot(ACTIVE_WORKBOOK_AVAILABLE_SNAPSHOT);
  try {
    await assertWorkbookRootHoverCases(
      document,
      mapExtensionWorksheetControlShapeNamePathInteractionCases(
        positiveHoverEntries,
        (entry) =>
          entry.rootKind === "workbook-qualified-matched"
            ? `${entry.anchor} は active workbook match 時に hover が control owner を指す`
            : `${entry.anchor} は ${entry.rootKind} root として hover が control owner を指す`
      )
    );
    await assertWorkbookRootNoHoverCases(
      document,
      mapExtensionWorksheetControlShapeNamePathInteractionCases(
        reasonHoverEntries,
        (entry) => `${entry.anchor} は ${entry.reason} のため match 中でも hover を解決しない`
      )
    );
    await assertWorkbookRootSignatureCases(
      document,
      mapExtensionWorksheetControlShapeNamePathInteractionCases(
        positiveSignatureEntries,
        (entry) =>
          entry.rootKind === "workbook-qualified-matched"
            ? `${entry.anchor} は active workbook match 時に signature help を解決する`
            : `${entry.anchor} は ${entry.rootKind} root として signature help を解決する`
      )
    );
    await assertWorkbookRootNoSignatureCases(
      document,
      mapExtensionWorksheetControlShapeNamePathInteractionCases(
        reasonSignatureEntries,
        (entry) => `${entry.anchor} は ${entry.reason} のため match 中でも signature help を解決しない`
      )
    );
    decodedTokens = await waitForSemanticTokensByCases(
      document,
      legend,
      text,
      mapExtensionWorksheetControlShapeNamePathSemanticCases(
        positiveSemanticEntries,
        (entry) =>
          entry.rootKind === "workbook-qualified-matched"
            ? `${entry.anchor} は active workbook match 時に semantic token を出す`
            : `${entry.anchor} は ${entry.rootKind} root として semantic token を出す`
      ),
      mapExtensionWorksheetControlShapeNamePathNoSemanticCases(
        reasonSemanticEntries,
        (entry) => `${entry.anchor} は ${entry.reason} のため match 中でも semantic token を出さない`
      )
    );
    assertWorkbookRootNoSemanticCases(
      text,
      decodedTokens,
      mapExtensionWorksheetControlShapeNamePathNoSemanticCases(
        reasonSemanticEntries,
        (entry) => `${entry.anchor} は ${entry.reason} のため match 中でも semantic token を出さない`
      )
    );
  } finally {
    await setActiveWorkbookIdentitySnapshot(originalSnapshot);
  }
}

async function assertWorkbookRootCompletionCases(
  document: vscode.TextDocument,
  cases: readonly WorkbookRootCompletionCase[]
): Promise<void> {
  for (const [token, detailFragment, blockedLabel, message, occurrenceIndex = 0] of cases) {
    const position = findPositionAfterToken(document, token, 0, occurrenceIndex);
    const items = await waitForCompletionLabelStateAtToken(document, token, "Value", true, occurrenceIndex);
    const completion = items.find((item) => getCompletionItemLabel(item) === "Value");
    const completionSummary = items
      .slice(0, 12)
      .map((item) => `${getCompletionItemLabel(item)} :: ${item.detail ?? ""}`)
      .join(" | ");
    const lineText = document.lineAt(position.line).text;
    const context = `line ${position.line + 1}, char ${position.character + 1}: ${lineText}`;

    assert.ok(completion?.detail?.includes(detailFragment), `${message} / ${context} / completions: ${completionSummary}`);
    assert.equal(
      hasCompletionItemLabel(items, blockedLabel),
      false,
      `${message} / ${context} / completions: ${completionSummary}`
    );
  }
}

async function assertWorkbookRootClosedCompletionCases(
  document: vscode.TextDocument,
  cases: readonly WorkbookRootClosedCompletionCase[]
): Promise<void> {
  for (const [token, message, occurrenceIndex = 0] of cases) {
    const items = await waitForCompletionLabelStateAtToken(document, token, "Value", false, occurrenceIndex);
    assert.equal(hasCompletionItemLabel(items, "Value"), false, message);
  }
}

async function assertWorkbookRootHoverCases(
  document: vscode.TextDocument,
  cases: readonly WorkbookRootHoverCase[],
  expectedFragment = "CheckBox.Value"
): Promise<void> {
  for (const [token, message, occurrenceIndex = 0] of cases) {
    const hovers = await waitForHoverAtToken(document, token, (items) => items.length > 0, occurrenceIndex);
    assert.equal(getHoverContentsText(hovers[0]).includes(expectedFragment), true, message);
  }
}

async function assertWorkbookRootNoHoverCases(
  document: vscode.TextDocument,
  cases: readonly WorkbookRootHoverCase[]
): Promise<void> {
  for (const [token, message, occurrenceIndex = 0] of cases) {
    assert.equal(await waitForNoHoverAtToken(document, token, occurrenceIndex), true, message);
  }
}

async function assertWorkbookRootSignatureCases(
  document: vscode.TextDocument,
  cases: readonly WorkbookRootSignatureCase[],
  expectedLabel = "Select(Replace) As Object"
): Promise<void> {
  for (const [token, message, occurrenceIndex = 0] of cases) {
    const signatureHelp = await waitForSignatureHelpAtToken(document, token, (help) => help.signatures.length > 0, occurrenceIndex);
    assert.equal(signatureHelp.signatures[0]?.label, expectedLabel, message);
  }
}

async function assertWorkbookRootNoSignatureCases(
  document: vscode.TextDocument,
  cases: readonly WorkbookRootSignatureCase[]
): Promise<void> {
  for (const [token, message, occurrenceIndex = 0] of cases) {
    assert.equal(await waitForNoSignatureHelpAtToken(document, token, occurrenceIndex), true, message);
  }
}

function assertWorkbookRootSemanticCases(
  text: string,
  tokens: Array<{ endCharacter: number; line: number; modifiers: string[]; startCharacter: number; type: string }>,
  cases: readonly WorkbookRootSemanticCase[]
): void {
  for (const [anchorToken, identifier, expected, message, occurrenceIndex = 0] of cases) {
    assertDecodedSemanticTokenByAnchor(text, tokens, anchorToken, identifier, expected, occurrenceIndex, message);
  }
}

function assertWorkbookRootNoSemanticCases(
  text: string,
  tokens: Array<{ endCharacter: number; line: number; modifiers: string[]; startCharacter: number; type: string }>,
  cases: readonly WorkbookRootNoSemanticCase[]
): void {
  for (const [anchorToken, identifier, message, occurrenceIndex = 0] of cases) {
    assertNoDecodedSemanticTokenByAnchor(text, tokens, anchorToken, identifier, occurrenceIndex, message);
  }
}

async function waitForRename(
  document: vscode.TextDocument,
  position: vscode.Position,
  newName: string,
  predicate: (edit: vscode.WorkspaceEdit) => boolean
): Promise<vscode.WorkspaceEdit> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const renameEdit = await vscode.commands.executeCommand<vscode.WorkspaceEdit>(
      "vscode.executeDocumentRenameProvider",
      document.uri,
      position,
      newName
    );

    if (renameEdit && predicate(renameEdit)) {
      return renameEdit;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return new vscode.WorkspaceEdit();
}

async function waitForSemanticTokensLegend(
  document: vscode.TextDocument,
  predicate: (legend: vscode.SemanticTokensLegend) => boolean
): Promise<vscode.SemanticTokensLegend> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const legend = await vscode.commands.executeCommand<vscode.SemanticTokensLegend>(
      "vscode.provideDocumentSemanticTokensLegend",
      document.uri
    );

    if (legend && predicate(legend)) {
      return legend;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return new vscode.SemanticTokensLegend([]);
}

async function waitForSemanticTokens(
  document: vscode.TextDocument,
  predicate: (tokens: vscode.SemanticTokens) => boolean
): Promise<vscode.SemanticTokens> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const tokens = await vscode.commands.executeCommand<vscode.SemanticTokens>(
      "vscode.provideDocumentSemanticTokens",
      document.uri
    );

    if (tokens && predicate(tokens)) {
      return tokens;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return new vscode.SemanticTokens(new Uint32Array());
}

async function waitForSemanticTokensByCases(
  document: vscode.TextDocument,
  legend: vscode.SemanticTokensLegend,
  text: string,
  positiveCases: readonly WorkbookRootSemanticCase[],
  noCases: readonly WorkbookRootNoSemanticCase[]
): Promise<Array<{ endCharacter: number; line: number; modifiers: string[]; startCharacter: number; type: string }>> {
  let latestTokenCount = 0;
  let latestFailureMessage = "semantic token の安定待機が開始されませんでした。";

  for (let attempt = 0; attempt < 30; attempt += 1) {
    const tokens =
      (await vscode.commands.executeCommand<vscode.SemanticTokens>("vscode.provideDocumentSemanticTokens", document.uri)) ??
      new vscode.SemanticTokens(new Uint32Array());
    const decodedTokens = decodeSemanticTokens(tokens, legend);

    latestTokenCount = tokens.data.length;
    try {
      assertWorkbookRootSemanticCases(text, decodedTokens, positiveCases);
      assertWorkbookRootNoSemanticCases(text, decodedTokens, noCases);
      return decodedTokens;
    } catch (error) {
      latestFailureMessage = error instanceof Error ? error.message : String(error);
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  assert.fail(
    [
      "waitForSemanticTokensByCases が対象 anchor の semantic 状態に到達しませんでした。",
      `document=${document.uri.fsPath}`,
      `latestTokenCount=${latestTokenCount}`,
      `lastFailure=${latestFailureMessage}`
    ].join(" ")
  );
}

async function waitForNoSemanticTokensByCases(
  document: vscode.TextDocument,
  legend: vscode.SemanticTokensLegend,
  text: string,
  cases: readonly WorkbookRootNoSemanticCase[]
): Promise<Array<{ endCharacter: number; line: number; modifiers: string[]; startCharacter: number; type: string }>> {
  let latestTokenCount = 0;
  let latestFailureMessage = "semantic token の安定待機が開始されませんでした。";

  for (let attempt = 0; attempt < 30; attempt += 1) {
    const tokens =
      (await vscode.commands.executeCommand<vscode.SemanticTokens>("vscode.provideDocumentSemanticTokens", document.uri)) ??
      new vscode.SemanticTokens(new Uint32Array());
    const decodedTokens = decodeSemanticTokens(tokens, legend);

    latestTokenCount = tokens.data.length;
    try {
      assertWorkbookRootNoSemanticCases(text, decodedTokens, cases);
      return decodedTokens;
    } catch (error) {
      latestFailureMessage = error instanceof Error ? error.message : String(error);
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  assert.fail(
    [
      "waitForNoSemanticTokensByCases が対象 anchor の no-semantic 状態に到達しませんでした。",
      `document=${document.uri.fsPath}`,
      `latestTokenCount=${latestTokenCount}`,
      `lastFailure=${latestFailureMessage}`
    ].join(" ")
  );
}

async function setActiveWorkbookIdentitySnapshot(snapshot: unknown): Promise<void> {
  let lastSetError: unknown;
  let lastObservedState: Record<string, unknown> | null | undefined;

  for (let attempt = 0; attempt < 30; attempt += 1) {
    try {
      await vscode.commands.executeCommand(TEST_SET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND, snapshot);
      lastSetError = undefined;
      break;
    } catch (error) {
      lastSetError = error;
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
  }

  if (lastSetError) {
    throw lastSetError instanceof Error ? lastSetError : new Error(String(lastSetError));
  }

  for (let attempt = 0; attempt < 30; attempt += 1) {
    let observedState: Record<string, unknown> | null | undefined;

    try {
      observedState = await vscode.commands.executeCommand<Record<string, unknown> | null>(
        TEST_GET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND
      );
    } catch {
      observedState = undefined;
    }

    lastObservedState = observedState;

    if (matchesActiveWorkbookIdentityState(observedState, snapshot)) {
      return;
    }

    await new Promise((resolve) => setTimeout(resolve, 100));
  }

  throw new Error(
    `active workbook identity snapshot が server へ反映されませんでした expected=${JSON.stringify(snapshot)} observed=${JSON.stringify(lastObservedState)}`
  );
}

async function getActiveWorkbookIdentitySnapshot(): Promise<Record<string, unknown> | null> {
  return (
    (await vscode.commands.executeCommand<Record<string, unknown> | null>(
      TEST_GET_ACTIVE_WORKBOOK_IDENTITY_SNAPSHOT_COMMAND
    )) ?? null
  );
}

function createRestorableActiveWorkbookIdentitySnapshot(
  state: Record<string, unknown> | null | undefined
): Record<string, unknown> {
  if (!state) {
    return { ...ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT };
  }

  const observedAt =
    typeof state.observedAt === "string" ? state.observedAt : ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT.observedAt;

  switch (typeof state.state === "string" ? state.state : undefined) {
    case "available": {
      const workbook = state.workbook;

      if (
        workbook &&
        typeof workbook === "object" &&
        typeof (workbook as { fullName?: unknown }).fullName === "string" &&
        typeof (workbook as { isAddin?: unknown }).isAddin === "boolean" &&
        typeof (workbook as { name?: unknown }).name === "string" &&
        typeof (workbook as { path?: unknown }).path === "string"
      ) {
        return {
          identity: {
            fullName: (workbook as { fullName: string }).fullName,
            isAddin: (workbook as { isAddin: boolean }).isAddin,
            name: (workbook as { name: string }).name,
            path: (workbook as { path: string }).path
          },
          observedAt,
          providerKind: "excel-active-workbook",
          state: "available",
          version: 1
        };
      }

      break;
    }
    case "unavailable":
      if (typeof state.reason === "string") {
        return {
          observedAt,
          providerKind: "excel-active-workbook",
          reason: state.reason,
          state: "unavailable",
          version: 1
        };
      }
      break;
    case "protected-view":
      if (state.protectedView && typeof state.protectedView === "object") {
        return {
          observedAt,
          protectedView: {
            ...(typeof (state.protectedView as { sourceName?: unknown }).sourceName === "string"
              ? { sourceName: (state.protectedView as { sourceName: string }).sourceName }
              : {}),
            ...(typeof (state.protectedView as { sourcePath?: unknown }).sourcePath === "string"
              ? { sourcePath: (state.protectedView as { sourcePath: string }).sourcePath }
              : {})
          },
          providerKind: "excel-active-workbook",
          state: "protected-view",
          version: 1
        };
      }
      break;
    case "unsupported": {
      const workbook = state.workbook;

      if (
        workbook &&
        typeof workbook === "object" &&
        typeof (workbook as { fullName?: unknown }).fullName === "string" &&
        typeof (workbook as { isAddin?: unknown }).isAddin === "boolean" &&
        typeof (workbook as { name?: unknown }).name === "string" &&
        typeof (workbook as { path?: unknown }).path === "string" &&
        typeof state.reason === "string"
      ) {
        return {
          identity: {
            fullName: (workbook as { fullName: string }).fullName,
            isAddin: (workbook as { isAddin: boolean }).isAddin,
            name: (workbook as { name: string }).name,
            path: (workbook as { path: string }).path
          },
          observedAt,
          providerKind: "excel-active-workbook",
          reason: state.reason,
          state: "unsupported",
          version: 1
        };
      }

      break;
    }
    default:
      break;
  }

  return { ...ACTIVE_WORKBOOK_UNAVAILABLE_SNAPSHOT };
}

function matchesActiveWorkbookIdentityState(
  observedState: Record<string, unknown> | null | undefined,
  snapshot: unknown
): boolean {
  if (!observedState || typeof snapshot !== "object" || snapshot === null) {
    return false;
  }

  const expectedState = typeof (snapshot as { state?: unknown }).state === "string" ? (snapshot as { state: string }).state : undefined;
  const observedKind = typeof observedState.state === "string" ? observedState.state : undefined;

  if (!expectedState || observedKind !== expectedState) {
    return false;
  }

  if (expectedState === "available") {
    const rawFullName = typeof observedState.rawFullName === "string" ? observedState.rawFullName : undefined;
    const expectedIdentity = (snapshot as { identity?: { fullName?: unknown } }).identity;
    const expectedFullName = typeof expectedIdentity?.fullName === "string" ? expectedIdentity.fullName : undefined;

    return normalizePathForComparison(rawFullName) === normalizePathForComparison(expectedFullName);
  }

  if ("reason" in snapshot) {
    return observedState.reason === (snapshot as { reason?: unknown }).reason;
  }

  return true;
}

function normalizePathForComparison(value: string | undefined): string {
  return value?.replace(/\//g, "\\").toLowerCase() ?? "";
}

async function waitForFormattedDocument(document: vscode.TextDocument, expectedText: string): Promise<string> {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    await vscode.commands.executeCommand("editor.action.formatDocument");
    const currentText = document.getText();

    if (normalizeText(currentText) === normalizeText(expectedText)) {
      return currentText;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return document.getText();
}

function normalizeText(text: string): string {
  return text.replace(/\r\n?/g, "\n").trimEnd();
}

function findPositionAfterToken(
  document: vscode.TextDocument,
  token: string,
  offsetFromEnd = 0,
  occurrenceIndex = 0
): vscode.Position {
  const source = document.getText();
  const startIndex = findTokenStartIndex(source, token, occurrenceIndex);
  return document.positionAt(startIndex + token.length + offsetFromEnd);
}

function assertVariadicWorksheetFunctionSignature(
  signatureHelp: vscode.SignatureHelp,
  expectedLabel: string,
  memberName: string
): void {
  assert.equal(
    signatureHelp.signatures[0]?.label,
    expectedLabel,
    `built-in member signature should be available for WorksheetFunction.${memberName}`
  );
  assert.equal(
    signatureHelp.signatures[0]?.parameters.length,
    30,
    `built-in member ${memberName} signature should expose variadic parameter metadata`
  );
  assert.ok(
    getSignatureDocumentation(signatureHelp.signatures[0]?.parameters[0]?.documentation).includes("想定型: Variant"),
    `built-in member ${memberName} first argument should include expected type`
  );
  assert.ok(
    getSignatureDocumentation(signatureHelp.signatures[0]?.parameters[0]?.documentation).includes("必須引数"),
    `built-in member ${memberName} first argument should be required`
  );
  assert.ok(
    getSignatureDocumentation(signatureHelp.signatures[0]?.parameters[1]?.documentation).includes("省略可能"),
    `built-in member ${memberName} second argument should be optional`
  );
  assert.ok(
    getSignatureDocumentation(signatureHelp.signatures[0]?.parameters[29]?.documentation).includes("省略可能"),
    `built-in member ${memberName} last argument should be optional`
  );
}

function getSignatureDocumentation(
  documentation: vscode.MarkdownString | vscode.ParameterInformation["documentation"] | vscode.SignatureInformation["documentation"] | undefined
): string {
  if (!documentation) {
    return "";
  }

  return typeof documentation === "string" ? documentation : documentation.value;
}

function getHoverContentsText(hover: vscode.Hover | undefined): string {
  if (!hover) {
    return "";
  }

  const contents = Array.isArray(hover.contents) ? hover.contents : [hover.contents];

  return contents
    .map((content) => {
      if (typeof content === "string") {
        return content;
      }

      if ("language" in content && "value" in content) {
        return content.value;
      }

      return content.value;
    })
    .join("\n");
}

function hasCompletionItemLabel(items: readonly vscode.CompletionItem[], label: string): boolean {
  return items.some((item) => getCompletionItemLabel(item) === label);
}

function getCompletionItemLabel(item: vscode.CompletionItem): string {
  return typeof item.label === "string" ? item.label : item.label.label;
}

function hasSnippetCompletion(items: readonly vscode.CompletionItem[], label: string): boolean {
  return items.some(
    (item) =>
      item.kind === vscode.CompletionItemKind.Snippet &&
      getCompletionItemLabel(item).toLowerCase() === label.toLowerCase()
  );
}

function decodeSemanticTokens(
  tokens: vscode.SemanticTokens,
  legend: vscode.SemanticTokensLegend
): Array<{ endCharacter: number; line: number; modifiers: string[]; startCharacter: number; type: string }> {
  const decodedTokens: Array<{ endCharacter: number; line: number; modifiers: string[]; startCharacter: number; type: string }> = [];
  let currentLine = 0;
  let currentCharacter = 0;

  for (let index = 0; index < tokens.data.length; index += 5) {
    const deltaLine = tokens.data[index] ?? 0;
    const deltaCharacter = tokens.data[index + 1] ?? 0;
    const length = tokens.data[index + 2] ?? 0;
    const tokenType = legend.tokenTypes[tokens.data[index + 3] ?? 0] ?? "";
    const modifierMask = tokens.data[index + 4] ?? 0;

    currentLine += deltaLine;
    currentCharacter = deltaLine === 0 ? currentCharacter + deltaCharacter : deltaCharacter;

    decodedTokens.push({
      endCharacter: currentCharacter + length,
      line: currentLine,
      modifiers: legend.tokenModifiers.filter((_, modifierIndex) => (modifierMask & (1 << modifierIndex)) !== 0),
      startCharacter: currentCharacter,
      type: tokenType
    });
  }

  return decodedTokens;
}

function assertDecodedSemanticToken(
  text: string,
  tokens: Array<{ endCharacter: number; line: number; modifiers: string[]; startCharacter: number; type: string }>,
  lineIndex: number,
  identifier: string,
  expected: { modifiers: readonly string[]; type: string },
  occurrence = 0
): void {
  const lines = text.split("\n");
  const line = lines[lineIndex] ?? "";
  let startCharacter = -1;
  let searchOffset = 0;

  for (let index = 0; index <= occurrence; index += 1) {
    startCharacter = line.indexOf(identifier, searchOffset);
    searchOffset = startCharacter + identifier.length;
  }

  assert.notEqual(startCharacter, -1, `identifier '${identifier}' must exist on line ${lineIndex}`);

  const token = tokens.find(
    (entry) =>
      entry.line === lineIndex &&
      entry.startCharacter === startCharacter &&
      entry.endCharacter === startCharacter + identifier.length
  );

  assert.ok(token, `semantic token '${identifier}' must exist at ${lineIndex}:${startCharacter}`);
  assert.equal(token.type, expected.type);
  assert.deepEqual([...token.modifiers].sort(), [...expected.modifiers].sort());
}

function assertDecodedSemanticTokenByAnchor(
  text: string,
  tokens: Array<{ endCharacter: number; line: number; modifiers: string[]; startCharacter: number; type: string }>,
  anchorToken: string,
  identifier: string,
  expected: { modifiers: readonly string[]; type: string },
  occurrenceIndex = 0,
  message?: string
): void {
  const { lineIndex, startCharacter } = findIdentifierPositionInTokenOccurrence(text, anchorToken, identifier, occurrenceIndex);
  const token = tokens.find(
    (entry) =>
      entry.line === lineIndex &&
      entry.startCharacter === startCharacter &&
      entry.endCharacter === startCharacter + identifier.length
  );

  assert.ok(token, message ?? `semantic token '${identifier}' must exist at ${lineIndex}:${startCharacter}`);
  assert.equal(token.type, expected.type, message);
  assert.deepEqual([...token.modifiers].sort(), [...expected.modifiers].sort(), message);
}

function assertNoDecodedSemanticToken(
  text: string,
  tokens: Array<{ endCharacter: number; line: number; modifiers: string[]; startCharacter: number; type: string }>,
  lineIndex: number,
  identifier: string,
  occurrence = 0
): void {
  const lines = text.split("\n");
  const line = lines[lineIndex] ?? "";
  let startCharacter = -1;
  let searchOffset = 0;

  for (let index = 0; index <= occurrence; index += 1) {
    startCharacter = line.indexOf(identifier, searchOffset);
    searchOffset = startCharacter + identifier.length;
  }

  assert.notEqual(startCharacter, -1, `identifier '${identifier}' must exist on line ${lineIndex}`);
  assert.equal(
    tokens.some(
      (entry) =>
        entry.line === lineIndex &&
        entry.startCharacter === startCharacter &&
        entry.endCharacter === startCharacter + identifier.length
    ),
    false,
    `semantic token '${identifier}' must not exist at ${lineIndex}:${startCharacter}`
  );
}

function assertNoDecodedSemanticTokenByAnchor(
  text: string,
  tokens: Array<{ endCharacter: number; line: number; modifiers: string[]; startCharacter: number; type: string }>,
  anchorToken: string,
  identifier: string,
  occurrenceIndex = 0,
  message?: string
): void {
  const { lineIndex, startCharacter } = findIdentifierPositionInTokenOccurrence(text, anchorToken, identifier, occurrenceIndex);

  assert.equal(
    tokens.some(
      (entry) =>
        entry.line === lineIndex &&
        entry.startCharacter === startCharacter &&
        entry.endCharacter === startCharacter + identifier.length
    ),
    false,
    message ?? `semantic token '${identifier}' must not exist at ${lineIndex}:${startCharacter}`
  );
}

function findIdentifierPositionInTokenOccurrence(
  text: string,
  anchorToken: string,
  identifier: string,
  occurrenceIndex = 0
): { lineIndex: number; startCharacter: number } {
  assert.equal(anchorToken.includes("\n"), false, `anchor token must stay on a single line: ${anchorToken}`);
  const startIndex = findTokenStartIndex(text, anchorToken, occurrenceIndex);
  const anchorPosition = positionAt(text, startIndex);
  const identifierOffset = anchorToken.lastIndexOf(identifier);

  assert.notEqual(identifierOffset, -1, `identifier '${identifier}' must exist in anchor token: ${anchorToken}`);

  return {
    lineIndex: anchorPosition.line,
    startCharacter: anchorPosition.character + identifierOffset
  };
}

function findTokenStartIndex(source: string, token: string, occurrenceIndex = 0): number {
  assert.ok(occurrenceIndex >= 0, `occurrence index must be non-negative: ${occurrenceIndex}`);

  let startIndex = -1;
  let searchFromIndex = 0;

  for (let index = 0; index <= occurrenceIndex; index += 1) {
    startIndex = source.indexOf(token, searchFromIndex);
    if (startIndex === -1) {
      break;
    }

    searchFromIndex = startIndex + token.length;
  }

  assert.notEqual(startIndex, -1, `token occurrence not found in document: ${token} [${occurrenceIndex}]`);
  return startIndex;
}

function positionAt(text: string, offset: number): { line: number; character: number } {
  const lines = text.split("\n");
  let remaining = offset;

  for (let line = 0; line < lines.length; line += 1) {
    const lineLength = lines[line]?.length ?? 0;
    if (remaining <= lineLength) {
      return { character: remaining, line };
    }

    remaining -= lineLength + 1;
  }

  return { character: 0, line: lines.length - 1 };
}

function getCodeAction(actions: readonly vscode.CodeAction[], title: string): vscode.CodeAction | undefined {
  return actions.find((action) => action.title === title);
}

function hasCodeAction(actions: readonly vscode.CodeAction[], title: string): boolean {
  return getCodeAction(actions, title) !== undefined;
}

function isCodeAction(action: vscode.CodeAction | vscode.Command): action is vscode.CodeAction {
  return "title" in action && "edit" in action;
}
