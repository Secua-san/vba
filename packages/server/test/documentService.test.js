const assert = require("node:assert/strict");
const { mkdtempSync, mkdirSync, rmSync, utimesSync, writeFileSync } = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");
const { pathToFileURL } = require("node:url");
const { createDocumentService } = require("../dist/index.js");
const { markIndexedAccessPathSegment, resolveBuiltinMemberOwnerFromRootType } = require("../../core/dist/index.js");

test("document service analyzes text and exposes LSP-ready data", () => {
  const service = createDocumentService();
  const state = service.analyzeText(
    "file:///C:/temp/Sample.bas",
    "vba",
    1,
    `Attribute VB_Name = "Sample"
Option Explicit

Public Sub Demo()
    Dim message As String
    message = "Hello"
End Sub`
  );

  assert.equal(state.analysis.module.name, "Sample");
  const messageCompletion = service
    .getCompletionSymbols(state.uri, { character: 4, line: 5 })
    .find((resolution) => resolution.symbol.name === "message");

  assert.equal(messageCompletion?.typeName, "String");
  assert.equal(service.getDefinition(state.uri, { character: 5, line: 5 })?.symbol.name, "message");
  assert.ok(service.getDocumentSymbols(state.uri)[0]?.children?.some((symbol) => symbol.name === "Demo"));
  assert.deepEqual(
    service
      .getReferences(state.uri, { character: 5, line: 5 }, true)
      .map((reference) => `${reference.uri}:${reference.range.start.line}:${reference.range.start.character}`),
    [`${state.uri}:4:8`, `${state.uri}:5:4`]
  );
});

test("document service exposes built-in completion items from the reference index", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/BuiltInCompletion.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "BuiltInCompletion"
Option Explicit

Public Sub Demo()
    App
    xlA
End Sub`
  );

  const applicationCompletions = service.getCompletionSymbols(uri, { character: 7, line: 4 });
  const excelConstantCompletions = service.getCompletionSymbols(uri, { character: 7, line: 5 });
  const application = applicationCompletions.find((resolution) => resolution.symbol.name === "Application");
  const excelConstant = excelConstantCompletions.find((resolution) => resolution.symbol.name === "xlAll");

  assert.equal(application?.isBuiltIn, true);
  assert.equal(application?.moduleName.includes("Excel"), true);
  assert.equal(excelConstant?.isBuiltIn, true);
  assert.equal(excelConstant?.typeName, "Long");
});

test("document service exposes built-in member completion items from the reference index", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/BuiltInMemberCompletion.bas";
  const thisWorkbookUri = "file:///C:/temp/ThisWorkbook.cls";
  const sheet1Uri = "file:///C:/temp/Sheet1.cls";
  const chart1Uri = "file:///C:/temp/Chart1.cls";
  const text = `Attribute VB_Name = "BuiltInMemberCompletion"
Option Explicit

Public Sub Demo()
    Dim i As Long

    Debug.Print Application.
    Debug.Print WorksheetFunction.Su
    Debug.Print Application.WorksheetFunction.Su
    Debug.Print ActiveWorkbook.
    Debug.Print ThisWorkbook.
    Debug.Print Sheet1.
    Debug.Print Chart1.
    Debug.Print ActiveWorkbook.Worksheets.
    Debug.Print Worksheets(1).
    Debug.Print Worksheets("A(1)").
    Debug.Print Worksheets(i + 1).
    Debug.Print Worksheets(Array("Sheet1", "Sheet2")).
    Debug.Print ActiveWorkbook.Worksheets(1).
    Debug.Print ActiveWorkbook.Worksheets(GetIndex()).
    Debug.Print Application.ActiveCell.
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function`;

  service.analyzeText(
    thisWorkbookUri,
    "vba",
    1,
    `Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(
    sheet1Uri,
    "vba",
    1,
    `Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(
    chart1Uri,
    "vba",
    1,
    `Attribute VB_Name = "Chart1"
Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(uri, "vba", 1, text);

  const applicationMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Application."));
  const worksheetFunctionMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Su"));
  const chainedWorksheetFunctionMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Application.WorksheetFunction.Su")
  );
  const activeWorkbookMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "ActiveWorkbook."));
  const thisWorkbookMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "ThisWorkbook."));
  const sheet1Members = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1."));
  const chart1Members = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Chart1."));
  const workbookWorksheetsMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.Worksheets.")
  );
  const indexedWorksheetMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Worksheets(1)."));
  const indexedWorksheetStringMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Worksheets(\"A(1)\").")
  );
  const indexedWorksheetExpressionMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Worksheets(i + 1).")
  );
  const groupedWorksheetMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'Worksheets(Array("Sheet1", "Sheet2")).')
  );
  const chainedIndexedWorksheetMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.Worksheets(1).")
  );
  const chainedIndexedWorksheetFunctionMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.Worksheets(GetIndex()).")
  );
  const applicationActiveCellMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Application.ActiveCell.")
  );

  const worksheetFunctionProperty = applicationMembers.find((resolution) => resolution.symbol.name === "WorksheetFunction");
  const activeCellProperty = applicationMembers.find((resolution) => resolution.symbol.name === "ActiveCell");
  const worksheetFunctionSum = worksheetFunctionMembers.find((resolution) => resolution.symbol.name === "Sum");
  const chainedWorksheetFunctionSum = chainedWorksheetFunctionMembers.find((resolution) => resolution.symbol.name === "Sum");
  const activeWorkbookSaveAs = activeWorkbookMembers.find((resolution) => resolution.symbol.name === "SaveAs");
  const activeWorkbookWorksheets = activeWorkbookMembers.find((resolution) => resolution.symbol.name === "Worksheets");
  const thisWorkbookSaveAs = thisWorkbookMembers.find((resolution) => resolution.symbol.name === "SaveAs");
  const sheet1Evaluate = sheet1Members.find((resolution) => resolution.symbol.name === "Evaluate");
  const sheet1SaveAs = sheet1Members.find((resolution) => resolution.symbol.name === "SaveAs");
  const chart1ChartArea = chart1Members.find((resolution) => resolution.symbol.name === "ChartArea");
  const chart1SetSourceData = chart1Members.find((resolution) => resolution.symbol.name === "SetSourceData");
  const workbookWorksheetsCount = workbookWorksheetsMembers.find((resolution) => resolution.symbol.name === "Count");
  const indexedWorksheetEvaluate = indexedWorksheetMembers.find((resolution) => resolution.symbol.name === "Evaluate");
  const indexedWorksheetSaveAs = indexedWorksheetMembers.find((resolution) => resolution.symbol.name === "SaveAs");
  const indexedWorksheetStringEvaluate = indexedWorksheetStringMembers.find(
    (resolution) => resolution.symbol.name === "Evaluate"
  );
  const indexedWorksheetExpressionSaveAs = indexedWorksheetExpressionMembers.find(
    (resolution) => resolution.symbol.name === "SaveAs"
  );
  const groupedWorksheetCount = groupedWorksheetMembers.find((resolution) => resolution.symbol.name === "Count");
  const chainedIndexedWorksheetExport = chainedIndexedWorksheetMembers.find(
    (resolution) => resolution.symbol.name === "ExportAsFixedFormat"
  );
  const chainedIndexedWorksheetFunctionCount = chainedIndexedWorksheetFunctionMembers.find(
    (resolution) => resolution.symbol.name === "Count"
  );
  const applicationActiveCellAddress = applicationActiveCellMembers.find((resolution) => resolution.symbol.name === "Address");

  assert.equal(worksheetFunctionProperty?.isBuiltIn, true);
  assert.equal(worksheetFunctionProperty?.moduleName, "Excel Application property");
  assert.equal(worksheetFunctionProperty?.typeName, "WorksheetFunction");
  assert.equal(worksheetFunctionProperty?.documentation?.includes("excel.application.worksheetfunction"), true);
  assert.equal(activeCellProperty?.moduleName, "Excel Application property");
  assert.equal(worksheetFunctionSum?.moduleName, "Excel WorksheetFunction method");
  assert.equal(worksheetFunctionSum?.documentation?.includes("excel.worksheetfunction.sum"), true);
  assert.equal(chainedWorksheetFunctionSum?.moduleName, "Excel WorksheetFunction method");
  assert.equal(activeWorkbookSaveAs?.moduleName, "Excel Workbook method");
  assert.equal(activeWorkbookSaveAs?.documentation?.includes("excel.workbook.saveas"), true);
  assert.equal(activeWorkbookWorksheets?.typeName, "Worksheets");
  assert.equal(thisWorkbookSaveAs?.moduleName, "Excel Workbook method");
  assert.equal(thisWorkbookSaveAs?.documentation?.includes("excel.workbook.saveas"), true);
  assert.equal(sheet1Evaluate?.moduleName, "Excel Worksheet method");
  assert.equal(sheet1Evaluate?.documentation?.includes("excel.worksheet.evaluate"), true);
  assert.equal(sheet1SaveAs?.documentation?.includes("excel.worksheet.saveas"), true);
  assert.equal(chart1ChartArea?.moduleName, "Excel Chart property");
  assert.equal(chart1ChartArea?.documentation?.includes("excel.chart.chartarea"), true);
  assert.equal(chart1SetSourceData?.moduleName, "Excel Chart method");
  assert.equal(chart1SetSourceData?.documentation?.includes("excel.chart.setsourcedata"), true);
  assert.equal(workbookWorksheetsCount?.moduleName, "Excel Worksheets property");
  assert.equal(indexedWorksheetEvaluate?.moduleName, "Excel Worksheet method");
  assert.equal(indexedWorksheetSaveAs?.documentation?.includes("excel.worksheet.saveas"), true);
  assert.equal(indexedWorksheetStringEvaluate?.documentation?.includes("excel.worksheet.evaluate"), true);
  assert.equal(indexedWorksheetExpressionSaveAs?.moduleName, "Excel Worksheet method");
  assert.equal(groupedWorksheetCount?.moduleName, "Excel Worksheets property");
  assert.equal(groupedWorksheetMembers.some((resolution) => resolution.symbol.name === "Evaluate"), false);
  assert.equal(chainedIndexedWorksheetExport?.documentation?.includes("excel.worksheet.exportasfixedformat"), true);
  assert.equal(chainedIndexedWorksheetFunctionCount?.moduleName, "Excel Worksheets property");
  assert.equal(chainedIndexedWorksheetFunctionMembers.some((resolution) => resolution.symbol.name === "ExportAsFixedFormat"), false);
  assert.equal(applicationActiveCellAddress?.documentation?.includes("excel.range.address"), true);
});

test("document service exposes DialogSheet common callable members through DialogSheets roots", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/DialogSheetBuiltIn.bas";
  const text = `Attribute VB_Name = "DialogSheetBuiltIn"
Option Explicit

Public Sub Demo()
    Debug.Print DialogSheets.
    Debug.Print DialogSheets(1).
    Debug.Print DialogSheets("Dialog1").
    Debug.Print DialogSheets(Array("Dialog1", "Dialog2")).
    Debug.Print DialogSheets(1).SaveAs
    Debug.Print DialogSheets(1).Evaluate
    Debug.Print DialogSheets.Item(1).
End Sub`;

  service.analyzeText(uri, "vba", 1, text);

  const dialogSheetsCollectionMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets."));
  const indexedDialogSheetMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1)."));
  const namedDialogSheetMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, 'DialogSheets("Dialog1").'));
  const groupedDialogSheetsMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(Array("Dialog1", "Dialog2")).')
  );
  const itemDialogSheetMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets.Item(1).")
  );
  const dialogSheetSaveAsHover = service.getHover(uri, findPositionAfterTokenInText(text, "DialogSheets(1).SaveA"));
  const tokens = service.getSemanticTokens(uri);

  const dialogSheetsCount = dialogSheetsCollectionMembers.find((resolution) => resolution.symbol.name === "Count");
  const dialogSheetSaveAs = indexedDialogSheetMembers.find((resolution) => resolution.symbol.name === "SaveAs");
  const dialogSheetEvaluate = indexedDialogSheetMembers.find((resolution) => resolution.symbol.name === "Evaluate");
  const namedDialogSheetActivate = namedDialogSheetMembers.find((resolution) => resolution.symbol.name === "Activate");
  const groupedDialogSheetsCount = groupedDialogSheetsMembers.find((resolution) => resolution.symbol.name === "Count");
  const itemDialogSheetSaveAs = itemDialogSheetMembers.find((resolution) => resolution.symbol.name === "SaveAs");

  assert.equal(dialogSheetsCount?.moduleName, "Excel DialogSheets property");
  assert.equal(dialogSheetSaveAs?.moduleName, "Excel DialogSheet method");
  assert.equal(dialogSheetSaveAs?.documentation?.includes("dialogsheet.saveas"), true);
  assert.equal(dialogSheetEvaluate?.documentation?.includes("dialogsheet.evaluate"), true);
  assert.equal(namedDialogSheetActivate?.documentation?.includes("dialogsheet.activate"), true);
  assert.equal(groupedDialogSheetsCount?.moduleName, "Excel DialogSheets property");
  assert.equal(itemDialogSheetSaveAs?.documentation?.includes("dialogsheet.saveas"), true);
  assert.equal(groupedDialogSheetsMembers.some((resolution) => resolution.symbol.name === "SaveAs"), false);
  assert.equal(dialogSheetSaveAsHover?.contents.includes("SaveAs(Filename, FileFormat, Password, ..., Local)"), true);
  assert.equal(dialogSheetSaveAsHover?.contents.includes("microsoft.office.interop.excel.dialogsheet.saveas"), true);
  assert.equal(
    tokens.some(
      (entry) =>
        entry.range.start.line === 8 &&
        entry.range.start.character === 32 &&
        entry.range.end.character === 38 &&
        entry.type === "function"
    ),
    true
  );
  assert.equal(
    tokens.some(
      (entry) =>
        entry.range.start.line === 4 &&
        entry.range.start.character === 16 &&
        entry.range.end.character === 28 &&
        entry.type === "type"
    ),
    true
  );
});

test("document service exposes OLEObject members through Worksheet and Chart OLEObjects roots", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/OleObjectBuiltIn.bas";
  const sheet1Uri = "file:///C:/temp/Sheet1.cls";
  const chart1Uri = "file:///C:/temp/Chart1.cls";
  const text = `Attribute VB_Name = "OleObjectBuiltIn"
Option Explicit

Public Sub Demo()
    Dim i As Long

    Debug.Print Sheet1.OLEObjects.
    Debug.Print Sheet1.OLEObjects(1).
    Debug.Print Sheet1.OLEObjects("CheckBox1").
    Debug.Print Sheet1.OLEObjects(i + 1).
    Debug.Print Sheet1.OLEObjects(GetIndex()).
    Debug.Print Sheet1.OLEObjects.Item(1).
    Debug.Print Sheet1.OLEObjects.Item("CheckBox1").
    Debug.Print Sheet1.OLEObjects.Item(i + 1).
    Debug.Print Sheet1.OLEObjects.Item(GetIndex()).
    Debug.Print Chart1.OLEObjects(1).
    Debug.Print Chart1.OLEObjects.Item(1).
    Call Sheet1.OLEObjects(1).Activate(
    Call Sheet1.OLEObjects(GetIndex()).Activate(
    Call Sheet1.OLEObjects.Item(1).Activate(
    Call Sheet1.OLEObjects.Item(GetIndex()).Activate(
    Debug.Print Sheet1.OLEObjects(1).Name
    Debug.Print Sheet1.OLEObjects.Item(1).Name
    Debug.Print Sheet1.OLEObjects(1).Object.
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function`;

  service.analyzeText(
    sheet1Uri,
    "vba",
    1,
    `Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(
    chart1Uri,
    "vba",
    1,
    `Attribute VB_Name = "Chart1"
Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(uri, "vba", 1, text);

  const sheetOleObjectsMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.OLEObjects."));
  const indexedOleObjectMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.OLEObjects(1)."));
  const namedOleObjectMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'Sheet1.OLEObjects("CheckBox1").')
  );
  const expressionOleObjectMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Sheet1.OLEObjects(i + 1).")
  );
  const functionOleObjectsMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Sheet1.OLEObjects(GetIndex()).")
  );
  const itemIndexedOleObjectMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Sheet1.OLEObjects.Item(1).")
  );
  const itemNamedOleObjectMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'Sheet1.OLEObjects.Item("CheckBox1").')
  );
  const itemExpressionOleObjectMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Sheet1.OLEObjects.Item(i + 1).")
  );
  const itemFunctionOleObjectsMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Sheet1.OLEObjects.Item(GetIndex()).")
  );
  const chartOleObjectMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Chart1.OLEObjects(1)."));
  const chartItemOleObjectMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Chart1.OLEObjects.Item(1).")
  );
  const activateSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Sheet1.OLEObjects(1).Activate("));
  const functionActivateSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Sheet1.OLEObjects(GetIndex()).Activate(")
  );
  const itemActivateSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Sheet1.OLEObjects.Item(1).Activate(")
  );
  const itemFunctionActivateSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Sheet1.OLEObjects.Item(GetIndex()).Activate(")
  );
  const nameHover = service.getHover(uri, findPositionAfterTokenInText(text, "Sheet1.OLEObjects(1).Nam"));
  const itemNameHover = service.getHover(uri, findPositionAfterTokenInText(text, "Sheet1.OLEObjects.Item(1).Nam"));
  const objectMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.OLEObjects(1).Object."));

  const sheetOleObjectsCount = sheetOleObjectsMembers.find((resolution) => resolution.symbol.name === "Count");
  const indexedOleObjectActivate = indexedOleObjectMembers.find((resolution) => resolution.symbol.name === "Activate");
  const indexedOleObjectName = indexedOleObjectMembers.find((resolution) => resolution.symbol.name === "Name");
  const namedOleObjectVisible = namedOleObjectMembers.find((resolution) => resolution.symbol.name === "Visible");
  const expressionOleObjectName = expressionOleObjectMembers.find((resolution) => resolution.symbol.name === "Name");
  const functionOleObjectsCount = functionOleObjectsMembers.find((resolution) => resolution.symbol.name === "Count");
  const itemIndexedOleObjectActivate = itemIndexedOleObjectMembers.find((resolution) => resolution.symbol.name === "Activate");
  const itemNamedOleObjectVisible = itemNamedOleObjectMembers.find((resolution) => resolution.symbol.name === "Visible");
  const itemExpressionOleObjectName = itemExpressionOleObjectMembers.find((resolution) => resolution.symbol.name === "Name");
  const itemFunctionOleObjectsCount = itemFunctionOleObjectsMembers.find((resolution) => resolution.symbol.name === "Count");
  const chartOleObjectName = chartOleObjectMembers.find((resolution) => resolution.symbol.name === "Name");
  const chartItemOleObjectName = chartItemOleObjectMembers.find((resolution) => resolution.symbol.name === "Name");

  assert.equal(sheetOleObjectsCount?.moduleName, "Excel OLEObjects property");
  assert.equal(indexedOleObjectActivate?.moduleName, "Excel OLEObject method");
  assert.equal(indexedOleObjectName?.documentation?.includes("excel.oleobject.name"), true);
  assert.equal(namedOleObjectVisible?.moduleName, "Excel OLEObject property");
  assert.equal(expressionOleObjectName?.moduleName, "Excel OLEObject property");
  assert.equal(functionOleObjectsCount?.moduleName, "Excel OLEObjects property");
  assert.equal(functionOleObjectsMembers.some((resolution) => resolution.symbol.name === "Activate"), false);
  assert.equal(itemIndexedOleObjectActivate?.moduleName, "Excel OLEObject method");
  assert.equal(itemNamedOleObjectVisible?.moduleName, "Excel OLEObject property");
  assert.equal(itemExpressionOleObjectName?.moduleName, "Excel OLEObject property");
  assert.equal(itemFunctionOleObjectsCount?.moduleName, "Excel OLEObjects property");
  assert.equal(itemFunctionOleObjectsMembers.some((resolution) => resolution.symbol.name === "Activate"), false);
  assert.equal(chartOleObjectName?.moduleName, "Excel OLEObject property");
  assert.equal(chartItemOleObjectName?.moduleName, "Excel OLEObject property");
  assert.equal(activateSignature?.label.includes("Activate()"), true);
  assert.equal(functionActivateSignature, undefined);
  assert.equal(itemActivateSignature?.label.includes("Activate()"), true);
  assert.equal(itemFunctionActivateSignature, undefined);
  assert.equal(nameHover?.contents.includes("OLEObject.Name"), true);
  assert.equal(nameHover?.contents.includes("excel.oleobject.name"), true);
  assert.equal(itemNameHover?.contents.includes("OLEObject.Name"), true);
  assert.equal(itemNameHover?.contents.includes("excel.oleobject.name"), true);
  assert.equal(objectMembers.some((resolution) => resolution.symbol.name === "Activate"), false);
});

test("document service exposes Shape members through Worksheet and Chart Shapes roots while limiting OLEFormat.Object promotion", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-shapes-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "book1");
  const moduleDirectory = path.join(bundleRoot, "modules");
  const thisWorkbookUri = pathToFileURL(path.join(bundleRoot, "ThisWorkbook.cls")).href;
  const sheet1Uri = pathToFileURL(path.join(bundleRoot, "Sheet1.cls")).href;
  const chart1Uri = pathToFileURL(path.join(bundleRoot, "Chart1.cls")).href;
  const uri = pathToFileURL(path.join(moduleDirectory, "ShapesBuiltIn.bas")).href;
  const text = `Attribute VB_Name = "ShapesBuiltIn"
Option Explicit

Public Sub Demo()
    Dim i As Long

    Debug.Print Sheet1.Shapes.
    Debug.Print Sheet1.Shapes(1).
    Debug.Print Sheet1.Shapes("CheckBox1").
    Debug.Print Sheet1.Shapes(i + 1).
    Debug.Print Sheet1.Shapes.Item(1).
    Debug.Print Sheet1.Shapes.Item("CheckBox1").
    Debug.Print Sheet1.Shapes.Item(i + 1).
    Debug.Print Chart1.Shapes(1).
    Debug.Print Chart1.Shapes.Item(1).
    Debug.Print Sheet1.Shapes("CheckBox1").Name
    Debug.Print Sheet1.Shapes("CheckBox1").OLEFormat.
    Debug.Print Sheet1.Shapes.Item("CheckBox1").OLEFormat.
    Debug.Print Sheet1.Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Sheet1.Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Call Sheet1.Shapes("CheckBox1").OLEFormat.Object.Select(
    Call Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Sheet1.Shapes(1).OLEFormat.Object.
    Debug.Print Sheet1.Shapes.Item(1).OLEFormat.Object.
    Debug.Print Chart1.Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Chart1.Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Sheet1.Shapes.Range(Array("CheckBox1")).OLEFormat.Object.
    Debug.Print Sheet1.Shapes("PlainShape").OLEFormat.Object.
    Debug.Print Sheet1.Shapes.Item("PlainShape").OLEFormat.Object.
    Debug.Print Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Worksheets("Sheet1").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Call ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Sheet1.Shapes("CheckBox1").OLEFormat.Object(1).
    Debug.Print ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Call ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
End Sub`;

  mkdirSync(moduleDirectory, { recursive: true });
  writeWorksheetControlMetadataSidecar(bundleRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [
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
        sheetName: "Sheet One",
        status: "supported"
      }
    ],
    version: 1,
    workbook: {
      name: "book1.xlsm",
      sourceKind: "openxml-package"
    }
  });
  writeWorkbookBindingManifest(bundleRoot, {
    artifact: "workbook-binding-manifest",
    bindingKind: "active-workbook-fullname",
    version: 1,
    workbook: {
      fullName: "C:\\Fixtures\\book1.xlsm",
      isAddIn: false,
      name: "book1.xlsm",
      path: "C:\\Fixtures",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });

    service.analyzeText(
      thisWorkbookUri,
      "vba",
      1,
      `Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(
      sheet1Uri,
      "vba",
      1,
      `Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(
      chart1Uri,
      "vba",
      1,
      `Attribute VB_Name = "Chart1"
Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(uri, "vba", 1, text);

    const shapesCollectionMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.Shapes."));
    const indexedShapeMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.Shapes(1)."));
    const namedShapeMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, 'Sheet1.Shapes("CheckBox1").'));
    const dynamicShapeMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.Shapes(i + 1)."));
    const itemIndexedShapeMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.Shapes.Item(1)."));
    const itemNamedShapeMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes.Item("CheckBox1").')
    );
    const itemDynamicShapeMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, "Sheet1.Shapes.Item(i + 1).")
    );
    const chartShapeMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Chart1.Shapes(1)."));
    const chartItemShapeMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, "Chart1.Shapes.Item(1).")
    );
    const oleFormatMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes("CheckBox1").OLEFormat.')
    );
    const itemOleFormatMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes.Item("CheckBox1").OLEFormat.')
    );
    const objectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes("CheckBox1").OLEFormat.Object.')
    );
    const itemObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.')
    );
    const indexedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, "Sheet1.Shapes(1).OLEFormat.Object.")
    );
    const itemIndexedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, "Sheet1.Shapes.Item(1).OLEFormat.Object.")
    );
    const chartNamedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Chart1.Shapes("CheckBox1").OLEFormat.Object.')
    );
    const chartItemNamedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Chart1.Shapes.Item("CheckBox1").OLEFormat.Object.')
    );
    const groupedShapeRangeObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes.Range(Array("CheckBox1")).OLEFormat.Object.')
    );
    const unmatchedShapeObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes("PlainShape").OLEFormat.Object.')
    );
    const itemUnmatchedShapeObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes.Item("PlainShape").OLEFormat.Object.')
    );
    const worksheetNameRootObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.')
    );
    const worksheetNameRootItemObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Worksheets("Sheet1").Shapes.Item("CheckBox1").OLEFormat.Object.')
    );
    const thisWorkbookWorksheetNameRootObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.')
    );
    const thisWorkbookWorksheetNameRootItemObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.')
    );
    const thisWorkbookWorksheetIndexedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.')
    );
    const activeWorkbookWorksheetNameRootObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.')
    );
    const thisWorkbookWorksheetCodeNameObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.')
    );
    const indexedObjectCallMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes("CheckBox1").OLEFormat.Object(1).')
    );
    const nameHover = service.getHover(uri, findPositionAfterTokenInText(text, 'Sheet1.Shapes("CheckBox1").Nam'));
    const namedValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes("CheckBox1").OLEFormat.Object.Valu')
    );
    const itemNamedValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.Valu')
    );
    const thisWorkbookNamedValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu')
    );
    const thisWorkbookItemNamedValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Valu')
    );
    const namedSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes("CheckBox1").OLEFormat.Object.Select(')
    );
    const itemNamedSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object.Select(')
    );
    const thisWorkbookNamedSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(')
    );
    const thisWorkbookItemNamedSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(')
    );
    const tokens = service.getSemanticTokens(uri);

    const indexedShapeName = indexedShapeMembers.find((resolution) => resolution.symbol.name === "Name");
    const itemIndexedShapeName = itemIndexedShapeMembers.find((resolution) => resolution.symbol.name === "Name");
    const chartShapeName = chartShapeMembers.find((resolution) => resolution.symbol.name === "Name");
    const chartItemShapeName = chartItemShapeMembers.find((resolution) => resolution.symbol.name === "Name");
    const oleFormatProgId = oleFormatMembers.find((resolution) => resolution.symbol.name === "progID");
    const itemOleFormatProgId = itemOleFormatMembers.find((resolution) => resolution.symbol.name === "progID");
    const namedObjectValue = objectMembers.find((resolution) => resolution.symbol.name === "Value");
    const namedObjectSelect = objectMembers.find((resolution) => resolution.symbol.name === "Select");
    const itemNamedObjectValue = itemObjectMembers.find((resolution) => resolution.symbol.name === "Value");
    const itemNamedObjectSelect = itemObjectMembers.find((resolution) => resolution.symbol.name === "Select");

    assert.equal(shapesCollectionMembers.some((resolution) => resolution.symbol.name === "Count"), true);
    assert.equal(shapesCollectionMembers.some((resolution) => resolution.symbol.name === "Name"), false);
    assert.equal(indexedShapeName?.moduleName, "Excel Shape property");
    assert.equal(namedShapeMembers.some((resolution) => resolution.symbol.name === "Name"), true);
    assert.equal(dynamicShapeMembers.some((resolution) => resolution.symbol.name === "Name"), true);
    assert.equal(itemIndexedShapeName?.moduleName, "Excel Shape property");
    assert.equal(itemNamedShapeMembers.some((resolution) => resolution.symbol.name === "Name"), true);
    assert.equal(itemDynamicShapeMembers.some((resolution) => resolution.symbol.name === "Name"), true);
    assert.equal(chartShapeName?.moduleName, "Excel Shape property");
    assert.equal(chartItemShapeName?.moduleName, "Excel Shape property");
    assert.equal(oleFormatProgId?.moduleName, "Excel OLEFormat property");
    assert.equal(itemOleFormatProgId?.moduleName, "Excel OLEFormat property");
    assert.equal(namedObjectValue?.moduleName.includes("CheckBox property"), true);
    assert.equal(namedObjectValue?.documentation?.includes("microsoft.office.interop.excel.checkbox.value"), true);
    assert.equal(namedObjectSelect?.moduleName.includes("CheckBox method"), true);
    assert.equal(objectMembers.some((resolution) => resolution.symbol.name === "Activate"), false);
    assert.equal(itemNamedObjectValue?.moduleName.includes("CheckBox property"), true);
    assert.equal(itemNamedObjectSelect?.moduleName.includes("CheckBox method"), true);
    assert.equal(itemObjectMembers.some((resolution) => resolution.symbol.name === "Activate"), false);
    assert.equal(indexedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(itemIndexedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(chartNamedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(chartItemNamedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(groupedShapeRangeObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(unmatchedShapeObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(itemUnmatchedShapeObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(worksheetNameRootObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(worksheetNameRootItemObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(thisWorkbookWorksheetNameRootObjectMembers.some((resolution) => resolution.symbol.name === "Value"), true);
    assert.equal(thisWorkbookWorksheetNameRootItemObjectMembers.some((resolution) => resolution.symbol.name === "Value"), true);
    assert.equal(thisWorkbookWorksheetIndexedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(activeWorkbookWorksheetNameRootObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(thisWorkbookWorksheetCodeNameObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(indexedObjectCallMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(nameHover?.contents.includes("Shape.Name"), true);
    assert.equal(nameHover?.contents.includes("excel.shape.name"), true);
    assert.equal(namedValueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(namedValueHover?.contents.includes("microsoft.office.interop.excel.checkbox.value"), true);
    assert.equal(itemNamedValueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(thisWorkbookNamedValueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(thisWorkbookItemNamedValueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(namedSelectSignature?.label, "Select(Replace) As Object");
    assert.equal(itemNamedSelectSignature?.label, "Select(Replace) As Object");
    assert.equal(thisWorkbookNamedSelectSignature?.label, "Select(Replace) As Object");
    assert.equal(thisWorkbookItemNamedSelectSignature?.label, "Select(Replace) As Object");
    assertSemanticToken(text, tokens, 15, "Name", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 16, "OLEFormat", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 20, "Value", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 22, "Select", { modifiers: [], type: "function" });
    assertSemanticToken(text, tokens, 38, "Value", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 40, "Select", { modifiers: [], type: "function" });

    service.setActiveWorkbookIdentitySnapshot({
      identity: {
        fullName: "c:/fixtures/BOOK1.xlsm",
        isAddin: false,
        name: "book1.xlsm",
        path: "c:/fixtures"
      },
      observedAt: "2026-03-21T00:00:00.000Z",
      providerKind: "excel-active-workbook",
      state: "available",
      version: 1
    });

    const activeWorkbookBoundObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.')
    );
    const activeWorkbookBoundValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu')
    );
    const activeWorkbookBoundSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(')
    );

    assert.equal(activeWorkbookBoundObjectMembers.some((resolution) => resolution.symbol.name === "Value"), true);
    assert.equal(activeWorkbookBoundObjectMembers.some((resolution) => resolution.symbol.name === "Delete"), false);
    assert.equal(activeWorkbookBoundValueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(activeWorkbookBoundSelectSignature?.label, "Select(Replace) As Object");

    service.setActiveWorkbookIdentitySnapshot({
      observedAt: "2026-03-21T00:01:00.000Z",
      providerKind: "excel-active-workbook",
      reason: "no-active-workbook",
      state: "unavailable",
      version: 1
    });

    const activeWorkbookUnavailableObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.')
    );
    const activeWorkbookUnavailableValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu')
    );
    const activeWorkbookUnavailableSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(')
    );

    assert.equal(activeWorkbookUnavailableObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(activeWorkbookUnavailableValueHover, undefined);
    assert.equal(activeWorkbookUnavailableSelectSignature, undefined);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service resolves OLEObject.Object through worksheet control metadata sidecar only for named worksheet selectors", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "book1");
  const moduleDirectory = path.join(bundleRoot, "modules");
  const thisWorkbookUri = pathToFileURL(path.join(bundleRoot, "ThisWorkbook.cls")).href;
  const sheet1Uri = pathToFileURL(path.join(bundleRoot, "Sheet1.cls")).href;
  const chart1Uri = pathToFileURL(path.join(bundleRoot, "Chart1.cls")).href;
  const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Dim i As Long

    Debug.Print Sheet1.OLEObjects("CheckBox1").Object.
    Debug.Print Sheet1.OLEObjects.Item("CheckBox1").Object.
    Debug.Print Sheet1.OLEObjects(1).Object.
    Debug.Print Sheet1.OLEObjects(i + 1).Object.
    Debug.Print Chart1.OLEObjects("CheckBox1").Object.
    Debug.Print ActiveSheet.OLEObjects("CheckBox1").Object.
    Debug.Print Sheet1.OLEObjects("CheckBox1").Object.Value
    Debug.Print Sheet1.OLEObjects.Item("CheckBox1").Object.Value
    Debug.Print ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.
    Debug.Print ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Call Sheet1.OLEObjects("CheckBox1").Object.Select(
    Call Sheet1.OLEObjects.Item("CheckBox1").Object.Select(
    Call ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Call Chart1.OLEObjects("CheckBox1").Object.Select(
    Debug.Print ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
End Sub`;

  mkdirSync(moduleDirectory, { recursive: true });
  writeWorksheetControlMetadataSidecar(bundleRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [
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
        sheetName: "Sheet One",
        status: "supported"
      },
      {
        controls: [
          {
            codeName: "chkChart",
            controlType: "CheckBox",
            progId: "Forms.CheckBox.1",
            shapeId: 8,
            shapeName: "CheckBox1"
          }
        ],
        ownerKind: "chartsheet",
        sheetCodeName: "Chart1",
        sheetName: "Chart1",
        status: "supported"
      }
    ],
    version: 1,
    workbook: {
      name: "book1.xlsm",
      sourceKind: "openxml-package"
    }
  });
  writeWorkbookBindingManifest(bundleRoot, {
    artifact: "workbook-binding-manifest",
    bindingKind: "active-workbook-fullname",
    version: 1,
    workbook: {
      fullName: "C:\\Fixtures\\book1.xlsm",
      isAddIn: false,
      name: "book1.xlsm",
      path: "C:\\Fixtures",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });

    service.analyzeText(
      thisWorkbookUri,
      "vba",
      1,
      `Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(
      sheet1Uri,
      "vba",
      1,
      `Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(
      chart1Uri,
      "vba",
      1,
      `Attribute VB_Name = "Chart1"
Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(uri, "vba", 1, text);

    const namedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.OLEObjects("CheckBox1").Object.')
    );
    const itemNamedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.OLEObjects.Item("CheckBox1").Object.')
    );
    const indexedObjectMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.OLEObjects(1).Object."));
    const dynamicObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, "Sheet1.OLEObjects(i + 1).Object.")
    );
    const chartObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'Chart1.OLEObjects("CheckBox1").Object.')
    );
    const activeSheetObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ActiveSheet.OLEObjects("CheckBox1").Object.')
    );
    const thisWorkbookNamedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.')
    );
    const thisWorkbookItemNamedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.')
    );
    const thisWorkbookIndexedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.')
    );
    const activeWorkbookNamedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.')
    );
    const thisWorkbookCodeNameObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.')
    );
    const namedValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.OLEObjects("CheckBox1").Object.Valu')
    );
    const itemNamedValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.OLEObjects.Item("CheckBox1").Object.Valu')
    );
    const thisWorkbookNamedValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu')
    );
    const thisWorkbookItemNamedValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu')
    );
    const namedSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.OLEObjects("CheckBox1").Object.Select(')
    );
    const itemNamedSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'Sheet1.OLEObjects.Item("CheckBox1").Object.Select(')
    );
    const thisWorkbookNamedSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(')
    );
    const thisWorkbookItemNamedSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'ThisWorkbook.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(')
    );
    const chartSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'Chart1.OLEObjects("CheckBox1").Object.Select(')
    );
    const tokens = service.getSemanticTokens(uri);

    const namedValue = namedObjectMembers.find((resolution) => resolution.symbol.name === "Value");
    const namedSelect = namedObjectMembers.find((resolution) => resolution.symbol.name === "Select");
    const itemNamedValue = itemNamedObjectMembers.find((resolution) => resolution.symbol.name === "Value");
    const itemNamedSelect = itemNamedObjectMembers.find((resolution) => resolution.symbol.name === "Select");

    assert.equal(namedValue?.moduleName.includes("CheckBox property"), true);
    assert.equal(namedValue?.documentation?.includes("microsoft.office.interop.excel.checkbox.value"), true);
    assert.equal(namedSelect?.moduleName.includes("CheckBox method"), true);
    assert.equal(namedObjectMembers.some((resolution) => resolution.symbol.name === "Activate"), false);
    assert.equal(itemNamedValue?.moduleName.includes("CheckBox property"), true);
    assert.equal(itemNamedSelect?.moduleName.includes("CheckBox method"), true);
    assert.equal(itemNamedObjectMembers.some((resolution) => resolution.symbol.name === "Activate"), false);
    assert.equal(indexedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(dynamicObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(chartObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(activeSheetObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(thisWorkbookNamedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), true);
    assert.equal(thisWorkbookItemNamedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), true);
    assert.equal(thisWorkbookIndexedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(activeWorkbookNamedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(thisWorkbookCodeNameObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(namedValueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(namedValueHover?.contents.includes("microsoft.office.interop.excel.checkbox.value"), true);
    assert.equal(itemNamedValueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(thisWorkbookNamedValueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(thisWorkbookItemNamedValueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(namedSelectSignature?.label, "Select(Replace) As Object");
    assert.equal(itemNamedSelectSignature?.label, "Select(Replace) As Object");
    assert.equal(thisWorkbookNamedSelectSignature?.label, "Select(Replace) As Object");
    assert.equal(thisWorkbookItemNamedSelectSignature?.label, "Select(Replace) As Object");
    assert.equal(chartSelectSignature, undefined);
    assertSemanticToken(text, tokens, 13, "Value", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 20, "Value", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 22, "Select", { modifiers: [], type: "function" });
    assertSemanticToken(text, tokens, 24, "Select", { modifiers: [], type: "function" });

    service.setActiveWorkbookIdentitySnapshot({
      identity: {
        fullName: "c:/fixtures/BOOK1.xlsm",
        isAddin: false,
        name: "book1.xlsm",
        path: "c:/fixtures"
      },
      observedAt: "2026-03-21T00:00:00.000Z",
      providerKind: "excel-active-workbook",
      state: "available",
      version: 1
    });

    const activeWorkbookBoundObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.')
    );
    const activeWorkbookBoundValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu')
    );
    const activeWorkbookBoundSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(')
    );

    assert.equal(activeWorkbookBoundObjectMembers.some((resolution) => resolution.symbol.name === "Value"), true);
    assert.equal(activeWorkbookBoundObjectMembers.some((resolution) => resolution.symbol.name === "Activate"), false);
    assert.equal(activeWorkbookBoundValueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(activeWorkbookBoundSelectSignature?.label, "Select(Replace) As Object");

    service.setActiveWorkbookIdentitySnapshot({
      identity: {
        fullName: "C:\\Fixtures\\OtherBook.xlsm",
        isAddin: false,
        name: "OtherBook.xlsm",
        path: "C:\\Fixtures"
      },
      observedAt: "2026-03-21T00:01:00.000Z",
      providerKind: "excel-active-workbook",
      state: "available",
      version: 1
    });

    const activeWorkbookMismatchedObjectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.')
    );
    const activeWorkbookMismatchedValueHover = service.getHover(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu')
    );
    const activeWorkbookMismatchedSelectSignature = service.getSignatureHelp(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(')
    );

    assert.equal(activeWorkbookMismatchedObjectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(activeWorkbookMismatchedValueHover, undefined);
    assert.equal(activeWorkbookMismatchedSelectSignature, undefined);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service resolves OLEObject.Object against the root document module sidecar", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleARoot = path.join(workspaceRoot, "bundle-a");
  const bundleBRoot = path.join(workspaceRoot, "bundle-b");
  const bundleBModuleDirectory = path.join(bundleBRoot, "modules");
  const sheetAUri = pathToFileURL(path.join(bundleARoot, "SheetA.cls")).href;
  const uri = pathToFileURL(path.join(bundleBModuleDirectory, "Module1.bas")).href;
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print SheetA.OLEObjects("Control1").Object.
End Sub`;

  mkdirSync(bundleBModuleDirectory, { recursive: true });
  writeWorksheetControlMetadataSidecar(bundleARoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [
      {
        controls: [
          {
            codeName: "chkFinished",
            controlType: "CheckBox",
            progId: "Forms.CheckBox.1",
            shapeId: 3,
            shapeName: "Control1"
          }
        ],
        ownerKind: "worksheet",
        sheetCodeName: "SheetA",
        sheetName: "SheetA",
        status: "supported"
      }
    ],
    version: 1,
    workbook: {
      name: "bundle-a.xlsm",
      sourceKind: "openxml-package"
    }
  });
  writeWorksheetControlMetadataSidecar(bundleBRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [
      {
        controls: [
          {
            codeName: "optFinished",
            controlType: "OptionButton",
            progId: "Forms.OptionButton.1",
            shapeId: 3,
            shapeName: "Control1"
          }
        ],
        ownerKind: "worksheet",
        sheetCodeName: "SheetA",
        sheetName: "SheetA",
        status: "supported"
      }
    ],
    version: 1,
    workbook: {
      name: "bundle-b.xlsm",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });

    service.analyzeText(
      sheetAUri,
      "vba",
      1,
      `Attribute VB_Name = "SheetA"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(uri, "vba", 1, text);

    const objectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'SheetA.OLEObjects("Control1").Object.')
    );
    const valueCompletion = objectMembers.find((resolution) => resolution.symbol.name === "Value");

    assert.equal(service.getState(sheetAUri)?.worksheetControlMetadata?.workbookName, "bundle-a.xlsm");
    assert.equal(service.getState(uri)?.worksheetControlMetadata?.workbookName, "bundle-b.xlsm");
    assert.equal(valueCompletion?.moduleName.includes("CheckBox"), true);
    assert.equal(valueCompletion?.moduleName.includes("OptionButton"), false);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service keeps ActiveWorkbook broad root closed when manifest and sidecar bundle roots differ", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const outerBundleRoot = path.join(workspaceRoot, "outer-bundle");
  const innerBundleRoot = path.join(outerBundleRoot, "inner-bundle");
  const moduleDirectory = path.join(innerBundleRoot, "modules");
  const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
End Sub`;

  mkdirSync(moduleDirectory, { recursive: true });
  writeWorkbookBindingManifest(outerBundleRoot, {
    artifact: "workbook-binding-manifest",
    bindingKind: "active-workbook-fullname",
    version: 1,
    workbook: {
      fullName: "C:\\Fixtures\\outer-bundle.xlsm",
      isAddIn: false,
      name: "outer-bundle.xlsm",
      path: "C:\\Fixtures",
      sourceKind: "openxml-package"
    }
  });
  writeWorksheetControlMetadataSidecar(innerBundleRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [
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
        sheetName: "Sheet One",
        status: "supported"
      }
    ],
    version: 1,
    workbook: {
      name: "inner-bundle.xlsm",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });

    service.setActiveWorkbookIdentitySnapshot({
      identity: {
        fullName: "C:\\Fixtures\\outer-bundle.xlsm",
        isAddin: false,
        name: "outer-bundle.xlsm",
        path: "C:\\Fixtures"
      },
      observedAt: "2026-03-21T00:00:00.000Z",
      providerKind: "excel-active-workbook",
      state: "available",
      version: 1
    });

    service.analyzeText(uri, "vba", 1, text);

    const state = service.getState(uri);
    const objectMembers = service.getCompletionSymbols(
      uri,
      findPositionAfterTokenInText(text, 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.')
    );

    assert.equal(state?.workbookBindingManifest?.bundleRoot, outerBundleRoot);
    assert.equal(state?.worksheetControlMetadata?.bundleRoot, innerBundleRoot);
    assert.equal(objectMembers.some((resolution) => resolution.symbol.name === "Value"), false);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service resolves unqualified worksheet broad roots only while active workbook binding matches", () => {
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Debug.Print Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Sheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print ActiveSheet.OLEObjects("CheckBox1").Object.Value
    Debug.Print Worksheets(1).OLEObjects("CheckBox1").Object.Value
    Debug.Print Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.Worksheets(1).OLEObjects("CheckBox1").Object.Value
    Debug.Print Application.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function`;
  const { service, uri, cleanup } = createWorksheetBroadRootFixture(text);
  const closedCompletionChecks = [
    ['Worksheets("Sheet One").OLEObjects("CheckBox1").Object.', "Value", 'unqualified Worksheets("Sheet One") は snapshot 未一致の間は broad root を開かない'],
    ['Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.', "Value", 'Application.Worksheets("Sheet One") は snapshot 未一致の間は broad root を開かない'],
    ['Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.', "Value", 'unqualified Worksheets("Sheet One").Shapes("CheckBox1") は snapshot 未一致の間は broad root を開かない'],
    ['Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.', "Value", 'Application.Worksheets("Sheet One").Shapes("CheckBox1") は snapshot 未一致の間は broad root を開かない']
  ];
  const matchedCompletionChecks = [
    ['Worksheets("Sheet One").OLEObjects("CheckBox1").Object.', "Value", "Activate", 'Worksheets("Sheet One").OLEObjects("CheckBox1").Object は control owner へ解決する'],
    ['Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.', "Value", "Activate", 'Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object は control owner へ解決する'],
    ['Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.', "Value", "Delete", 'Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object は control owner へ解決する'],
    ['Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.', "Value", "Delete", 'Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object は control owner へ解決する']
  ];
  const nonTargetCompletionChecks = [
    ['Sheets("Sheet One").OLEObjects("CheckBox1").Object.', "Value", 'Sheets root は broad root family の対象外を維持する'],
    ['ActiveSheet.OLEObjects("CheckBox1").Object.', "Value", 'ActiveSheet root は broad root family の対象外を維持する'],
    ['Worksheets(1).OLEObjects("CheckBox1").Object.', "Value", 'numeric selector の Worksheets root は broad root family の対象外を維持する'],
    ['Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.', "Value", 'dynamic selector の Worksheets root は broad root family の対象外を維持する'],
    ['Application.Worksheets(1).OLEObjects("CheckBox1").Object.', "Value", 'numeric selector の Application.Worksheets root は broad root family の対象外を維持する'],
    ['Application.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.', "Value", 'dynamic selector の Application.Worksheets root は broad root family の対象外を維持する']
  ];
  const hoverChecks = [
    ['Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu', "CheckBox.Value", 'Worksheets("Sheet One") の OLEObject.Object hover は control owner へ解決する'],
    ['Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu', "CheckBox.Value", 'Worksheets("Sheet One") の Shape.OLEFormat.Object hover は control owner へ解決する']
  ];
  const signatureChecks = [
    ['Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(', "Select(Replace) As Object", 'Application.Worksheets("Sheet One") の OLEObject.Object signature help は control owner へ解決する'],
    ['Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(', "Select(Replace) As Object", 'Application.Worksheets("Sheet One") の Shape.OLEFormat.Object signature help は control owner へ解決する']
  ];

  try {
    for (const [token, symbolName, message] of closedCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), false, message);
    }
    assert.equal(getHoverAfterToken(service, uri, text, 'Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu'), undefined);
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select('
      ),
      undefined
    );
    assert.equal(getHoverAfterToken(service, uri, text, 'Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu'), undefined);
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select('
      ),
      undefined
    );

    service.setActiveWorkbookIdentitySnapshot(createMatchedActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName, forbiddenSymbolName, message] of matchedCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), true, message);
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, forbiddenSymbolName), false, message);
    }
    for (const [token, symbolName, message] of nonTargetCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), false, message);
    }
    for (const [token, expectedFragment, message] of hoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token)?.contents.includes(expectedFragment), true, message);
    }
    for (const [token, expectedLabel, message] of signatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token)?.label, expectedLabel, message);
    }

    service.setActiveWorkbookIdentitySnapshot(createMismatchedActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName] of closedCompletionChecks) {
      assert.equal(
        hasCompletionSymbolAfterToken(service, uri, text, token, symbolName),
        false,
        `mismatch snapshot では ${token} broad root を開かない`
      );
    }
    for (const [token, symbolName] of nonTargetCompletionChecks) {
      assert.equal(
        hasCompletionSymbolAfterToken(service, uri, text, token, symbolName),
        false,
        `mismatch snapshot でも ${token} は broad root family の対象外を維持する`
      );
    }
    for (const [token] of hoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, `mismatch snapshot では ${token} hover を出さない`);
    }
    for (const [token] of signatureChecks) {
      assert.equal(
        getSignatureHelpAfterToken(service, uri, text, token),
        undefined,
        `mismatch snapshot では ${token} signature help を出さない`
      );
    }

    service.setActiveWorkbookIdentitySnapshot(createUnavailableActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName] of closedCompletionChecks) {
      assert.equal(
        hasCompletionSymbolAfterToken(service, uri, text, token, symbolName),
        false,
        `unavailable snapshot では ${token} broad root を開かない`
      );
    }
    for (const [token, symbolName] of nonTargetCompletionChecks) {
      assert.equal(
        hasCompletionSymbolAfterToken(service, uri, text, token, symbolName),
        false,
        `unavailable snapshot でも ${token} は broad root family の対象外を維持する`
      );
    }
    for (const [token] of hoverChecks) {
      assert.equal(
        getHoverAfterToken(service, uri, text, token),
        undefined,
        `unavailable snapshot では ${token} hover を出さない`
      );
    }
    for (const [token] of signatureChecks) {
      assert.equal(
        getSignatureHelpAfterToken(service, uri, text, token),
        undefined,
        `unavailable snapshot では ${token} signature help を出さない`
      );
    }
  } finally {
    cleanup();
  }
});

test("document service resolves unqualified worksheet broad root item selectors including root Item forms", () => {
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Application.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Call Application.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Debug.Print Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Call Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Debug.Print Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Call Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Call Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
End Sub`;
  const { service, uri, cleanup } = createWorksheetBroadRootFixture(text);
  const closedCompletionChecks = [
    ['Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.', "Value", 'Worksheets("Sheet One").OLEObjects.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない'],
    ['Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.', "Value", 'Application.Worksheets("Sheet One").Shapes.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない'],
    ['Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.', "Value", 'Worksheets.Item("Sheet One").OLEObjects("CheckBox1") は snapshot 未一致の間は broad root を開かない'],
    ['Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.', "Value", 'Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない'],
    ['Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.', "Value", 'Application.Worksheets.Item("Sheet One").Shapes("CheckBox1") は snapshot 未一致の間は broad root を開かない'],
    ['Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.', "Value", 'Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない']
  ];
  const matchedCompletionChecks = [
    ['Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.', "Value", "Activate", 'Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object は control owner へ解決する'],
    ['Application.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.', "Value", "Activate", 'Application.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object は control owner へ解決する'],
    ['Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.', "Value", "Delete", 'Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object は control owner へ解決する'],
    ['Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.', "Value", "Delete", 'Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object は control owner へ解決する'],
    ['Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.', "Value", "Activate", 'Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object は control owner へ解決する'],
    ['Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.', "Value", "Activate", 'Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object は control owner へ解決する'],
    ['Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.', "Value", "Delete", 'Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object は control owner へ解決する'],
    ['Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.', "Value", "Delete", 'Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object は control owner へ解決する']
  ];
  const matchedHoverChecks = [
    'Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu',
    'Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Valu',
    'Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Valu',
    'Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu'
  ];
  const matchedSignatureChecks = [
    'Application.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
    'Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
    'Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(',
    'Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select('
  ];

  try {
    for (const [token, symbolName, message] of closedCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), false, message);
    }
    for (const token of matchedHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, `snapshot 未一致では ${token} hover を出さない`);
    }
    for (const token of matchedSignatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token), undefined, `snapshot 未一致では ${token} signature help を出さない`);
    }

    service.setActiveWorkbookIdentitySnapshot(createMatchedActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName, forbiddenSymbolName, message] of matchedCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), true, message);
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, forbiddenSymbolName), false, message);
    }
    for (const token of matchedHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token)?.contents.includes("CheckBox.Value"), true);
    }
    for (const token of matchedSignatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token)?.label, "Select(Replace) As Object");
    }

    service.setActiveWorkbookIdentitySnapshot(createMismatchedActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName] of closedCompletionChecks) {
      assert.equal(
        hasCompletionSymbolAfterToken(service, uri, text, token, symbolName),
        false,
        `mismatch snapshot では ${token} broad root を開かない`
      );
    }
    for (const token of matchedHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, `mismatch snapshot では ${token} hover を出さない`);
    }
    for (const token of matchedSignatureChecks) {
      assert.equal(
        getSignatureHelpAfterToken(service, uri, text, token),
        undefined,
        `mismatch snapshot では ${token} signature help を出さない`
      );
    }

    service.setActiveWorkbookIdentitySnapshot(createUnavailableActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName] of closedCompletionChecks) {
      assert.equal(
        hasCompletionSymbolAfterToken(service, uri, text, token, symbolName),
        false,
        `unavailable snapshot では ${token} broad root を開かない`
      );
    }
    for (const token of matchedHoverChecks) {
      assert.equal(
        getHoverAfterToken(service, uri, text, token),
        undefined,
        `unavailable snapshot では ${token} hover を出さない`
      );
    }
    for (const token of matchedSignatureChecks) {
      assert.equal(
        getSignatureHelpAfterToken(service, uri, text, token),
        undefined,
        `unavailable snapshot では ${token} signature help を出さない`
      );
    }
  } finally {
    cleanup();
  }
});

test("document service resolves workbook-qualified worksheet root item selectors for OLEObject.Object", () => {
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Call ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Value
    Debug.Print ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value
    Call ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Select(
    Call ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Value
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Value
    Call ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Call ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.
    Debug.Print ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.
    Debug.Print ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Value
    Call ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Select(
    Call ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(
End Sub`;
  const { service, uri, cleanup } = createWorkbookQualifiedWorksheetRootFixture(text);
  const nonTargetCompletionChecks = [
    ['ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.', "Value", 'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'],
    ['ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.', "Value", 'ThisWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'],
    ['ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.', "Value", 'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'],
    ['ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.', "Value", 'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない']
  ];
  const staticCompletionChecks = [
    ['ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.', "Value", "Activate", 'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object は control owner へ解決する'],
    ['ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.', "Value", "Activate", 'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object は control owner へ解決する']
  ];
  const staticHoverChecks = [
    'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Valu',
    'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu'
  ];
  const staticSignatureChecks = [
    'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(',
    'ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select('
  ];
  const nonTargetHoverChecks = [
    ['ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Valu', 'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'],
    ['ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu', 'ThisWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'],
    ['ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Valu', 'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'],
    ['ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu', 'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない']
  ];
  const nonTargetSignatureChecks = [
    'ThisWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Select(',
    'ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(',
    'ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("CheckBox1").Object.Select(',
    'ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select('
  ];
  const closedCompletionChecks = [
    ['ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.', "Value", 'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1") は snapshot 未一致の間は broad root を開かない'],
    ['ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.', "Value", 'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない']
  ];
  const matchedCompletionChecks = [
    ['ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.', "Value", "Activate", 'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object は control owner へ解決する'],
    ['ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.', "Value", "Activate", 'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object は control owner へ解決する']
  ];
  const matchedHoverChecks = [
    'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Valu',
    'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu'
  ];
  const matchedSignatureChecks = [
    'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(',
    'ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select('
  ];
  const nonTargetSemanticChecks = [
    [10, "Value", 'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので semantic token を出さない'],
    [11, "Value", 'ThisWorkbook.Worksheets.Item(1) は numeric selector なので semantic token を出さない'],
    [12, "Select", 'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので semantic token を出さない'],
    [13, "Select", 'ThisWorkbook.Worksheets.Item(1) は numeric selector なので semantic token を出さない'],
    [22, "Value", 'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので semantic token を出さない'],
    [23, "Value", 'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので semantic token を出さない'],
    [24, "Select", 'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので semantic token を出さない'],
    [25, "Select", 'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので semantic token を出さない']
  ];

  try {
    for (const [token, symbolName, blockedSymbolName, message] of staticCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), true, message);
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, blockedSymbolName), false, message);
    }
    for (const token of staticHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token)?.contents.includes("CheckBox.Value"), true);
    }
    for (const token of staticSignatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token)?.label, "Select(Replace) As Object");
    }
    for (const [token, symbolName, message] of nonTargetCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), false, message);
    }
    for (const [token, message] of nonTargetHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, message);
    }
    for (const token of nonTargetSignatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token), undefined);
    }
    for (const [token, symbolName, message] of closedCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), false, message);
    }
    for (const token of matchedHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, `snapshot 未一致では ${token} hover を出さない`);
    }
    for (const token of matchedSignatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token), undefined, `snapshot 未一致では ${token} signature help を出さない`);
    }
    let tokens = service.getSemanticTokens(uri);
    for (const [lineIndex, identifier, message] of nonTargetSemanticChecks) {
      assertNoSemanticToken(text, tokens, lineIndex, identifier, 0, message);
    }

    service.setActiveWorkbookIdentitySnapshot(createMatchedActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName, blockedSymbolName, message] of matchedCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), true, message);
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, blockedSymbolName), false, message);
    }
    for (const token of matchedHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token)?.contents.includes("CheckBox.Value"), true);
    }
    for (const token of matchedSignatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token)?.label, "Select(Replace) As Object");
    }
    for (const [token, symbolName, message] of nonTargetCompletionChecks.slice(2)) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), false, message);
    }
    for (const [token, message] of nonTargetHoverChecks.slice(2)) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, message);
    }
    for (const token of nonTargetSignatureChecks.slice(2)) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token), undefined);
    }
    tokens = service.getSemanticTokens(uri);
    for (const [lineIndex, identifier, message] of nonTargetSemanticChecks) {
      assertNoSemanticToken(text, tokens, lineIndex, identifier, 0, message);
    }

    service.setActiveWorkbookIdentitySnapshot(createMismatchedActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName] of matchedCompletionChecks) {
      assert.equal(
        hasCompletionSymbolAfterToken(service, uri, text, token, symbolName),
        false,
        `mismatch snapshot では ${token} broad root を開かない`
      );
    }
    for (const token of matchedHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, `mismatch snapshot では ${token} hover を出さない`);
    }
    for (const token of matchedSignatureChecks) {
      assert.equal(
        getSignatureHelpAfterToken(service, uri, text, token),
        undefined,
        `mismatch snapshot では ${token} signature help を出さない`
      );
    }
    tokens = service.getSemanticTokens(uri);
    for (const [lineIndex, identifier, message] of nonTargetSemanticChecks) {
      assertNoSemanticToken(text, tokens, lineIndex, identifier, 0, message);
    }

    service.setActiveWorkbookIdentitySnapshot(createUnavailableActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName] of matchedCompletionChecks) {
      assert.equal(
        hasCompletionSymbolAfterToken(service, uri, text, token, symbolName),
        false,
        `unavailable snapshot では ${token} broad root を開かない`
      );
    }
    tokens = service.getSemanticTokens(uri);
    for (const [lineIndex, identifier, message] of nonTargetSemanticChecks) {
      assertNoSemanticToken(text, tokens, lineIndex, identifier, 0, message);
    }
  } finally {
    cleanup();
  }
});

test("document service resolves workbook-qualified worksheet root item selectors for Shape.OLEFormat.Object", () => {
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Call ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Value
    Call ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Select(
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Value
    Call ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Call ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(
    Call ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Select(
End Sub`;
  const { service, uri, cleanup } = createWorkbookQualifiedWorksheetRootFixture(text);
  const nonTargetCompletionChecks = [
    ['ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.', "Value", 'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'],
    ['ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.', "Value", 'ThisWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'],
    ['ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.', "Value", 'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'],
    ['ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.', "Value", 'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない']
  ];
  const staticCompletionChecks = [
    ['ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.', "Value", "Delete", 'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object は control owner へ解決する'],
    ['ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.', "Value", "Delete", 'ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object は control owner へ解決する']
  ];
  const staticHoverChecks = [
    'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu',
    'ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Valu'
  ];
  const staticSignatureChecks = [
    'ThisWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(',
    'ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select('
  ];
  const nonTargetHoverChecks = [
    ['ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu', 'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'],
    ['ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Valu', 'ThisWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない'],
    ['ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu', 'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので control owner に昇格しない'],
    ['ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Valu', 'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので control owner に昇格しない']
  ];
  const nonTargetSignatureChecks = [
    'ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(',
    'ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Select(',
    'ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(',
    'ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Select('
  ];
  const closedCompletionChecks = [
    ['ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.', "Value", 'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1") は snapshot 未一致の間は broad root を開かない'],
    ['ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.', "Value", 'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1") は snapshot 未一致の間は broad root を開かない']
  ];
  const matchedCompletionChecks = [
    ['ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.', "Value", "Delete", 'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object は control owner へ解決する'],
    ['ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.', "Value", "Delete", 'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object は control owner へ解決する']
  ];
  const matchedHoverChecks = [
    'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu',
    'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Valu'
  ];
  const matchedSignatureChecks = [
    'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(',
    'ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select('
  ];
  const nonTargetSemanticChecks = [
    [10, "Value", 'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので semantic token を出さない'],
    [11, "Value", 'ThisWorkbook.Worksheets.Item(1) は numeric selector なので semantic token を出さない'],
    [12, "Select", 'ThisWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので semantic token を出さない'],
    [13, "Select", 'ThisWorkbook.Worksheets.Item(1) は numeric selector なので semantic token を出さない'],
    [22, "Value", 'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので semantic token を出さない'],
    [23, "Value", 'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので semantic token を出さない'],
    [24, "Select", 'ActiveWorkbook.Worksheets.Item("Sheet1") は codeName 指定なので semantic token を出さない'],
    [25, "Select", 'ActiveWorkbook.Worksheets.Item(1) は numeric selector なので semantic token を出さない']
  ];

  try {
    for (const [token, symbolName, blockedSymbolName, message] of staticCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), true, message);
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, blockedSymbolName), false, message);
    }
    for (const token of staticHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token)?.contents.includes("CheckBox.Value"), true);
    }
    for (const token of staticSignatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token)?.label, "Select(Replace) As Object");
    }
    for (const [token, symbolName, message] of nonTargetCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), false, message);
    }
    for (const [token, message] of nonTargetHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, message);
    }
    for (const token of nonTargetSignatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token), undefined);
    }
    for (const [token, symbolName, message] of closedCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), false, message);
    }
    for (const token of matchedHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, `snapshot 未一致では ${token} hover を出さない`);
    }
    for (const token of matchedSignatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token), undefined, `snapshot 未一致では ${token} signature help を出さない`);
    }
    let tokens = service.getSemanticTokens(uri);
    for (const [lineIndex, identifier, message] of nonTargetSemanticChecks) {
      assertNoSemanticToken(text, tokens, lineIndex, identifier, 0, message);
    }

    service.setActiveWorkbookIdentitySnapshot(createMatchedActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName, blockedSymbolName, message] of matchedCompletionChecks) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), true, message);
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, blockedSymbolName), false, message);
    }
    for (const token of matchedHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token)?.contents.includes("CheckBox.Value"), true);
    }
    for (const token of matchedSignatureChecks) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token)?.label, "Select(Replace) As Object");
    }
    for (const [token, symbolName, message] of nonTargetCompletionChecks.slice(2)) {
      assert.equal(hasCompletionSymbolAfterToken(service, uri, text, token, symbolName), false, message);
    }
    for (const [token, message] of nonTargetHoverChecks.slice(2)) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, message);
    }
    for (const token of nonTargetSignatureChecks.slice(2)) {
      assert.equal(getSignatureHelpAfterToken(service, uri, text, token), undefined);
    }
    tokens = service.getSemanticTokens(uri);
    for (const [lineIndex, identifier, message] of nonTargetSemanticChecks) {
      assertNoSemanticToken(text, tokens, lineIndex, identifier, 0, message);
    }

    service.setActiveWorkbookIdentitySnapshot(createMismatchedActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName] of matchedCompletionChecks) {
      assert.equal(
        hasCompletionSymbolAfterToken(service, uri, text, token, symbolName),
        false,
        `mismatch snapshot では ${token} broad root を開かない`
      );
    }
    for (const token of matchedHoverChecks) {
      assert.equal(getHoverAfterToken(service, uri, text, token), undefined, `mismatch snapshot では ${token} hover を出さない`);
    }
    for (const token of matchedSignatureChecks) {
      assert.equal(
        getSignatureHelpAfterToken(service, uri, text, token),
        undefined,
        `mismatch snapshot では ${token} signature help を出さない`
      );
    }
    tokens = service.getSemanticTokens(uri);
    for (const [lineIndex, identifier, message] of nonTargetSemanticChecks) {
      assertNoSemanticToken(text, tokens, lineIndex, identifier, 0, message);
    }

    service.setActiveWorkbookIdentitySnapshot(createUnavailableActiveWorkbookIdentitySnapshot());

    for (const [token, symbolName] of matchedCompletionChecks) {
      assert.equal(
        hasCompletionSymbolAfterToken(service, uri, text, token, symbolName),
        false,
        `unavailable snapshot では ${token} broad root を開かない`
      );
    }
    tokens = service.getSemanticTokens(uri);
    for (const [lineIndex, identifier, message] of nonTargetSemanticChecks) {
      assertNoSemanticToken(text, tokens, lineIndex, identifier, 0, message);
    }
  } finally {
    cleanup();
  }
});

test("document service keeps unqualified worksheet broad root closed when Worksheets is shadowed", () => {
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Private Function Worksheets(ByVal sheetName As String) As String
    Worksheets = sheetName
End Function

Public Sub Demo()
    Debug.Print Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
End Sub`;
  const { service, uri, cleanup } = createWorksheetBroadRootFixture(text);

  try {
    service.setActiveWorkbookIdentitySnapshot(createMatchedActiveWorkbookIdentitySnapshot());

    assert.equal(
      hasCompletionSymbolAfterToken(service, uri, text, 'Worksheets("Sheet One").OLEObjects("CheckBox1").Object.', "Value"),
      false
    );
    assert.equal(getHoverAfterToken(service, uri, text, 'Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu'), undefined);
    assert.equal(
      getSignatureHelpAfterToken(service, uri, text, 'Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select('),
      undefined
    );
  } finally {
    cleanup();
  }
});

test("document service keeps Application worksheet broad root closed when Application is shadowed", () => {
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Private Type Application
    Name As String
End Type

Public Sub Demo()
    Dim Application As Application
    Debug.Print Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
End Sub`;
  const { service, uri, cleanup } = createWorksheetBroadRootFixture(text);

  try {
    service.setActiveWorkbookIdentitySnapshot(createMatchedActiveWorkbookIdentitySnapshot());

    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(service, uri, text, 'Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu'),
      undefined
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select('
      ),
      undefined
    );
  } finally {
    cleanup();
  }
});

test("document service resolves Application.ThisWorkbook worksheet OLEObject roots and gates Application.ActiveWorkbook", () => {
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Application.ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Value
    Debug.Print Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value
    Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Value
    Call Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Value
    Debug.Print Application.ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value
    Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.
    Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Value
    Call Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function`;
  const { service, uri, cleanup } = createWorkbookQualifiedWorksheetRootFixture(text);

  try {
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        "Value"
      ),
      true
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.',
        "Value"
      ),
      true
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu'
      )?.contents.includes("CheckBox.Value"),
      true
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select('
      )?.label,
      "Select(Replace) As Object"
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(service, uri, text, 'Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu'),
      undefined
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Valu'
      ),
      undefined
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select('
      ),
      undefined
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu'
      ),
      undefined
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select('
      ),
      undefined
    );

    let tokens = service.getSemanticTokens(uri);
    assertSemanticToken(text, tokens, 6, "Value", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 7, "Select", { modifiers: [], type: "function" });
    assertNoSemanticToken(text, tokens, 15, "Value");
    assertNoSemanticToken(text, tokens, 20, "Value");

    service.setActiveWorkbookIdentitySnapshot(createMatchedActiveWorkbookIdentitySnapshot());

    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        "Value"
      ),
      true
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.',
        "Value"
      ),
      true
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu'
      )?.contents.includes("CheckBox.Value"),
      true
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select('
      )?.label,
      "Select(Replace) As Object"
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu'
      ),
      undefined
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Valu'
      ),
      undefined
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select('
      ),
      undefined
    );

    tokens = service.getSemanticTokens(uri);
    assertSemanticToken(text, tokens, 15, "Value", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 16, "Select", { modifiers: [], type: "function" });
    assertNoSemanticToken(text, tokens, 17, "Value");
    assertNoSemanticToken(text, tokens, 20, "Value");
  } finally {
    cleanup();
  }
});

test("document service resolves Application.ThisWorkbook worksheet Shape roots and gates Application.ActiveWorkbook", () => {
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Value
    Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select(
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function`;
  const { service, uri, cleanup } = createWorkbookQualifiedWorksheetRootFixture(text);

  try {
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        "Value"
      ),
      true
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
        "Value"
      ),
      true
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu'
      )?.contents.includes("CheckBox.Value"),
      true
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select('
      )?.label,
      "Select(Replace) As Object"
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Valu'
      ),
      undefined
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Valu'
      ),
      undefined
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select('
      ),
      undefined
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu'
      ),
      undefined
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select('
      ),
      undefined
    );

    let tokens = service.getSemanticTokens(uri);
    assertSemanticToken(text, tokens, 6, "Value", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 7, "Select", { modifiers: [], type: "function" });
    assertNoSemanticToken(text, tokens, 15, "Value");
    assertNoSemanticToken(text, tokens, 20, "Value");

    service.setActiveWorkbookIdentitySnapshot(createMatchedActiveWorkbookIdentitySnapshot());

    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        "Value"
      ),
      true
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
        "Value"
      ),
      true
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu'
      )?.contents.includes("CheckBox.Value"),
      true
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select('
      )?.label,
      "Select(Replace) As Object"
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Valu'
      ),
      undefined
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Valu'
      ),
      undefined
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select('
      ),
      undefined
    );

    tokens = service.getSemanticTokens(uri);
    assertSemanticToken(text, tokens, 15, "Value", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 16, "Select", { modifiers: [], type: "function" });
    assertNoSemanticToken(text, tokens, 17, "Value");
    assertNoSemanticToken(text, tokens, 20, "Value");
  } finally {
    cleanup();
  }
});

test("document service keeps Application workbook roots closed when Application is shadowed", () => {
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Private Type Application
    Name As String
End Type

Public Sub Demo()
    Dim Application As Application
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.
    Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value
    Call Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.
    Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value
    Call Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(
End Sub`;
  const { service, uri, cleanup } = createWorksheetBroadRootFixture(text);

  try {
    service.setActiveWorkbookIdentitySnapshot(createMatchedActiveWorkbookIdentitySnapshot());

    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu'
      ),
      undefined
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select('
      ),
      undefined
    );
    assert.equal(
      hasCompletionSymbolAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        "Value"
      ),
      false
    );
    assert.equal(
      getHoverAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu'
      ),
      undefined
    );
    assert.equal(
      getSignatureHelpAfterToken(
        service,
        uri,
        text,
        'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select('
      ),
      undefined
    );
  } finally {
    cleanup();
  }
});

test("document service reloads the root document module sidecar after sidecar-only regeneration", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "bundle-a");
  const moduleDirectory = path.join(bundleRoot, "modules");
  const sheetAUri = pathToFileURL(path.join(bundleRoot, "SheetA.cls")).href;
  const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print SheetA.OLEObjects("Control1").Object.
End Sub`;

  mkdirSync(moduleDirectory, { recursive: true });
  writeWorksheetControlMetadataSidecar(bundleRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [
      {
        controls: [
          {
            codeName: "chkFinished",
            controlType: "CheckBox",
            progId: "Forms.CheckBox.1",
            shapeId: 3,
            shapeName: "Control1"
          }
        ],
        ownerKind: "worksheet",
        sheetCodeName: "SheetA",
        sheetName: "SheetA",
        status: "supported"
      }
    ],
    version: 1,
    workbook: {
      name: "bundle-a.xlsm",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });

    service.analyzeText(
      sheetAUri,
      "vba",
      1,
      `Attribute VB_Name = "SheetA"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(uri, "vba", 1, text);

    const objectPosition = findPositionAfterTokenInText(text, 'SheetA.OLEObjects("Control1").Object.');
    const initialMembers = service.getCompletionSymbols(uri, objectPosition);
    const initialValueCompletion = initialMembers.find((resolution) => resolution.symbol.name === "Value");

    assert.equal(initialValueCompletion?.moduleName.includes("CheckBox"), true);

    writeWorksheetControlMetadataSidecar(bundleRoot, {
      artifact: "worksheet-control-metadata-sidecar",
      owners: [
        {
          controls: [
            {
              codeName: "optFinished",
              controlType: "OptionButton",
              progId: "Forms.OptionButton.1",
              shapeId: 3,
              shapeName: "Control1"
            }
          ],
          ownerKind: "worksheet",
          sheetCodeName: "SheetA",
          sheetName: "SheetA",
          status: "supported"
        }
      ],
      version: 1,
      workbook: {
        name: "bundle-a.xlsm",
        sourceKind: "openxml-package"
      }
    });

    const refreshedMembers = service.getCompletionSymbols(uri, objectPosition);
    const refreshedValueCompletion = refreshedMembers.find((resolution) => resolution.symbol.name === "Value");

    assert.equal(refreshedValueCompletion?.moduleName.includes("OptionButton"), true);
    assert.equal(refreshedValueCompletion?.moduleName.includes("CheckBox"), false);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service resolves worksheet control code names through the worksheet root sidecar only", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "book1");
  const moduleDirectory = path.join(bundleRoot, "modules");
  const sheet1Uri = pathToFileURL(path.join(bundleRoot, "Sheet1.cls")).href;
  const chart1Uri = pathToFileURL(path.join(bundleRoot, "Chart1.cls")).href;
  const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print Sheet1.chkFinished.
    Debug.Print Sheet1.CheckBox1.
    Debug.Print Chart1.chkFinished.
    Debug.Print ActiveSheet.chkFinished.
    Debug.Print Sheet1.chkFinished.Value
    Call Sheet1.chkFinished.Select(
End Sub`;

  mkdirSync(moduleDirectory, { recursive: true });
  writeWorksheetControlMetadataSidecar(bundleRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [
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
      },
      {
        controls: [
          {
            codeName: "chkFinished",
            controlType: "CheckBox",
            progId: "Forms.CheckBox.1",
            shapeId: 8,
            shapeName: "ChartCheckBox1"
          }
        ],
        ownerKind: "chartsheet",
        sheetCodeName: "Chart1",
        sheetName: "Chart1",
        status: "supported"
      }
    ],
    version: 1,
    workbook: {
      name: "book1.xlsm",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });

    service.analyzeText(
      sheet1Uri,
      "vba",
      1,
      `Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(
      chart1Uri,
      "vba",
      1,
      `Attribute VB_Name = "Chart1"
Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(uri, "vba", 1, text);

    const controlCodeNameMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.chkFinished."));
    const shapeNameMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.CheckBox1."));
    const chartMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Chart1.chkFinished."));
    const activeSheetMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "ActiveSheet.chkFinished."));
    const valueHover = service.getHover(uri, findPositionAfterTokenInText(text, "Sheet1.chkFinished.Valu"));
    const selectSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Sheet1.chkFinished.Select("));
    const tokens = service.getSemanticTokens(uri);

    const valueCompletion = controlCodeNameMembers.find((resolution) => resolution.symbol.name === "Value");
    const selectCompletion = controlCodeNameMembers.find((resolution) => resolution.symbol.name === "Select");

    assert.equal(valueCompletion?.moduleName.includes("CheckBox property"), true);
    assert.equal(selectCompletion?.moduleName.includes("CheckBox method"), true);
    assert.equal(controlCodeNameMembers.some((resolution) => resolution.symbol.name === "Activate"), false);
    assert.equal(shapeNameMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(chartMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(activeSheetMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(valueHover?.contents.includes("CheckBox.Value"), true);
    assert.equal(valueHover?.contents.includes("microsoft.office.interop.excel.checkbox.value"), true);
    assert.equal(selectSignature?.label, "Select(Replace) As Object");
    assertSemanticToken(text, tokens, 8, "Value", { modifiers: [], type: "variable" });
    assertSemanticToken(text, tokens, 9, "Select", { modifiers: [], type: "function" });
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service keeps worksheet control code names unresolved without a sidecar", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "book1");
  const moduleDirectory = path.join(bundleRoot, "modules");
  const sheet1Uri = pathToFileURL(path.join(bundleRoot, "Sheet1.cls")).href;
  const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
  const text = `Attribute VB_Name = "Module1"
Option Explicit

Public Sub Demo()
    Debug.Print Sheet1.chkFinished.
    Debug.Print Sheet1.chkFinished.Value
    Call Sheet1.chkFinished.Select(
End Sub`;

  mkdirSync(moduleDirectory, { recursive: true });

  try {
    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });

    service.analyzeText(
      sheet1Uri,
      "vba",
      1,
      `Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(uri, "vba", 1, text);

    const controlCodeNameMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Sheet1.chkFinished."));
    const valueHover = service.getHover(uri, findPositionAfterTokenInText(text, "Sheet1.chkFinished.Valu"));
    const selectSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Sheet1.chkFinished.Select("));
    const tokens = service.getSemanticTokens(uri);

    assert.equal(controlCodeNameMembers.some((resolution) => resolution.symbol.name === "Value"), false);
    assert.equal(valueHover, undefined);
    assert.equal(selectSignature, undefined);
    assertNoSemanticToken(text, tokens, 5, "Value");
    assertNoSemanticToken(text, tokens, 6, "Select");
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service keeps built-in document roots conservative for unknown predeclared class modules", () => {
  const service = createDocumentService();
  const chart1Uri = "file:///C:/temp/Chart1.cls";
  const chart2Uri = "file:///C:/temp/Chart2.cls";
  const dialogSheet1Uri = "file:///C:/temp/DialogSheet1.cls";
  const uri = "file:///C:/temp/BuiltInChartShadowing.bas";
  const text = `Attribute VB_Name = "BuiltInChartShadowing"
Option Explicit

Public Sub Demo()
    Debug.Print Chart1.
    Debug.Print Chart1.Evaluate
    Debug.Print Chart2.
    Debug.Print Chart2.ChartArea
    Debug.Print DialogSheet1.
    Debug.Print DialogSheet1.ChartArea
    Call Chart2.SetSourceData(Range("A1:B2"))
    Call DialogSheet1.SetSourceData(Range("A1:B2"))
End Sub`;

  service.analyzeText(
    chart1Uri,
    "vba",
    1,
    `Attribute VB_Name = "Chart1"
Attribute VB_Base = "0{11111111-1111-1111-1111-111111111111}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(
    chart2Uri,
    "vba",
    1,
    `Attribute VB_Name = "Chart2"
Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = False
Option Explicit`
  );
  service.analyzeText(
    dialogSheet1Uri,
    "vba",
    1,
    `Attribute VB_Name = "DialogSheet1"
Attribute VB_Base = "0{00020830-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(uri, "vba", 1, text);

  const chart1Completions = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Chart1."));
  const chart1Hover = service.getHover(uri, findPositionAfterTokenInText(text, "Chart1.Evalu"));
  const chart2Completions = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "Chart2."));
  const chart2Hover = service.getHover(uri, findPositionAfterTokenInText(text, "Chart2.ChartA"));
  const dialogSheet1Completions = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheet1."));
  const dialogSheet1Hover = service.getHover(uri, findPositionAfterTokenInText(text, "DialogSheet1.ChartA"));
  const chart2Signature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Chart2.SetSourceData("));
  const dialogSheet1Signature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheet1.SetSourceData(")
  );
  const tokens = service.getSemanticTokens(uri);

  assert.deepEqual(chart1Completions, []);
  assert.equal(chart1Hover, undefined);
  assert.deepEqual(chart2Completions, []);
  assert.equal(chart2Hover, undefined);
  assert.equal(chart2Signature, undefined);
  assert.deepEqual(dialogSheet1Completions, []);
  assert.equal(dialogSheet1Hover, undefined);
  assert.equal(dialogSheet1Signature, undefined);
  assert.equal(
    tokens.some(
      (entry) =>
        entry.range.start.line === 6 &&
        entry.range.start.character === 23 &&
        entry.range.end.character === 31 &&
        entry.type === "function"
    ),
    false
  );
  assert.equal(
    tokens.some(
      (entry) =>
        entry.range.start.line === 10 &&
        entry.range.start.character === 29 &&
        entry.range.end.character === 38 &&
        entry.type === "variable"
    ),
    false
  );
});

test("document service exposes DialogSheet interop signature help conservatively through indexed DialogSheets access", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/DialogSheetBuiltInSignature.bas";
  const text = `Attribute VB_Name = "DialogSheetBuiltInSignature"
Option Explicit

Public Sub Demo()
    Debug.Print DialogSheets(1).Evaluate("A1")
    Call DialogSheets(1).SaveAs("Dialog1.xlsx")
    Call DialogSheets(1).ExportAsFixedFormat(xlTypePDF)
    Call DialogSheets(Array("Dialog1", "Dialog2")).SaveAs("Dialog1.xlsx")
    Call DialogSheets.Item(1).SaveAs("Dialog1.xlsx")
End Sub`;

  service.analyzeText(uri, "vba", 1, text);

  const evaluateSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Evaluate("));
  const saveAsSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "DialogSheets(1).SaveAs("));
  const exportSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).ExportAsFixedFormat(")
  );
  const groupedSaveAsSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(Array("Dialog1", "Dialog2")).SaveAs(')
  );
  const itemSaveAsSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets.Item(1).SaveAs(")
  );

  assert.equal(evaluateSignature?.label, "Evaluate(Name) As Object");
  assert.equal(
    evaluateSignature?.parameters[0]?.documentation?.includes("想定型: Object"),
    true,
  );
  assert.equal(saveAsSignature?.label, "SaveAs(Filename, FileFormat, Password, ..., Local)");
  assert.equal(saveAsSignature?.parameters.length, 10);
  assert.equal(
    saveAsSignature?.parameters[0]?.documentation?.includes("必須引数"),
    true,
  );
  assert.equal(
    saveAsSignature?.parameters[1]?.documentation?.includes("省略可能"),
    true,
  );
  assert.equal(exportSignature?.label, "ExportAsFixedFormat(Type, Filename, Quality, ..., FixedFormatExtClassPtr)");
  assert.equal(exportSignature?.parameters.length, 9);
  assert.equal(itemSaveAsSignature?.label, "SaveAs(Filename, FileFormat, Password, ..., Local)");
  assert.equal(itemSaveAsSignature?.parameters.length, 10);
  assert.equal(groupedSaveAsSignature, undefined);
});

test("document service exposes DialogSheet common callable members through Application and Workbook DialogSheets roots", () => {
  const service = createDocumentService();
  const thisWorkbookUri = "file:///C:/temp/ThisWorkbook.cls";
  const uri = "file:///C:/temp/DialogSheetBuiltInRoots.bas";
  const text = `Attribute VB_Name = "DialogSheetBuiltInRoots"
Option Explicit

Public Sub Demo()
    Debug.Print Application.DialogSheets.
    Debug.Print Application.DialogSheets(1).
    Debug.Print ActiveWorkbook.DialogSheets.
    Debug.Print ActiveWorkbook.DialogSheets(1).
    Debug.Print Application.DialogSheets(Array("Dialog1", "Dialog2")).
    Debug.Print Application.DialogSheets(1).Evaluate("A1")
    Call ActiveWorkbook.DialogSheets(1).SaveAs("Dialog1.xlsx")
    Call Application.DialogSheets(Array("Dialog1", "Dialog2")).SaveAs("Dialog1.xlsx")
    Call ThisWorkbook.DialogSheets(1).SaveAs("Dialog1.xlsx")
End Sub`;

  service.analyzeText(
    thisWorkbookUri,
    "vba",
    1,
    `Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(uri, "vba", 1, text);

  const applicationDialogSheetsMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Application.DialogSheets.")
  );
  const indexedApplicationDialogSheetMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Application.DialogSheets(1).")
  );
  const workbookDialogSheetsMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.DialogSheets.")
  );
  const indexedWorkbookDialogSheetMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.DialogSheets(1).")
  );
  const indexedThisWorkbookDialogSheetMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "ThisWorkbook.DialogSheets(1).")
  );
  const groupedApplicationDialogSheetsMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'Application.DialogSheets(Array("Dialog1", "Dialog2")).')
  );
  const evaluateSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Application.DialogSheets(1).Evaluate(")
  );
  const saveAsSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.DialogSheets(1).SaveAs(")
  );
  const thisWorkbookSaveAsSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "ThisWorkbook.DialogSheets(1).SaveAs(")
  );
  const groupedSaveAsSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, 'Application.DialogSheets(Array("Dialog1", "Dialog2")).SaveAs(')
  );
  const dialogSheetSaveAsHover = service.getHover(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.DialogSheets(1).SaveA")
  );
  const tokens = service.getSemanticTokens(uri);

  const applicationDialogSheetsCount = applicationDialogSheetsMembers.find(
    (resolution) => resolution.symbol.name === "Count"
  );
  const indexedApplicationDialogSheetEvaluate = indexedApplicationDialogSheetMembers.find(
    (resolution) => resolution.symbol.name === "Evaluate"
  );
  const workbookDialogSheetsCount = workbookDialogSheetsMembers.find(
    (resolution) => resolution.symbol.name === "Count"
  );
  const indexedWorkbookDialogSheetSaveAs = indexedWorkbookDialogSheetMembers.find(
    (resolution) => resolution.symbol.name === "SaveAs"
  );
  const indexedThisWorkbookDialogSheetSaveAs = indexedThisWorkbookDialogSheetMembers.find(
    (resolution) => resolution.symbol.name === "SaveAs"
  );
  const groupedApplicationDialogSheetsCount = groupedApplicationDialogSheetsMembers.find(
    (resolution) => resolution.symbol.name === "Count"
  );

  assert.equal(applicationDialogSheetsCount?.moduleName, "Excel DialogSheets property");
  assert.equal(indexedApplicationDialogSheetEvaluate?.documentation?.includes("dialogsheet.evaluate"), true);
  assert.equal(workbookDialogSheetsCount?.moduleName, "Excel DialogSheets property");
  assert.equal(indexedWorkbookDialogSheetSaveAs?.moduleName, "Excel DialogSheet method");
  assert.equal(indexedWorkbookDialogSheetSaveAs?.documentation?.includes("dialogsheet.saveas"), true);
  assert.equal(indexedThisWorkbookDialogSheetSaveAs?.moduleName, "Excel DialogSheet method");
  assert.equal(groupedApplicationDialogSheetsCount?.moduleName, "Excel DialogSheets property");
  assert.equal(groupedApplicationDialogSheetsMembers.some((resolution) => resolution.symbol.name === "SaveAs"), false);
  assert.equal(evaluateSignature?.label, "Evaluate(Name) As Object");
  assert.equal(saveAsSignature?.label, "SaveAs(Filename, FileFormat, Password, ..., Local)");
  assert.equal(thisWorkbookSaveAsSignature?.label, "SaveAs(Filename, FileFormat, Password, ..., Local)");
  assert.equal(groupedSaveAsSignature, undefined);
  assert.equal(dialogSheetSaveAsHover?.contents.includes("microsoft.office.interop.excel.dialogsheet.saveas"), true);
  assertSemanticToken(text, tokens, 4, "DialogSheets", {
    modifiers: [],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 10, "SaveAs", {
    modifiers: [],
    type: "function"
  });
  assertNoSemanticToken(text, tokens, 11, "SaveAs");
});

test("document service exposes DialogFrame supplemental members through DialogSheet roots", () => {
  const service = createDocumentService();
  const thisWorkbookUri = "file:///C:/temp/ThisWorkbook.cls";
  const uri = "file:///C:/temp/DialogFrameBuiltIn.bas";
  const text = `Attribute VB_Name = "DialogFrameBuiltIn"
Option Explicit

Public Sub Demo()
    Debug.Print DialogSheets(1).
    Debug.Print DialogSheets(1).DialogFrame.
    Debug.Print DialogSheets(1).DialogFrame.Caption
    Call DialogSheets(1).DialogFrame.Select("DialogFrame1")
    Debug.Print Application.DialogSheets(1).DialogFrame.
    Debug.Print Application.DialogSheets(1).DialogFrame.Text
    Call ActiveWorkbook.DialogSheets(1).DialogFrame.Select("DialogFrame1")
    Call ThisWorkbook.DialogSheets(1).DialogFrame.Select("DialogFrame1")
    Call DialogSheets(Array("Dialog1", "Dialog2")).DialogFrame.Select("DialogFrame1")
    Debug.Print DialogSheets("Dialog1").DialogFrame.
    Call DialogSheets.Item(1).DialogFrame.Select("DialogFrame1")
    Debug.Print DialogSheets(1).DialogFrame.Caption(
    Debug.Print Application.DialogSheets(1).DialogFrame.Text(
End Sub`;

  service.analyzeText(
    thisWorkbookUri,
    "vba",
    1,
    `Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(uri, "vba", 1, text);

  const indexedDialogSheetMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1)."));
  const indexedDialogFrameMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).DialogFrame.")
  );
  const applicationDialogFrameMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "Application.DialogSheets(1).DialogFrame.")
  );
  const namedDialogFrameMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets("Dialog1").DialogFrame.')
  );
  const groupedDialogSheetsMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(Array("Dialog1", "Dialog2")).')
  );
  const captionHover = service.getHover(uri, findPositionAfterTokenInText(text, "DialogSheets(1).DialogFrame.Capti"));
  const applicationTextHover = service.getHover(
    uri,
    findPositionAfterTokenInText(text, "Application.DialogSheets(1).DialogFrame.Tex")
  );
  const selectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).DialogFrame.Select(")
  );
  const activeWorkbookSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.DialogSheets(1).DialogFrame.Select(")
  );
  const itemSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets.Item(1).DialogFrame.Select(")
  );
  const thisWorkbookSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "ThisWorkbook.DialogSheets(1).DialogFrame.Select(")
  );
  const groupedSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(Array("Dialog1", "Dialog2")).DialogFrame.Select(')
  );
  const captionPropertySignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).DialogFrame.Caption(")
  );
  const applicationTextPropertySignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Application.DialogSheets(1).DialogFrame.Text(")
  );
  const tokens = service.getSemanticTokens(uri);

  const dialogFrameProperty = indexedDialogSheetMembers.find((resolution) => resolution.symbol.name === "DialogFrame");
  const dialogFrameCaption = indexedDialogFrameMembers.find((resolution) => resolution.symbol.name === "Caption");
  const dialogFrameSelect = indexedDialogFrameMembers.find((resolution) => resolution.symbol.name === "Select");
  const applicationDialogFrameText = applicationDialogFrameMembers.find((resolution) => resolution.symbol.name === "Text");
  const namedDialogFrameCaption = namedDialogFrameMembers.find((resolution) => resolution.symbol.name === "Caption");

  assert.equal(dialogFrameProperty?.moduleName, "Excel DialogSheet property");
  assert.equal(dialogFrameProperty?.typeName, "DialogFrame");
  assert.equal(dialogFrameProperty?.documentation?.includes("dialogsheet.dialogframe"), true);
  assert.equal(dialogFrameCaption?.moduleName, "Excel DialogFrame property");
  assert.equal(dialogFrameCaption?.typeName, "String");
  assert.equal(dialogFrameCaption?.documentation?.includes("dialogframe.caption"), true);
  assert.equal(dialogFrameSelect?.moduleName, "Excel DialogFrame method");
  assert.equal(dialogFrameSelect?.documentation?.includes("dialogframe.select"), true);
  assert.equal(applicationDialogFrameText?.moduleName, "Excel DialogFrame property");
  assert.equal(applicationDialogFrameText?.documentation?.includes("dialogframe.text"), true);
  assert.equal(namedDialogFrameCaption?.moduleName, "Excel DialogFrame property");
  assert.equal(namedDialogFrameCaption?.documentation?.includes("dialogframe.caption"), true);
  assert.equal(groupedDialogSheetsMembers.some((resolution) => resolution.symbol.name === "DialogFrame"), false);
  assert.equal(captionHover?.contents.includes("DialogFrame.Caption"), true);
  assert.equal(captionHover?.contents.includes("microsoft.office.interop.excel.dialogframe.caption"), true);
  assert.equal(applicationTextHover?.contents.includes("DialogFrame.Text"), true);
  assert.equal(applicationTextHover?.contents.includes("microsoft.office.interop.excel.dialogframe.text"), true);
  assert.equal(selectSignature?.label, "Select(Replace) As Object");
  assert.equal(selectSignature?.parameters.length, 1);
  assert.equal(selectSignature?.parameters[0]?.documentation?.includes("想定型: Object"), true);
  assert.equal(selectSignature?.parameters[0]?.documentation?.includes("省略可能"), true);
  assert.equal(activeWorkbookSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(itemSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(thisWorkbookSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(groupedSelectSignature, undefined);
  assert.equal(captionPropertySignature, undefined);
  assert.equal(applicationTextPropertySignature, undefined);
  assertSemanticToken(text, tokens, 5, "DialogFrame", {
    modifiers: [],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 6, "DialogFrame", {
    modifiers: [],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 6, "Caption", {
    modifiers: [],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 7, "Select", {
    modifiers: [],
    type: "function"
  });
  assertSemanticToken(text, tokens, 9, "Text", {
    modifiers: [],
    type: "variable"
  });
  assertNoSemanticToken(text, tokens, 12, "Select");
});

test("document service normalizes DialogSheet control collection selectors conservatively", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/DialogSheetControlCollection.bas";
  const text = `Attribute VB_Name = "DialogSheetControlCollection"
Option Explicit

Public Sub Demo()
    Dim index As Long
    Debug.Print DialogSheets(1).Buttons.
    Debug.Print DialogSheets(1).Buttons(1).
    Debug.Print DialogSheets(1).Buttons(&H1).
    Debug.Print DialogSheets(1).Buttons(&O7).
    Debug.Print DialogSheets(1).Buttons(1#).
    Debug.Print DialogSheets(1).Buttons(1E+2).
    Debug.Print DialogSheets(1).Buttons("Button 1").
    Debug.Print DialogSheets(1).Buttons(index).
    Debug.Print DialogSheets(1).Buttons(Array(1, 2)).
    Debug.Print DialogSheets(1).Buttons.Item(1).
    Debug.Print DialogSheets(1).Buttons.Item(index).
    Debug.Print DialogSheets(1).CheckBoxes(1).
    Debug.Print DialogSheets(1).OptionButtons("Option 1").
    Call DialogSheets(1).Buttons(1).Select("Button 1")
    Call DialogSheets(1).Buttons.Item(1).Select("Button 1")
    Call DialogSheets(1).CheckBoxes(1).Select("Check 1")
    Call DialogSheets(1).CheckBoxes.Item(1).Select("Check 1")
    Call DialogSheets(1).OptionButtons("Option 1").Select("Option 1")
    Call DialogSheets(1).OptionButtons.Item(1).Select("Option 1")
    Call Application.DialogSheets(1).Buttons(1).Select("Button 1")
    Debug.Print DialogSheets(1).Buttons(1).Caption
    Debug.Print DialogSheets(1).CheckBoxes(1).Value
    Debug.Print DialogSheets(1).OptionButtons("Option 1").Value
    Debug.Print DialogSheets(1).CheckBoxes(1).Value(
    Debug.Print DialogSheets(1).OptionButtons("Option 1").Value(
    Call DialogSheets(1).Buttons(index).Select("Button 1")
    Call DialogSheets(1).Buttons.Item(index).Select("Button 1")
    Call DialogSheets(1).Buttons(Array(1, 2)).Select("Button 1")
    Debug.Print DialogSheets(1).Buttons.Item("Button 1").
    Debug.Print DialogSheets(1).CheckBoxes.Item("Check 1").
    Debug.Print DialogSheets(1).OptionButtons.Item("Option 1").
    Call DialogSheets(1).Buttons(&H1).Select("Button 1")
    Call DialogSheets(1).Buttons(&O7).Select("Button 1")
    Call DialogSheets(1).Buttons(1#).Select("Button 1")
    Call DialogSheets(1).Buttons(1E+2).Select("Button 1")
    Call DialogSheets(1).Buttons.Item("Button 1").Select("Button 1")
    Call DialogSheets(1).CheckBoxes.Item("Check 1").Select("Check 1")
    Call DialogSheets(1).OptionButtons.Item("Option 1").Select("Option 1")
End Sub`;

  service.analyzeText(uri, "vba", 1, text);

  const buttonsCollectionMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Buttons."));
  const indexedButtonMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(1)."));
  const hexButtonMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(&H1)."));
  const octalButtonMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(&O7)."));
  const suffixButtonMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(1#)."));
  const exponentButtonMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(1E+2)."));
  const namedButtonMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).Buttons("Button 1").')
  );
  const dynamicButtonMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(index)."));
  const groupedButtonMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(Array(1, 2)).")
  );
  const itemButtonMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Buttons.Item(1)."));
  const namedItemButtonMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).Buttons.Item("Button 1").')
  );
  const dynamicItemButtonMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).Buttons.Item(index).")
  );
  const indexedCheckBoxMembers = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "DialogSheets(1).CheckBoxes(1)."));
  const namedItemCheckBoxMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).CheckBoxes.Item("Check 1").')
  );
  const namedOptionButtonMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).OptionButtons("Option 1").')
  );
  const namedItemOptionButtonMembers = service.getCompletionSymbols(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).OptionButtons.Item("Option 1").')
  );
  const buttonCaptionHover = service.getHover(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(1).Capti"));
  const checkBoxValueHover = service.getHover(uri, findPositionAfterTokenInText(text, "DialogSheets(1).CheckBoxes(1).Valu"));
  const optionButtonValueHover = service.getHover(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).OptionButtons("Option 1").Valu')
  );
  const buttonSelectSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(1).Select("));
  const itemButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).Buttons.Item(1).Select(")
  );
  const hexButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(&H1).Select(")
  );
  const octalButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(&O7).Select(")
  );
  const suffixButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(1#).Select(")
  );
  const exponentButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(1E+2).Select(")
  );
  const namedItemButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).Buttons.Item("Button 1").Select(')
  );
  const checkBoxSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).CheckBoxes(1).Select(")
  );
  const itemCheckBoxSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).CheckBoxes.Item(1).Select(")
  );
  const namedItemCheckBoxSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).CheckBoxes.Item("Check 1").Select(')
  );
  const optionButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).OptionButtons("Option 1").Select(')
  );
  const itemOptionButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).OptionButtons.Item(1).Select(")
  );
  const namedItemOptionButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).OptionButtons.Item("Option 1").Select(')
  );
  const applicationButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Application.DialogSheets(1).Buttons(1).Select(")
  );
  const dynamicButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(index).Select(")
  );
  const dynamicItemButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).Buttons.Item(index).Select(")
  );
  const groupedButtonSelectSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).Buttons(Array(1, 2)).Select(")
  );
  const checkBoxValueSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "DialogSheets(1).CheckBoxes(1).Value(")
  );
  const optionButtonValueSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, 'DialogSheets(1).OptionButtons("Option 1").Value(')
  );
  const tokens = service.getSemanticTokens(uri);

  const buttonsCount = buttonsCollectionMembers.find((resolution) => resolution.symbol.name === "Count");
  const buttonsItem = buttonsCollectionMembers.find((resolution) => resolution.symbol.name === "Item");
  const indexedButtonCaption = indexedButtonMembers.find((resolution) => resolution.symbol.name === "Caption");
  const hexButtonCaption = hexButtonMembers.find((resolution) => resolution.symbol.name === "Caption");
  const octalButtonCaption = octalButtonMembers.find((resolution) => resolution.symbol.name === "Caption");
  const suffixButtonCaption = suffixButtonMembers.find((resolution) => resolution.symbol.name === "Caption");
  const exponentButtonCaption = exponentButtonMembers.find((resolution) => resolution.symbol.name === "Caption");
  const indexedButtonSelect = indexedButtonMembers.find((resolution) => resolution.symbol.name === "Select");
  const namedButtonCaption = namedButtonMembers.find((resolution) => resolution.symbol.name === "Caption");
  const dynamicButtonsCount = dynamicButtonMembers.find((resolution) => resolution.symbol.name === "Count");
  const groupedButtonsCount = groupedButtonMembers.find((resolution) => resolution.symbol.name === "Count");
  const itemButtonCaption = itemButtonMembers.find((resolution) => resolution.symbol.name === "Caption");
  const namedItemButtonCaption = namedItemButtonMembers.find((resolution) => resolution.symbol.name === "Caption");
  const dynamicItemButtonsCount = dynamicItemButtonMembers.find((resolution) => resolution.symbol.name === "Count");
  const checkBoxValue = indexedCheckBoxMembers.find((resolution) => resolution.symbol.name === "Value");
  const namedItemCheckBoxValue = namedItemCheckBoxMembers.find((resolution) => resolution.symbol.name === "Value");
  const optionButtonValue = namedOptionButtonMembers.find((resolution) => resolution.symbol.name === "Value");
  const namedItemOptionButtonValue = namedItemOptionButtonMembers.find((resolution) => resolution.symbol.name === "Value");

  assert.equal(buttonsCount?.moduleName, "Excel Buttons property");
  assert.equal(buttonsItem?.moduleName, "Excel Buttons method");
  assert.equal(indexedButtonCaption?.moduleName, "Excel Button property");
  assert.equal(hexButtonCaption?.moduleName, "Excel Button property");
  assert.equal(octalButtonCaption?.moduleName, "Excel Button property");
  assert.equal(suffixButtonCaption?.moduleName, "Excel Button property");
  assert.equal(exponentButtonCaption?.moduleName, "Excel Button property");
  assert.equal(indexedButtonCaption?.documentation?.includes("excel.button.caption"), true);
  assert.equal(indexedButtonSelect?.moduleName, "Excel Button method");
  assert.equal(namedButtonCaption?.moduleName, "Excel Button property");
  assert.equal(dynamicButtonsCount?.moduleName, "Excel Buttons property");
  assert.equal(groupedButtonsCount?.moduleName, "Excel Buttons property");
  assert.equal(itemButtonCaption?.moduleName, "Excel Button property");
  assert.equal(namedItemButtonCaption?.moduleName, "Excel Button property");
  assert.equal(dynamicItemButtonsCount?.moduleName, "Excel Buttons property");
  assert.equal(checkBoxValue?.moduleName, "Excel CheckBox property");
  assert.equal(namedItemCheckBoxValue?.moduleName, "Excel CheckBox property");
  assert.equal(checkBoxValue?.documentation?.includes("excel.checkbox.value"), true);
  assert.equal(optionButtonValue?.moduleName, "Excel OptionButton property");
  assert.equal(namedItemOptionButtonValue?.moduleName, "Excel OptionButton property");
  assert.equal(optionButtonValue?.documentation?.includes("excel.optionbutton.value"), true);
  assert.equal(dynamicButtonMembers.some((resolution) => resolution.symbol.name === "Caption"), false);
  assert.equal(groupedButtonMembers.some((resolution) => resolution.symbol.name === "Caption"), false);
  assert.equal(dynamicItemButtonMembers.some((resolution) => resolution.symbol.name === "Caption"), false);
  assert.equal(buttonCaptionHover?.contents.includes("Button.Caption"), true);
  assert.equal(buttonCaptionHover?.contents.includes("microsoft.office.interop.excel.button.caption"), true);
  assert.equal(checkBoxValueHover?.contents.includes("CheckBox.Value"), true);
  assert.equal(checkBoxValueHover?.contents.includes("microsoft.office.interop.excel.checkbox.value"), true);
  assert.equal(optionButtonValueHover?.contents.includes("OptionButton.Value"), true);
  assert.equal(optionButtonValueHover?.contents.includes("microsoft.office.interop.excel.optionbutton.value"), true);
  assert.equal(buttonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(itemButtonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(hexButtonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(octalButtonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(suffixButtonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(exponentButtonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(namedItemButtonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(checkBoxSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(itemCheckBoxSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(namedItemCheckBoxSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(optionButtonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(itemOptionButtonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(namedItemOptionButtonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(applicationButtonSelectSignature?.label, "Select(Replace) As Object");
  assert.equal(dynamicButtonSelectSignature, undefined);
  assert.equal(dynamicItemButtonSelectSignature, undefined);
  assert.equal(groupedButtonSelectSignature, undefined);
  assert.equal(checkBoxValueSignature, undefined);
  assert.equal(optionButtonValueSignature, undefined);
  assertSemanticToken(text, tokens, 5, "Buttons", {
    modifiers: [],
    type: "function"
  });
  assertSemanticToken(text, tokens, 18, "Select", {
    modifiers: [],
    type: "function"
  });
  assertSemanticToken(text, tokens, 25, "Caption", {
    modifiers: [],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 26, "Value", {
    modifiers: [],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 27, "Value", {
    modifiers: [],
    type: "variable"
  });
  assertNoSemanticToken(text, tokens, 30, "Select");
  assertNoSemanticToken(text, tokens, 31, "Select");
  assertNoSemanticToken(text, tokens, 32, "Select");
});

test("DialogSheet control collection owner mapping stays aligned with supplemental config", async () => {
  const supplementalConfigModule = await import(
    pathToFileURL(path.resolve(__dirname, "..", "..", "..", "scripts", "lib", "supplementalReferenceConfig.mjs")).href
  );

  for (const config of supplementalConfigModule.dialogSheetControlCollectionOwnerConfigs) {
    assert.equal(
      resolveBuiltinMemberOwnerFromRootType("DialogSheet", [config.collectionName]),
      config.collectionName,
      `DialogSheet.${config.collectionName} は collection owner を返す必要があります`
    );
    assert.equal(
      resolveBuiltinMemberOwnerFromRootType("DialogSheet", [markIndexedAccessPathSegment(config.collectionName, "single")]),
      config.collectionName,
      `DialogSheet.${config.collectionName}(<expr>) は collection owner を維持する必要があります`
    );
    assert.equal(
      resolveBuiltinMemberOwnerFromRootType("DialogSheet", [markIndexedAccessPathSegment(config.collectionName, "literal")]),
      config.itemName,
      `DialogSheet.${config.collectionName}(<literal>) は ${config.itemName} owner を返す必要があります`
    );
  }
});

test("Worksheet and Chart OLEObjects owner mapping stays aligned with indexed access rules", () => {
  for (const ownerName of ["Worksheet", "Chart"]) {
    assert.equal(
      resolveBuiltinMemberOwnerFromRootType(ownerName, ["OLEObjects"]),
      "OLEObjects",
      `${ownerName}.OLEObjects は collection owner を返す必要があります`
    );
    assert.equal(
      resolveBuiltinMemberOwnerFromRootType(ownerName, [markIndexedAccessPathSegment("OLEObjects", "single")]),
      "OLEObject",
      `${ownerName}.OLEObjects(<expr>) は OLEObject owner を返す必要があります`
    );
    assert.equal(
      resolveBuiltinMemberOwnerFromRootType(ownerName, [markIndexedAccessPathSegment("OLEObjects", "literal")]),
      "OLEObject",
      `${ownerName}.OLEObjects(<literal>) は OLEObject owner を返す必要があります`
    );
    assert.equal(
      resolveBuiltinMemberOwnerFromRootType(ownerName, ["OLEObjects", "Item"]),
      "OLEObjects",
      `${ownerName}.OLEObjects.Item は collection owner を返す必要があります`
    );
    assert.equal(
      resolveBuiltinMemberOwnerFromRootType(ownerName, ["OLEObjects", markIndexedAccessPathSegment("Item", "single")]),
      "OLEObject",
      `${ownerName}.OLEObjects.Item(<expr>) は OLEObject owner を返す必要があります`
    );
    assert.equal(
      resolveBuiltinMemberOwnerFromRootType(ownerName, ["OLEObjects", markIndexedAccessPathSegment("Item", "literal")]),
      "OLEObject",
      `${ownerName}.OLEObjects.Item(<literal>) は OLEObject owner を返す必要があります`
    );
  }
});

test("document service keeps built-in member completion and semantic tokens conservative", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/BuiltInMemberShadowing.bas";
  const text = `Attribute VB_Name = "BuiltInMemberShadowing"
Option Explicit

Private Type Application
    Name As String
End Type

Public Sub Demo()
    Dim Application As Application
    Dim foo As Collection
    Debug.Print Application.Name
    Debug.Print foo.Count
End Sub`;

  service.analyzeText(uri, "vba", 1, text);

  const shadowedApplicationMembers = service.getCompletionSymbols(uri, { character: 28, line: 10 });
  const unknownOwnerMembers = service.getCompletionSymbols(uri, { character: 20, line: 11 });
  const tokens = service.getSemanticTokens(uri);

  assert.deepEqual(shadowedApplicationMembers, []);
  assert.deepEqual(unknownOwnerMembers, []);
  assert.equal(
    tokens.some(
      (entry) =>
        entry.range.start.line === 10 &&
        entry.range.start.character === 28 &&
        entry.range.end.character === 32 &&
        entry.type === "variable"
    ),
    false
  );
  assert.equal(
    tokens.some(
      (entry) =>
        entry.range.start.line === 11 &&
        entry.range.start.character === 20 &&
        entry.range.end.character === 25 &&
        entry.type === "function"
    ),
    false
  );
});

test("document service keeps ThisWorkbook built-in alias conservative for non-document class modules", () => {
  const service = createDocumentService();
  const thisWorkbookUri = "file:///C:/temp/ThisWorkbook.cls";
  const uri = "file:///C:/temp/BuiltInThisWorkbookShadowing.bas";
  const text = `Attribute VB_Name = "BuiltInThisWorkbookShadowing"
Option Explicit

Public Sub Demo()
    Debug.Print ThisWorkbook.
    Debug.Print ThisWorkbook.SaveAs
End Sub`;

  service.analyzeText(
    thisWorkbookUri,
    "vba",
    1,
    `Attribute VB_Name = "ThisWorkbook"
Option Explicit

Public Sub SaveAs()
End Sub`
  );
  service.analyzeText(uri, "vba", 1, text);

  const completions = service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, "ThisWorkbook."));
  const hover = service.getHover(uri, findPositionAfterTokenInText(text, "ThisWorkbook.Save"));
  const tokens = service.getSemanticTokens(uri);

  assert.deepEqual(completions, []);
  assert.equal(hover, undefined);
  assert.equal(
    tokens.some(
      (entry) =>
        entry.range.start.line === 5 &&
        entry.range.start.character === 28 &&
        entry.range.end.character === 34 &&
        entry.type === "function"
    ),
    false
  );
});

test("document service offers an Option Explicit code action after existing option lines", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/MissingOptionExplicit.bas";
  const text = `Attribute VB_Name = "MissingOptionExplicit"
Option Compare Text

Public Sub Demo()
    Debug.Print "ready"
End Sub`;

  service.analyzeText(uri, "vba", 1, text);

  const actions = service.getCodeActions(uri);

  assert.equal(actions.length, 1);
  assert.equal(actions[0]?.title, "Option Explicit を追加");
  assert.equal(
    applyTextEdit(text, actions[0].edit),
    `Attribute VB_Name = "MissingOptionExplicit"
Option Compare Text
Option Explicit

Public Sub Demo()
    Debug.Print "ready"
End Sub`
  );
});

test("document service offers an Option Explicit code action for form modules without touching the designer area", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/MissingOptionExplicit.frm";
  const text = `VERSION 5.00
Begin VB.Form MissingOptionExplicit
   Caption = "MissingOptionExplicit"
End
Attribute VB_Name = "MissingOptionExplicit"
Public Sub Demo()
    Debug.Print "ready"
End Sub`;

  service.analyzeText(uri, "vba", 1, text);

  const actions = service.getCodeActions(uri);

  assert.equal(actions.length, 1);
  assert.equal(
    applyTextEdit(text, actions[0].edit),
    `VERSION 5.00
Begin VB.Form MissingOptionExplicit
   Caption = "MissingOptionExplicit"
End
Attribute VB_Name = "MissingOptionExplicit"
Option Explicit

Public Sub Demo()
    Debug.Print "ready"
End Sub`
  );
});

test("document service omits the Option Explicit code action when the module already declares it", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/AlreadyExplicit.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "AlreadyExplicit"
Option Explicit

Public Sub Demo()
End Sub`
  );

  assert.deepEqual(service.getCodeActions(uri), []);
});

test("document service resolves exported symbols across VBA modules", () => {
  const service = createDocumentService();
  const libraryUri = "file:///C:/temp/PublicApi.bas";
  const consumerUri = "file:///C:/temp/Consumer.bas";

  service.analyzeText(
    libraryUri,
    "vba",
    1,
    `Attribute VB_Name = "PublicApi"
Option Explicit

Public Function PublicMessage() As String
    PublicMessage = "Hello"
End Function`
  );
  service.analyzeText(
    consumerUri,
    "vba",
    1,
    `Attribute VB_Name = "Consumer"
Option Explicit

Public Sub UseLibrary()
    Dim message As String
    message = PublicMessage()
End Sub`
  );

  const completions = service.getCompletionSymbols(consumerUri, { character: 4, line: 5 });
  const definition = service.getDefinition(consumerUri, { character: 18, line: 5 });
  const diagnostics = service.getDiagnostics(consumerUri);
  const references = service.getReferences(consumerUri, { character: 18, line: 5 }, true);

  const publicMessageCompletion = completions.find((resolution) => resolution.uri === libraryUri && resolution.symbol.name === "PublicMessage");

  assert.equal(publicMessageCompletion?.typeName, "String");
  assert.equal(definition?.uri, libraryUri);
  assert.equal(definition?.moduleName, "PublicApi");
  assert.equal(definition?.symbol.name, "PublicMessage");
  assert.equal(diagnostics.some((diagnostic) => diagnostic.code === "undeclared-variable"), false);
  assert.deepEqual(
    references.map((reference) => `${reference.uri}:${reference.range.start.line}:${reference.range.start.character}`),
    [`${libraryUri}:3:16`, `${consumerUri}:5:14`]
  );
});

test("document service narrows completion candidates by inferred assignment type", () => {
  const service = createDocumentService();
  const consumerUri = "file:///C:/temp/ConsumerCompletion.bas";

  service.analyzeText(
    "file:///C:/temp/PublicApi.bas",
    "vba",
    1,
    `Attribute VB_Name = "PublicApi"
Option Explicit

Public Function PublicMessage() As String
    PublicMessage = "Hello"
End Function`
  );
  service.analyzeText(
    "file:///C:/temp/NumberApi.bas",
    "vba",
    1,
    `Attribute VB_Name = "NumberApi"
Option Explicit

Public Function PublicNumber() As Long
    PublicNumber = 42
End Function`
  );
  service.analyzeText(
    consumerUri,
    "vba",
    1,
    `Attribute VB_Name = "ConsumerCompletion"
Option Explicit

Public Sub UseLibraryCompletion()
    Dim message As String
    message = Pub
End Sub`
  );

  const completions = service.getCompletionSymbols(consumerUri, { character: 17, line: 5 });

  assert.ok(completions.some((resolution) => resolution.symbol.name === "PublicMessage"));
  assert.equal(completions.some((resolution) => resolution.symbol.name === "PublicNumber"), false);
});

test("document service exposes signature help with inferred argument types", () => {
  const service = createDocumentService();
  const consumerUri = "file:///C:/temp/ConsumerSignature.bas";
  const formatterUri = "file:///C:/temp/FormatterApi.bas";

  service.analyzeText(
    formatterUri,
    "vba",
    1,
    `Attribute VB_Name = "FormatterApi"
Option Explicit

Public Function FormatMessage(ByVal value As String, ByVal count As Long) As String
    FormatMessage = value & CStr(count)
End Function`
  );
  service.analyzeText(
    consumerUri,
    "vba",
    1,
    `Attribute VB_Name = "ConsumerSignature"
Option Explicit

Public Sub UseSignature()
    Dim message As String
    message = FormatMessage(message, 1)
End Sub`
  );

  const signature = service.getSignatureHelp(consumerUri, { character: 38, line: 5 });

  assert.equal(signature?.activeParameter, 1);
  assert.equal(signature?.label, "FormatMessage(ByVal value As String, ByVal count As Long) As String");
  assert.equal(signature?.documentation, "FormatterApi モジュール");
  assert.equal(signature?.parameters[1]?.documentation?.includes("現在の引数型: Long"), true);
});

test("document service exposes built-in member signature help and hover", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/BuiltInSignature.bas";
  const thisWorkbookUri = "file:///C:/temp/ThisWorkbook.cls";
  const sheet1Uri = "file:///C:/temp/Sheet1.cls";
  const chart1Uri = "file:///C:/temp/Chart1.cls";
  const text = `Attribute VB_Name = "BuiltInSignature"
Option Explicit

Public Sub Demo()
    Dim i As Long
    Dim transposedResult As Variant
    Debug.Print WorksheetFunction.Sum(1, 2)
    Debug.Print Application.WorksheetFunction.Sum(1, 2)
    Debug.Print Application.WorksheetFunction.Power(2, 3)
    Debug.Print WorksheetFunction.Average(1, 2, 3)
    Debug.Print WorksheetFunction.Max(1, 2, 3)
    Debug.Print WorksheetFunction.Min(1, 2, 3)
    Debug.Print WorksheetFunction.EDate(Date, 1)
    Debug.Print WorksheetFunction.EoMonth(Date, 1)
    Debug.Print WorksheetFunction.Find("A", "ABC")
    Debug.Print WorksheetFunction.Search("A", "ABC")
    Debug.Print WorksheetFunction.And(True, False, True)
    Debug.Print WorksheetFunction.Or(True, False, True)
    Debug.Print WorksheetFunction.Xor(True, False, True)
    Debug.Print WorksheetFunction.CountA("A", "")
    Debug.Print WorksheetFunction.CountBlank(Range("A1:A2"))
    Debug.Print WorksheetFunction.Text(Now, "yyyy-mm-dd")
    Debug.Print WorksheetFunction.VLookup("A", Range("A1:B2"), 2, False)
    Debug.Print WorksheetFunction.Match("A", Range("A1:A2"), 0)
    Debug.Print WorksheetFunction.Index(Range("A1:B2"), 1, 2)
    Debug.Print WorksheetFunction.Lookup("A", Range("A1:A2"), Range("B1:B2"))
    Debug.Print WorksheetFunction.HLookup("A", Range("A1:B2"), 2, False)
    Debug.Print WorksheetFunction.Choose(1, "A", "B")
    transposedResult = WorksheetFunction.Transpose(Range("A1:B2"))
    Debug.Print UBound(transposedResult, 1), UBound(transposedResult, 2)
    Debug.Print ActiveCell.Address(False, False, xlA1, False)
    Debug.Print Application.ActiveCell.Address(False, False, xlA1, False)
    Debug.Print Cells.AddressLocal(False, False)
    Debug.Print ActiveWorkbook.Worksheets.Count
    Debug.Print Worksheets(1).Evaluate("A1")
    Debug.Print Worksheets("A(1)").Evaluate("A1")
    Debug.Print Worksheets(Array("Sheet1", "Sheet2")).Evaluate("A1")
    Call Worksheets(1).SaveAs("Sheet1.csv")
    Call Worksheets(i + 1).SaveAs("Sheet1.csv")
    Call ActiveWorkbook.Worksheets(1).ExportAsFixedFormat(xlTypePDF)
    Call ActiveWorkbook.Worksheets(GetIndex()).ExportAsFixedFormat(xlTypePDF)
    Debug.Print ThisWorkbook.SaveAs
    Call ThisWorkbook.SaveAs("Book1.xlsx")
    Debug.Print Sheet1.Evaluate("A1")
    Call Sheet1.SaveAs("Sheet1.csv")
    Call Chart1.SetSourceData(Range("A1:B2"))
    Debug.Print Chart1.ChartArea
    Call ActiveWorkbook.Close(False)
    Call ActiveWorkbook.ExportAsFixedFormat(xlTypePDF)
    Call Application.CalculateFull()
    Application.OnTime(Now, "BuiltInSignature.Demo")
    Call Application.WorksheetFunction()
    Call Application.AfterCalculate()
    Call Application.ActiveCell()
    Call Application.NewWorkbook()
    Debug.Print Application.Calculate
End Sub

Private Function GetIndex() As Long
    GetIndex = 1
End Function`;

  service.analyzeText(
    thisWorkbookUri,
    "vba",
    1,
    `Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(
    sheet1Uri,
    "vba",
    1,
    `Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(
    chart1Uri,
    "vba",
    1,
    `Attribute VB_Name = "Chart1"
Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(uri, "vba", 1, text);

  const worksheetSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Sum(1, 2"));
  const chainedSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Application.WorksheetFunction.Sum(")
  );
  const powerSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Application.WorksheetFunction.Power(")
  );
  const averageSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Average("));
  const maxSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Max("));
  const minSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Min("));
  const edateSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.EDate("));
  const eomonthSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.EoMonth("));
  const findSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Find("));
  const searchSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Search("));
  const andSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.And("));
  const orSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Or("));
  const xorSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Xor("));
  const countASignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.CountA("));
  const countBlankSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "WorksheetFunction.CountBlank(")
  );
  const textSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Text("));
  const vlookupSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.VLookup("));
  const matchSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Match("));
  const indexSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Index("));
  const lookupSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Lookup("));
  const hlookupSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.HLookup("));
  const chooseSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Choose("));
  const transposeSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "WorksheetFunction.Transpose(")
  );
  const addressSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "ActiveCell.Address("));
  const chainedAddressSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Application.ActiveCell.Address(")
  );
  const addressLocalSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Cells.AddressLocal("));
  const worksheetEvaluateSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Worksheets(1).Evaluate("));
  const worksheetStringEvaluateSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Worksheets(\"A(1)\").Evaluate(")
  );
  const groupedWorksheetEvaluateSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, 'Worksheets(Array("Sheet1", "Sheet2")).Evaluate(')
  );
  const worksheetSaveAsSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Worksheets(1).SaveAs("));
  const worksheetExpressionSaveAsSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Worksheets(i + 1).SaveAs(")
  );
  const worksheetExportSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.Worksheets(1).ExportAsFixedFormat(")
  );
  const worksheetFunctionExportSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.Worksheets(GetIndex()).ExportAsFixedFormat(")
  );
  const workbookSaveAsSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "ThisWorkbook.SaveAs("));
  const sheet1EvaluateSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Sheet1.Evaluate("));
  const sheet1SaveAsSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Sheet1.SaveAs("));
  const chart1SetSourceDataSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Chart1.SetSourceData(")
  );
  const workbookCloseSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "ActiveWorkbook.Close("));
  const workbookExportSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "ActiveWorkbook.ExportAsFixedFormat(")
  );
  const extractedZeroArgSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Application.CalculateFull("));
  const fallbackSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Application.OnTime("));
  const propertyFallbackSignature = service.getSignatureHelp(
    uri,
    findPositionAfterTokenInText(text, "Application.WorksheetFunction(")
  );
  const eventFallbackSignature = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Application.AfterCalculate("));
  const propertyFallbackSignature2 = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Application.ActiveCell("));
  const eventFallbackSignature2 = service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "Application.NewWorkbook("));
  const hover = service.getHover(uri, findPositionAfterTokenInText(text, "Debug.Print Application.Calcu"));
  const workbookHover = service.getHover(uri, findPositionAfterTokenInText(text, "Debug.Print ThisWorkbook.Save"));
  const chartHover = service.getHover(uri, findPositionAfterTokenInText(text, "Debug.Print Chart1.ChartA"));

  assert.equal(worksheetSignature?.activeParameter, 1);
  assert.equal(worksheetSignature?.label, "Sum(Arg1, Arg2, Arg3, ..., Arg30) As Double");
  assert.equal(worksheetSignature?.parameters.length, 30);
  assert.equal(worksheetSignature?.parameters[0]?.documentation?.includes("必須引数"), true);
  assert.equal(worksheetSignature?.parameters[1]?.documentation?.includes("想定型: Variant"), true);
  assert.equal(worksheetSignature?.parameters[1]?.documentation?.includes("省略可能"), true);
  assert.equal(worksheetSignature?.parameters[1]?.documentation?.includes("現在の引数型: Long"), true);
  assert.equal(chainedSignature?.label, "Sum(Arg1, Arg2, Arg3, ..., Arg30) As Double");
  assert.equal(powerSignature?.label.includes("Power("), true);
  assert.equal(powerSignature?.parameters.length, 2);
  assert.equal(averageSignature?.label, "Average(Arg1, Arg2, Arg3, ..., Arg30) As Double");
  assert.equal(averageSignature?.parameters[1]?.documentation?.includes("省略可能"), true);
  assert.equal(maxSignature?.label, "Max(Arg1, Arg2, Arg3, ..., Arg30) As Double");
  assert.equal(maxSignature?.parameters.length, 30);
  assert.equal(maxSignature?.parameters[0]?.documentation?.includes("想定型: Variant"), true);
  assert.equal(maxSignature?.parameters[0]?.documentation?.includes("必須引数"), true);
  assert.equal(maxSignature?.parameters[1]?.documentation?.includes("省略可能"), true);
  assert.equal(maxSignature?.parameters[29]?.documentation?.includes("省略可能"), true);
  assert.equal(minSignature?.label, "Min(Arg1, Arg2, Arg3, ..., Arg30) As Double");
  assert.equal(minSignature?.parameters.length, 30);
  assert.equal(minSignature?.parameters[0]?.documentation?.includes("想定型: Variant"), true);
  assert.equal(minSignature?.parameters[0]?.documentation?.includes("必須引数"), true);
  assert.equal(minSignature?.parameters[1]?.documentation?.includes("省略可能"), true);
  assert.equal(minSignature?.parameters[29]?.documentation?.includes("省略可能"), true);
  assert.equal(edateSignature?.label, "EDate(Arg1, Arg2) As Double");
  assert.equal(edateSignature?.parameters.length, 2);
  assert.equal(eomonthSignature?.label, "EoMonth(Arg1, Arg2) As Double");
  assert.equal(eomonthSignature?.parameters.length, 2);
  assert.equal(findSignature?.label, "Find(Arg1, Arg2, Arg3) As Double");
  assert.equal(findSignature?.parameters.length, 3);
  assert.equal(findSignature?.parameters[2]?.documentation?.includes("省略可能"), true);
  assert.equal(searchSignature?.label, "Search(Arg1, Arg2, Arg3) As Double");
  assert.equal(searchSignature?.parameters.length, 3);
  assert.equal(searchSignature?.parameters[2]?.documentation?.includes("省略可能"), true);
  assert.equal(andSignature?.label, "And(Arg1, Arg2, Arg3, ..., Arg30) As Boolean");
  assert.equal(andSignature?.parameters.length, 30);
  assert.equal(andSignature?.parameters[1]?.documentation?.includes("省略可能"), true);
  assert.equal(orSignature?.label, "Or(Arg1, Arg2, Arg3, ..., Arg30) As Boolean");
  assert.equal(orSignature?.parameters.length, 30);
  assert.equal(orSignature?.parameters[1]?.documentation?.includes("想定型: Variant"), true);
  assert.equal(orSignature?.parameters[1]?.documentation?.includes("省略可能"), true);
  assert.equal(xorSignature?.label, "Xor(Arg1, Arg2, Arg3, ..., Arg30) As Boolean");
  assert.equal(xorSignature?.parameters.length, 30);
  assert.equal(xorSignature?.parameters[1]?.documentation?.includes("想定型: Variant"), true);
  assert.equal(xorSignature?.parameters[1]?.documentation?.includes("省略可能"), true);
  assert.equal(countASignature?.label, "CountA(Arg1, Arg2, Arg3, ..., Arg30) As Double");
  assert.equal(countASignature?.parameters.length, 30);
  assert.equal(countASignature?.parameters[1]?.documentation?.includes("省略可能"), true);
  assert.equal(countBlankSignature?.label, "CountBlank(Arg1) As Double");
  assert.equal(countBlankSignature?.parameters.length, 1);
  assert.equal(textSignature?.label, "Text(Arg1, Arg2) As String");
  assert.equal(textSignature?.parameters.length, 2);
  assert.equal(vlookupSignature?.label, "VLookup(Arg1, Arg2, Arg3, Arg4) As Variant");
  assert.equal(vlookupSignature?.parameters[3]?.documentation?.includes("省略可能"), true);
  assert.equal(matchSignature?.label, "Match(Arg1, Arg2, Arg3) As Double");
  assert.equal(matchSignature?.parameters.length, 3);
  assert.equal(matchSignature?.parameters[2]?.documentation?.includes("省略可能"), true);
  assert.equal(indexSignature?.label, "Index(Arg1, Arg2, Arg3, Arg4) As Variant");
  assert.equal(indexSignature?.parameters.length, 4);
  assert.equal(indexSignature?.parameters[2]?.documentation?.includes("省略可能"), true);
  assert.equal(indexSignature?.parameters[3]?.documentation?.includes("省略可能"), true);
  assert.equal(lookupSignature?.label, "Lookup(Arg1, Arg2, Arg3) As Variant");
  assert.equal(lookupSignature?.parameters.length, 3);
  assert.equal(lookupSignature?.parameters[2]?.documentation?.includes("省略可能"), true);
  assert.equal(hlookupSignature?.label, "HLookup(Arg1, Arg2, Arg3, Arg4) As Variant");
  assert.equal(hlookupSignature?.parameters.length, 4);
  assert.equal(hlookupSignature?.parameters[3]?.documentation?.includes("省略可能"), true);
  assert.equal(chooseSignature?.label, "Choose(Arg1, Arg2, Arg3, ..., Arg30) As Variant");
  assert.equal(chooseSignature?.parameters.length, 30);
  assert.equal(chooseSignature?.parameters[1]?.documentation?.includes("想定型: Variant"), true);
  assert.equal(chooseSignature?.parameters[1]?.documentation?.includes("省略可能"), false);
  assert.equal(chooseSignature?.parameters[29]?.documentation?.includes("省略可能"), false);
  assert.equal(transposeSignature?.label, "Transpose(Arg1) As Variant");
  assert.equal(transposeSignature?.parameters.length, 1);
  assert.equal(transposeSignature?.parameters[0]?.documentation?.includes("必須引数"), true);
  assert.equal(addressSignature?.label, "Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo) As String");
  assert.equal(addressSignature?.parameters.length, 5);
  assert.equal(addressSignature?.parameters[0]?.documentation?.includes("省略可能"), true);
  assert.equal(addressSignature?.parameters[2]?.documentation?.includes("想定型: XlReferenceStyle"), true);
  assert.equal(addressSignature?.parameters[4]?.documentation?.includes("省略可能"), true);
  assert.equal(chainedAddressSignature?.label, "Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo) As String");
  assert.equal(chainedAddressSignature?.parameters.length, 5);
  assert.equal(chainedAddressSignature?.parameters[2]?.documentation?.includes("想定型: XlReferenceStyle"), true);
  assert.equal(addressLocalSignature?.label, "AddressLocal(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo) As String");
  assert.equal(addressLocalSignature?.parameters.length, 5);
  assert.equal(addressLocalSignature?.parameters[2]?.documentation?.includes("想定型: XlReferenceStyle"), true);
  assert.equal(addressLocalSignature?.parameters[4]?.documentation?.includes("省略可能"), true);
  assert.equal(worksheetEvaluateSignature?.label, "Evaluate(Name) As Variant");
  assert.equal(worksheetEvaluateSignature?.parameters.length, 1);
  assert.equal(worksheetEvaluateSignature?.parameters[0]?.documentation?.includes("必須引数"), true);
  assert.equal(worksheetStringEvaluateSignature?.label, "Evaluate(Name) As Variant");
  assert.equal(worksheetStringEvaluateSignature?.parameters[0]?.documentation?.includes("必須引数"), true);
  assert.equal(groupedWorksheetEvaluateSignature, undefined);
  assert.equal(
    worksheetSaveAsSignature?.label,
    "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local)"
  );
  assert.equal(worksheetSaveAsSignature?.parameters.length, 10);
  assert.equal(worksheetSaveAsSignature?.parameters[0]?.documentation?.includes("必須引数"), true);
  assert.equal(worksheetSaveAsSignature?.parameters[0]?.documentation?.includes("想定型: String"), true);
  assert.equal(
    worksheetExpressionSaveAsSignature?.label,
    "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local)"
  );
  assert.equal(worksheetExpressionSaveAsSignature?.parameters[0]?.documentation?.includes("想定型: String"), true);
  assert.equal(
    worksheetExportSignature?.label,
    "ExportAsFixedFormat(Type, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr)"
  );
  assert.equal(worksheetExportSignature?.parameters.length, 9);
  assert.equal(worksheetExportSignature?.parameters[0]?.documentation?.includes("必須引数"), true);
  assert.equal(worksheetExportSignature?.parameters[0]?.documentation?.includes("想定型: XlFixedFormatType"), true);
  assert.equal(worksheetFunctionExportSignature, undefined);
  assert.equal(
    workbookSaveAsSignature?.label,
    "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)"
  );
  assert.equal(sheet1EvaluateSignature?.label, "Evaluate(Name) As Variant");
  assert.equal(
    sheet1SaveAsSignature?.label,
    "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local)"
  );
  assert.equal(chart1SetSourceDataSignature?.label, "Chart.SetSourceData()");
  assert.equal(chart1SetSourceDataSignature?.documentation?.includes("excel.chart.setsourcedata"), true);
  assert.equal(workbookSaveAsSignature?.parameters.length, 12);
  assert.equal(workbookSaveAsSignature?.parameters[0]?.documentation?.includes("省略可能"), true);
  assert.equal(workbookSaveAsSignature?.parameters[6]?.documentation?.includes("想定型: XlSaveAsAccessMode"), true);
  assert.equal(workbookSaveAsSignature?.parameters[7]?.documentation?.includes("想定型: XlSaveConflictResolution"), true);
  assert.equal(workbookCloseSignature?.label, "Close(SaveChanges, FileName, RouteWorkbook)");
  assert.equal(workbookCloseSignature?.parameters.length, 3);
  assert.equal(workbookCloseSignature?.parameters[0]?.documentation?.includes("省略可能"), true);
  assert.equal(
    workbookExportSignature?.label,
    "ExportAsFixedFormat(Type, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr)"
  );
  assert.equal(workbookExportSignature?.parameters.length, 9);
  assert.equal(workbookExportSignature?.parameters[0]?.documentation?.includes("必須引数"), true);
  assert.equal(workbookExportSignature?.parameters[0]?.documentation?.includes("想定型: XlFixedFormatType"), true);
  assert.equal(extractedZeroArgSignature?.label, "CalculateFull()");
  assert.equal(extractedZeroArgSignature?.parameters.length, 0);
  assert.equal(fallbackSignature?.label, "Application.OnTime()");
  assert.equal(fallbackSignature?.parameters.length, 0);
  assert.equal(fallbackSignature?.documentation?.includes("excel.application.ontime"), true);
  assert.equal(propertyFallbackSignature, undefined);
  assert.equal(eventFallbackSignature, undefined);
  assert.equal(propertyFallbackSignature2, undefined);
  assert.equal(eventFallbackSignature2, undefined);
  assert.equal(hover?.contents.includes("Calculate()"), true);
  assert.equal(hover?.contents.includes("Calculates all open workbooks"), true);
  assert.equal(hover?.contents.includes("Microsoft Learn"), true);
  assert.equal(
    workbookHover?.contents.includes(
      "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)"
    ),
    true
  );
  assert.equal(workbookHover?.contents.includes("excel.workbook.saveas"), true);
  assert.equal(chartHover?.contents.includes("Chart.ChartArea"), true);
  assert.equal(chartHover?.contents.includes("excel.chart.chartarea"), true);
});

test("document service prioritizes built-in member signature help over workspace callable collisions", () => {
  const service = createDocumentService();
  const consumerUri = "file:///C:/temp/BuiltInCollisionConsumer.bas";
  const collisionUri = "file:///C:/temp/BuiltInCollisionDefinitions.bas";

  service.analyzeText(
    collisionUri,
    "vba",
    1,
    `Attribute VB_Name = "BuiltInCollisionDefinitions"
Option Explicit

Public Function Sum(ByVal value As Long) As Long
    Sum = value
End Function`
  );
  service.analyzeText(
    consumerUri,
    "vba",
    1,
    `Attribute VB_Name = "BuiltInCollisionConsumer"
Option Explicit

Public Sub Demo()
    Debug.Print WorksheetFunction.Sum(1, 2)
End Sub`
  );

  const signature = service.getSignatureHelp(consumerUri, { character: 42, line: 4 });

  assert.equal(signature?.label, "Sum(Arg1, Arg2, Arg3, ..., Arg30) As Double");
  assert.equal(signature?.documentation?.includes("excel.worksheetfunction.sum"), true);
});

test("document service keeps built-in signature and hover conservative for shadowed roots", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/BuiltInSignatureShadowed.bas";
  const text = `Attribute VB_Name = "BuiltInSignatureShadowed"
Option Explicit

Public Sub Demo()
    Dim WorksheetFunction As String
    Dim ActiveCell As String
    Debug.Print WorksheetFunction.Sum(1, 2)
    Debug.Print ActiveCell.Address(False, False)
End Sub`;

  service.analyzeText(uri, "vba", 1, text);

  assert.equal(service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Sum(")), undefined);
  assert.equal(service.getHover(uri, findPositionAfterTokenInText(text, "WorksheetFunction.Su")), undefined);
  assert.equal(service.getSignatureHelp(uri, findPositionAfterTokenInText(text, "ActiveCell.Address(")), undefined);
  assert.equal(service.getHover(uri, findPositionAfterTokenInText(text, "ActiveCell.Addre")), undefined);
});

test("document service formats VBA indentation through the shared core formatter", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/FormatDocument.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
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

  const formatted = service.formatDocument(uri, { insertSpaces: true, tabSize: 4 });

  assert.equal(
    formatted,
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
});

test("document service formats continued assignments, argument lists, and method chains", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/ContinuationFormatting.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "ContinuationFormatting"
Option Explicit

Public Sub Demo()
Dim message As String
message =   _
"prefix" &  _
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

  const formatted = service.formatDocument(uri, { insertSpaces: true, tabSize: 4 });

  assert.equal(
    formatted,
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
});

test("document service expands compressed block layout through the shared core formatter", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/BlockLayoutFormatting.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "BlockLayoutFormatting"
Option Explicit

Public Sub Demo()
Dim value As Long: value = 0
If value = 0 Then: Debug.Print "zero": ElseIf value = 1 Then: Debug.Print "one": Else: Debug.Print "other": End If
Select Case value: Case 0: Debug.Print "case zero": Case Else: With Application: .StatusBar = "fallback": End With: End Select
#If VBA7 Then: value = value + 1: #Else: value = value - 1: #End If
End Sub`
  );

  const formatted = service.formatDocument(uri, { insertSpaces: true, tabSize: 4 });

  assert.equal(
    formatted,
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
});

test("document service aligns declaration blocks through the shared core formatter", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/DeclarationAlignment.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "DeclarationAlignment"
Option Explicit

Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr

Public Sub Demo()
Dim title As String
Dim count       As Long
Dim enabled As Boolean

Const DefaultTitle As String = "Ready"
Const RetryCount  As Long=3
Const IsEnabled As Boolean   = True

Debug.Print title, count, enabled
End Sub`
  );

  const formatted = service.formatDocument(uri, { insertSpaces: true, tabSize: 4 });

  assert.equal(
    formatted,
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
});

test("document service normalizes comment spacing through the shared core formatter", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/CommentFormatting.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "CommentFormatting"
Option Explicit

Public Sub Demo()
'leading comment
Dim value As Long'counter
If True Then'true branch
'inner comment
value = 1'updated
Rem    status
#If VBA7 Then'requires vba7
'conditional comment
#Else'fallback path
Rem    fallback comment
#End If
End If
End Sub`
  );

  const formatted = service.formatDocument(uri, { insertSpaces: true, tabSize: 4 });

  assert.equal(
    formatted,
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
});

test("document service keeps ambiguous cross-file symbols conservative", () => {
  const service = createDocumentService();
  const consumerUri = "file:///C:/temp/Consumer.bas";

  service.analyzeText(
    "file:///C:/temp/PublicApi.bas",
    "vba",
    1,
    `Attribute VB_Name = "PublicApi"
Option Explicit

Public Function PublicMessage() As String
    PublicMessage = "Hello"
End Function`
  );
  service.analyzeText(
    "file:///C:/temp/AnotherApi.bas",
    "vba",
    1,
    `Attribute VB_Name = "AnotherApi"
Option Explicit

Public Function PublicMessage() As String
    PublicMessage = "World"
End Function`
  );
  service.analyzeText(
    consumerUri,
    "vba",
    1,
    `Attribute VB_Name = "Consumer"
Option Explicit

Public Sub UseLibrary()
    Dim message As String
    message = PublicMessage()
End Sub`
  );

  assert.equal(service.getDefinition(consumerUri, { character: 18, line: 5 }), undefined);
  assert.equal(service.getDiagnostics(consumerUri).some((diagnostic) => diagnostic.code === "undeclared-variable"), true);
  assert.deepEqual(service.getReferences(consumerUri, { character: 18, line: 5 }, true), []);
});

test("document service keeps local references scoped when module names are shadowed", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/Shadowing.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "Shadowing"
Option Explicit

Public Const SharedValue As Long = 1

Public Sub Demo()
    Dim SharedValue As Long
    SharedValue = 2
End Sub`
  );

  const definition = service.getDefinition(uri, { character: 8, line: 6 });
  const references = service.getReferences(uri, { character: 8, line: 6 }, true);

  assert.equal(definition?.symbol.scope, "procedure");
  assert.equal(definition?.symbol.kind, "variable");
  assert.deepEqual(
    references.map((reference) => `${reference.uri}:${reference.range.start.line}:${reference.range.start.character}`),
    [`${uri}:6:8`, `${uri}:7:4`]
  );
});

test("document service prepares safe local rename edits within one procedure", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/RenameLocal.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "RenameLocal"
Option Explicit

Public Sub Demo()
    Dim totalCount As Long
    Dim message As String
    totalCount = 1
    message = totalCount
    Debug.Print totalCount
End Sub

Public Sub OtherDemo()
    Dim totalCount As Long
    totalCount = 2
End Sub`
  );

  const target = service.prepareRename(uri, { character: 6, line: 6 });
  const edits = service.getRenameEdits(uri, { character: 6, line: 6 }, "currentCount");

  assert.equal(target?.placeholder, "totalCount");
  assert.equal(`${target?.range.start.line}:${target?.range.start.character}`, "6:4");
  assert.deepEqual(
    edits?.map((edit) => `${edit.uri}:${edit.range.start.line}:${edit.range.start.character}:${edit.newText}`),
    [
      `${uri}:4:8:currentCount`,
      `${uri}:6:4:currentCount`,
      `${uri}:7:14:currentCount`,
      `${uri}:8:16:currentCount`
    ]
  );
});

test("document service rejects unsafe local rename targets and names", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/RenameLocal.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "RenameLocal"
Option Explicit

Public Sub Demo(ByVal inputValue As Long)
    Dim totalCount As Long
    Dim message As String
    totalCount = inputValue
    message = totalCount
End Sub`
  );

  assert.equal(service.prepareRename(uri, { character: 20, line: 3 }), undefined);
  assert.equal(service.getRenameEdits(uri, { character: 6, line: 6 }, "message"), undefined);
  assert.equal(service.getRenameEdits(uri, { character: 6, line: 6 }, "Sub"), undefined);
});

test("document service exposes semantic tokens for declarations and references", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/SemanticTokens.bas";
  const text = `Attribute VB_Name = "SemanticTokens"
Option Explicit

Private Type CustomerRecord
    Name As String
End Type

Private Const DefaultName As String = "A"

Public Function BuildCustomer(ByVal sourceName As String) As CustomerRecord
    Dim customer As CustomerRecord
    customer.Name = sourceName
    BuildCustomer = customer
End Function

Public Sub Demo()
    Dim current As CustomerRecord
    current = BuildCustomer(DefaultName)
End Sub`;

  service.analyzeText(uri, "vba", 1, text);

  const tokens = service.getSemanticTokens(uri);

  assertSemanticToken(text, tokens, 3, "CustomerRecord", {
    modifiers: ["declaration"],
    type: "type"
  });
  assertSemanticToken(text, tokens, 7, "DefaultName", {
    modifiers: ["declaration", "readonly"],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 9, "BuildCustomer", {
    modifiers: ["declaration"],
    type: "function"
  });
  assertSemanticToken(text, tokens, 9, "sourceName", {
    modifiers: ["declaration"],
    type: "parameter"
  });
  assertSemanticToken(text, tokens, 9, "CustomerRecord", {
    modifiers: [],
    type: "type"
  });
  assertSemanticToken(text, tokens, 10, "customer", {
    modifiers: ["declaration"],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 10, "CustomerRecord", {
    modifiers: [],
    type: "type"
  });
  assertSemanticToken(text, tokens, 11, "sourceName", {
    modifiers: [],
    type: "parameter"
  });
  assertSemanticToken(text, tokens, 16, "current", {
    modifiers: ["declaration"],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 17, "BuildCustomer", {
    modifiers: [],
    type: "function"
  });
  assertSemanticToken(text, tokens, 17, "DefaultName", {
    modifiers: ["readonly"],
    type: "variable"
  });
});

test("document service exposes semantic tokens for built-in keywords, functions, constants, and members", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/BuiltInSemantic.bas";
  const thisWorkbookUri = "file:///C:/temp/ThisWorkbook.cls";
  const sheet1Uri = "file:///C:/temp/Sheet1.cls";
  const chart1Uri = "file:///C:/temp/Chart1.cls";
  const text = `Attribute VB_Name = "BuiltInSemantic"
Option Explicit

Public Sub Demo()
    Beep
    MsgBox xlAll
    Debug.Print Application.Name
    Debug.Print Application.WorksheetFunction.Sum(1, 2)
    Debug.Print Worksheets("A(1)").Evaluate("A1")
    Debug.Print Worksheets(Array("Sheet1", "Sheet2")).Evaluate("A1")
    Debug.Print ThisWorkbook.SaveAs
    Debug.Print Sheet1.Evaluate("A1")
    Call Chart1.SetSourceData(Range("A1:B2"))
    Debug.Print Application.ActiveCell.Address
End Sub`;

  service.analyzeText(
    thisWorkbookUri,
    "vba",
    1,
    `Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(
    sheet1Uri,
    "vba",
    1,
    `Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(
    chart1Uri,
    "vba",
    1,
    `Attribute VB_Name = "Chart1"
Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
  );
  service.analyzeText(uri, "vba", 1, text);

  const tokens = service.getSemanticTokens(uri);

  assertSemanticToken(text, tokens, 4, "Beep", {
    modifiers: [],
    type: "keyword"
  });
  assertSemanticToken(text, tokens, 5, "MsgBox", {
    modifiers: [],
    type: "function"
  });
  assertSemanticToken(text, tokens, 5, "xlAll", {
    modifiers: ["readonly"],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 6, "Application", {
    modifiers: [],
    type: "type"
  });
  assertSemanticToken(text, tokens, 6, "Name", {
    modifiers: [],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 7, "WorksheetFunction", {
    modifiers: [],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 7, "Sum", {
    modifiers: [],
    type: "function"
  });
  assertSemanticToken(text, tokens, 8, "Evaluate", {
    modifiers: [],
    type: "function"
  });
  assertNoSemanticToken(text, tokens, 9, "Evaluate");
  assertSemanticToken(text, tokens, 10, "ThisWorkbook", {
    modifiers: [],
    type: "variable"
  });
  assertSemanticToken(text, tokens, 10, "SaveAs", {
    modifiers: [],
    type: "function"
  });
  assertSemanticToken(text, tokens, 11, "Evaluate", {
    modifiers: [],
    type: "function"
  });
  assertSemanticToken(text, tokens, 12, "SetSourceData", {
    modifiers: [],
    type: "function"
  });
  assertSemanticToken(text, tokens, 13, "Address", {
    modifiers: [],
    type: "variable"
  });
});

test("document service exposes inferred type mismatch diagnostics", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/Mismatch.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "Mismatch"
Option Explicit

Public Sub Demo()
    Dim title As String
    title = True
End Sub`
  );

  const diagnostics = service.getDiagnostics(uri).filter((diagnostic) => diagnostic.code === "type-mismatch");

  assert.equal(diagnostics.length, 1);
  assert.equal(diagnostics[0]?.severity, "warning");
});

test("document service exposes type mismatch diagnostics for continued assignments", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/ContinuedMismatch.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "ContinuedMismatch"
Option Explicit

Public Sub Demo()
    Dim title As String
    title = _
        True
End Sub`
  );

  const diagnostics = service.getDiagnostics(uri).filter((diagnostic) => diagnostic.code === "type-mismatch");

  assert.equal(diagnostics.length, 1);
  assert.equal(diagnostics[0]?.message, "Type mismatch: cannot assign Boolean to String.");
});

test("document service exposes expanded type mismatch diagnostics for compound and Set assignments", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/ExpandedMismatch.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "ExpandedMismatch"
Option Explicit

Public Sub Demo()
    Dim title As String
    Dim count As Long
    Dim flag As Boolean
    Dim holder As Object
    Dim loose As Variant
    title = 1 + 2
    count = "1" & loose
    flag = 1 < 2
    Set holder = Nothing
    Set holder = 1
End Sub`
  );

  const diagnostics = service.getDiagnostics(uri).filter((diagnostic) => diagnostic.code === "type-mismatch");

  assert.equal(diagnostics.length, 3);
  assert.deepEqual(
    diagnostics.map((diagnostic) => diagnostic.message),
    [
      "Type mismatch: cannot assign Long to String.",
      "Type mismatch: cannot assign String to Long.",
      "Type mismatch: cannot assign Long to Object."
    ]
  );
});

test("document service augments diagnostics for cross-file ByRef argument risks", () => {
  const service = createDocumentService();
  const libraryUri = "file:///C:/temp/PublicByRefApi.bas";
  const consumerUri = "file:///C:/temp/PublicByRefConsumer.bas";

  service.analyzeText(
    libraryUri,
    "vba",
    1,
    `Attribute VB_Name = "PublicByRefApi"
Option Explicit

Public Sub UpdateCount(ByRef count As Long)
End Sub`
  );
  service.analyzeText(
    consumerUri,
    "vba",
    1,
    `Attribute VB_Name = "PublicByRefConsumer"
Option Explicit

Public Sub Demo()
    Dim wrongCount As String
    UpdateCount wrongCount
End Sub`
  );

  const diagnostics = service.getDiagnostics(consumerUri).filter((diagnostic) => diagnostic.code.startsWith("byref-"));

  assert.equal(diagnostics.length, 1);
  assert.equal(
    diagnostics[0]?.message,
    "ByRef parameter 'count' in UpdateCount expects Long but receives String. VBA may raise a ByRef argument type mismatch."
  );
});

test("document service exposes set-required diagnostics for local object assignments", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/SetRequired.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "SetRequired"
Option Explicit

Private Function BuildItems() As Collection
    Set BuildItems = New Collection
End Function

Public Sub Demo()
    Dim items As Collection
    Dim holder As Object
    items = New Collection
    holder = BuildItems()
    items = Nothing
    Set items = New Collection
End Sub`
  );

  const diagnostics = service.getDiagnostics(uri).filter((diagnostic) => diagnostic.code === "set-required");

  assert.equal(diagnostics.length, 3);
  assert.deepEqual(
    diagnostics.map((diagnostic) => diagnostic.message),
    [
      "Set is required to assign Collection to Collection.",
      "Set is required to assign Collection to Object.",
      "Set is required to assign Nothing to Collection."
    ]
  );
});

test("document service exposes duplicate-definition diagnostics", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/Duplicates.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "Duplicates"
Option Explicit

Private Type CustomerRecord
    Id As Long
End Type

Private Type CustomerRecord
    Name As String
End Type

Public Enum StatusKind
    StatusOpen = 1
End Enum

Public Enum StatusKind
    StatusClosed = 2
End Enum

Private Sub SharedName()
End Sub

Private Sub SharedName()
End Sub

Public Sub Demo(ByVal value As Long)
    Dim value As Long
    Const title As String = "A"
    Const title As String = "B"
End Sub`
  );

  const diagnostics = service.getDiagnostics(uri).filter((diagnostic) => diagnostic.code === "duplicate-definition");

  assert.equal(diagnostics.length, 5);
  assert.deepEqual(
    diagnostics.map((diagnostic) => diagnostic.message),
    [
      "Duplicate definition 'CustomerRecord' in module scope.",
      "Duplicate definition 'StatusKind' in module scope.",
      "Duplicate definition 'SharedName' in module scope.",
      "Duplicate definition 'value' in procedure 'Demo'.",
      "Duplicate definition 'title' in procedure 'Demo'."
    ]
  );
});

test("document service exposes unreachable-code diagnostics conservatively", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/Unreachable.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "Unreachable"
Option Explicit

Public Sub Demo()
    Dim ready As Boolean
    Dim keepRunning As Boolean
    Dim marker As Long
    Exit Sub
    marker = 1
JumpHere:
    marker = 6

    If ready Then
        Exit Sub
        marker = 2
    Else
        marker = 3
    End If

    Do While keepRunning
        End
        marker = 4
    Loop

    marker = 5
End Sub`
  );

  const diagnostics = service.getDiagnostics(uri).filter((diagnostic) => diagnostic.code === "unreachable-code");

  assert.equal(diagnostics.length, 3);
  assert.deepEqual(
    diagnostics.map((diagnostic) => diagnostic.message),
    [
      "Unreachable code after Exit Sub.",
      "Unreachable code after Exit Sub.",
      "Unreachable code after End."
    ]
  );
});

test("document service exposes unused-variable diagnostics for locals and parameters", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/UnusedLocals.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "UnusedLocals"
Option Explicit

Public Sub Demo(ByVal usedArg As Long, ByVal unusedArg As Long)
    Dim usedValue As Long
    Dim writeOnlyValue As Long
    Dim unusedValue As String
    usedValue = usedArg
    writeOnlyValue = 1
    Debug.Print usedValue
End Sub`
  );

  const diagnostics = service.getDiagnostics(uri).filter((diagnostic) => diagnostic.code === "unused-variable");

  assert.equal(diagnostics.length, 2);
  assert.deepEqual(
    diagnostics.map((diagnostic) => diagnostic.message),
    [
      "Unused local declaration 'unusedArg'.",
      "Unused local declaration 'unusedValue'."
    ]
  );
});

test("document service exposes write-only-variable diagnostics for assigned-only locals", () => {
  const service = createDocumentService();
  const uri = "file:///C:/temp/WriteOnlyLocals.bas";

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "WriteOnlyLocals"
Option Explicit

Public Sub Demo()
    Dim readValue As Long
    Dim writeOnlyValue As Long
    Dim objectHolder As Collection
    readValue = 1
    writeOnlyValue = readValue
    Set objectHolder = New Collection
    Debug.Print readValue
End Sub`
  );

  const diagnostics = service.getDiagnostics(uri).filter((diagnostic) => diagnostic.code === "write-only-variable");
  const unusedDiagnostics = service.getDiagnostics(uri).filter((diagnostic) => diagnostic.code === "unused-variable");

  assert.equal(diagnostics.length, 2);
  assert.deepEqual(
    diagnostics.map((diagnostic) => diagnostic.message),
    [
      "Write-only local variable 'writeOnlyValue'.",
      "Write-only local variable 'objectHolder'."
    ]
  );
  assert.equal(unusedDiagnostics.length, 0);
});

test("document service loads nearest worksheet control metadata sidecar as read-only state", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "samples", "book1");
  const moduleDirectory = path.join(bundleRoot, "modules");
  const logs = [];

  mkdirSync(moduleDirectory, { recursive: true });
  writeWorksheetControlMetadataSidecar(workspaceRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [],
    version: 1,
    workbook: {
      name: "outer.xlsm",
      sourceKind: "openxml-package"
    }
  });
  writeWorksheetControlMetadataSidecar(bundleRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [
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
      },
      {
        ownerKind: "chartsheet",
        reason: "chart-sheet-metadata-unproven",
        sheetCodeName: "Chart1",
        sheetName: "Chart1",
        status: "unsupported"
      }
    ],
    version: 1,
    workbook: {
      name: "inner.xlsm",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({
      logger: (entry) => logs.push(entry),
      workspaceRoots: [workspaceRoot]
    });
    const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
    const state = service.analyzeText(
      uri,
      "vba",
      1,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );

    assert.equal(state.worksheetControlMetadata?.status, "loaded");
    assert.equal(state.worksheetControlMetadata?.bundleRoot, bundleRoot);
    assert.equal(state.worksheetControlMetadata?.workbookName, "inner.xlsm");
    assert.equal(state.worksheetControlMetadata?.supportedOwners.length, 1);
    assert.equal(state.worksheetControlMetadata?.supportedOwners[0]?.sheetCodeName, "Sheet1");
    assert.equal(state.worksheetControlMetadata?.supportedOwners[0]?.controls[0]?.shapeName, "CheckBox1");
    assert.equal(logs.some((entry) => entry.code === "worksheet-control-metadata.loaded"), true);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service は workspaceRoots 未指定時に worksheet control metadata sidecar を読まない", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const moduleDirectory = path.join(workspaceRoot, "src");
  const logs = [];

  mkdirSync(moduleDirectory, { recursive: true });
  writeWorksheetControlMetadataSidecar(workspaceRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [],
    version: 1,
    workbook: {
      name: "book1.xlsm",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({
      logger: (entry) => logs.push(entry)
    });
    const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
    const state = service.analyzeText(
      uri,
      "vba",
      1,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );

    assert.equal(state.worksheetControlMetadata, undefined);
    assert.equal(logs.some((entry) => entry.code === "worksheet-control-metadata.loaded"), false);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service reloads worksheet control metadata sidecar when file stats change", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const moduleDirectory = path.join(workspaceRoot, "src");
  const sidecarPath = path.join(workspaceRoot, ".vba", "worksheet-control-metadata.json");
  const logs = [];

  mkdirSync(moduleDirectory, { recursive: true });
  writeWorksheetControlMetadataSidecar(workspaceRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [
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
    ],
    version: 1,
    workbook: {
      name: "before.xlsm",
      sourceKind: "openxml-package"
    }
  });
  utimesSync(sidecarPath, new Date("2026-03-14T00:00:01.000Z"), new Date("2026-03-14T00:00:01.000Z"));

  try {
    const service = createDocumentService({
      logger: (entry) => logs.push(entry),
      workspaceRoots: [workspaceRoot]
    });
    const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;

    const firstState = service.analyzeText(
      uri,
      "vba",
      1,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );

    writeWorksheetControlMetadataSidecar(workspaceRoot, {
      artifact: "worksheet-control-metadata-sidecar",
      owners: [
        {
          controls: [
            {
              codeName: "chkFinished",
              controlType: "CheckBox",
              progId: "Forms.CheckBox.1",
              shapeId: 3,
              shapeName: "CheckBox1"
            },
            {
              codeName: "cmdApply",
              controlType: "CommandButton",
              progId: "Forms.CommandButton.1",
              shapeId: 4,
              shapeName: "CommandButton1"
            }
          ],
          ownerKind: "worksheet",
          sheetCodeName: "Sheet1",
          sheetName: "Sheet1",
          status: "supported"
        }
      ],
      version: 1,
      workbook: {
        name: "after-longer.xlsm",
        sourceKind: "openxml-package"
      }
    });
    utimesSync(sidecarPath, new Date("2026-03-14T00:00:05.000Z"), new Date("2026-03-14T00:00:05.000Z"));

    const secondState = service.analyzeText(
      uri,
      "vba",
      2,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );
    const thirdState = service.analyzeText(
      uri,
      "vba",
      3,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );

    assert.equal(firstState.worksheetControlMetadata?.workbookName, "before.xlsm");
    assert.equal(firstState.worksheetControlMetadata?.supportedOwners[0]?.controls.length, 1);
    assert.equal(secondState.worksheetControlMetadata?.workbookName, "after-longer.xlsm");
    assert.equal(secondState.worksheetControlMetadata?.supportedOwners[0]?.controls.length, 2);
    assert.equal(thirdState.worksheetControlMetadata?.workbookName, "after-longer.xlsm");
    assert.equal(thirdState.worksheetControlMetadata?.supportedOwners[0]?.controls.length, 2);
    assert.equal(logs.filter((entry) => entry.code === "worksheet-control-metadata.loaded").length, 2);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service sidecar lookup は workspace root を越えない", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const parentRoot = path.join(temporaryDirectory, "outside");
  const workspaceRoot = path.join(parentRoot, "workspace");
  const moduleDirectory = path.join(workspaceRoot, "src");

  mkdirSync(moduleDirectory, { recursive: true });
  writeWorksheetControlMetadataSidecar(parentRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [],
    version: 1,
    workbook: {
      name: "outside.xlsm",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });
    const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
    const state = service.analyzeText(
      uri,
      "vba",
      1,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );

    assert.equal(state.worksheetControlMetadata, undefined);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service ignores invalid worksheet control metadata sidecar and keeps analysis alive", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const moduleDirectory = path.join(workspaceRoot, "src");
  const logs = [];

  mkdirSync(path.join(workspaceRoot, ".vba"), { recursive: true });
  mkdirSync(moduleDirectory, { recursive: true });
  writeFileSync(
    path.join(workspaceRoot, ".vba", "worksheet-control-metadata.json"),
    `{
      "version": 2,
      "artifact": "wrong-artifact",
      "owners": []
    }\n`
  );

  try {
    const service = createDocumentService({
      logger: (entry) => logs.push(entry),
      workspaceRoots: [workspaceRoot]
    });
    const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
    const state = service.analyzeText(
      uri,
      "vba",
      1,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );

    assert.equal(state.analysis.module.name, "Module1");
    assert.equal(state.worksheetControlMetadata?.status, "ignored");
    assert.equal(state.worksheetControlMetadata?.supportedOwners.length, 0);
    assert.equal(logs.some((entry) => entry.code === "worksheet-control-metadata.invalid-version"), true);
    assert.equal(logs.some((entry) => entry.code === "worksheet-control-metadata.invalid-artifact"), true);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service は workspace root 変更時に worksheet control metadata sidecar state を再解決する", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const moduleDirectory = path.join(workspaceRoot, "src");

  mkdirSync(moduleDirectory, { recursive: true });
  writeWorksheetControlMetadataSidecar(workspaceRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [],
    version: 1,
    workbook: {
      name: "book1.xlsm",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });
    const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
    const state = service.analyzeText(
      uri,
      "vba",
      1,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );

    assert.equal(state.worksheetControlMetadata?.workbookName, "book1.xlsm");

    service.setWorkspaceRoots([]);

    assert.equal(service.getState(uri)?.worksheetControlMetadata, undefined);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service caches active workbook identity snapshot and logs manifest match / mismatch", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-active-workbook-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleA = path.join(workspaceRoot, "bundle-a");
  const bundleB = path.join(workspaceRoot, "bundle-b");
  const moduleAUri = pathToFileURL(path.join(bundleA, "Module1.bas")).href;
  const moduleBUri = pathToFileURL(path.join(bundleB, "Module2.bas")).href;
  const logs = [];

  mkdirSync(bundleA, { recursive: true });
  mkdirSync(bundleB, { recursive: true });
  writeWorkbookBindingManifest(bundleA, {
    artifact: "workbook-binding-manifest",
    bindingKind: "active-workbook-fullname",
    version: 1,
    workbook: {
      fullName: "C:\\Work\\Book1.xlsm",
      isAddIn: false,
      name: "Book1.xlsm",
      path: "C:\\Work",
      sourceKind: "openxml-package"
    }
  });
  writeWorkbookBindingManifest(bundleB, {
    artifact: "workbook-binding-manifest",
    bindingKind: "active-workbook-fullname",
    version: 1,
    workbook: {
      fullName: "C:\\Work\\OtherBook.xlsm",
      isAddIn: false,
      name: "OtherBook.xlsm",
      path: "C:\\Work",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({
      logger: (entry) => logs.push(entry),
      workspaceRoots: [workspaceRoot]
    });

    service.setActiveWorkbookIdentitySnapshot({
      identity: {
        fullName: "c:/work/BOOK1.xlsm",
        isAddin: false,
        name: "Book1.xlsm",
        path: "c:/work"
      },
      observedAt: "2026-03-21T00:00:00.000Z",
      providerKind: "excel-active-workbook",
      state: "available",
      version: 1
    });

    const stateA = service.analyzeText(
      moduleAUri,
      "vba",
      1,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );
    const stateB = service.analyzeText(
      moduleBUri,
      "vba",
      1,
      `Attribute VB_Name = "Module2"
Option Explicit`
    );

    assert.equal(stateA.activeWorkbookIdentity?.state, "available");
    assert.equal(stateA.workbookBindingManifest?.status, "loaded");
    assert.equal(stateA.workbookBindingManifest?.workbookName, "Book1.xlsm");
    assert.equal(stateB.workbookBindingManifest?.workbookName, "OtherBook.xlsm");
    assert.equal(logs.some((entry) => entry.code === "active-workbook-identity.updated"), true);
    assert.equal(logs.some((entry) => entry.code === "active-workbook-identity.match"), true);
    assert.equal(logs.some((entry) => entry.code === "active-workbook-identity.mismatch"), true);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service logs binding-missing when active workbook is available but manifest is absent", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-active-workbook-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const moduleDirectory = path.join(workspaceRoot, "src");
  const logs = [];

  mkdirSync(moduleDirectory, { recursive: true });

  try {
    const service = createDocumentService({
      logger: (entry) => logs.push(entry),
      workspaceRoots: [workspaceRoot]
    });
    const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;

    service.setActiveWorkbookIdentitySnapshot({
      identity: {
        fullName: "C:\\Work\\Book1.xlsm",
        isAddin: false,
        name: "Book1.xlsm",
        path: "C:\\Work"
      },
      observedAt: "2026-03-21T00:00:00.000Z",
      providerKind: "excel-active-workbook",
      state: "available",
      version: 1
    });

    const state = service.analyzeText(
      uri,
      "vba",
      1,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );

    assert.equal(state.workbookBindingManifest, undefined);
    assert.equal(logs.some((entry) => entry.code === "active-workbook-identity.binding-missing"), true);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service logs binding-disabled when active workbook snapshot is unavailable", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-active-workbook-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "bundle");
  const logs = [];

  mkdirSync(bundleRoot, { recursive: true });
  writeWorkbookBindingManifest(bundleRoot, {
    artifact: "workbook-binding-manifest",
    bindingKind: "active-workbook-fullname",
    version: 1,
    workbook: {
      fullName: "C:\\Work\\Book1.xlsm",
      isAddIn: false,
      name: "Book1.xlsm",
      path: "C:\\Work",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({
      logger: (entry) => logs.push(entry),
      workspaceRoots: [workspaceRoot]
    });
    const uri = pathToFileURL(path.join(bundleRoot, "Module1.bas")).href;

    service.setActiveWorkbookIdentitySnapshot({
      observedAt: "2026-03-21T00:00:00.000Z",
      providerKind: "excel-active-workbook",
      reason: "host-unreachable",
      state: "unavailable",
      version: 1
    });

    const state = service.analyzeText(
      uri,
      "vba",
      1,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );
    service.analyzeText(
      uri,
      "vba",
      2,
      `Attribute VB_Name = "Module1"
Option Explicit
' edit`
    );

    assert.equal(state.activeWorkbookIdentity?.state, "unavailable");
    assert.equal(logs.some((entry) => entry.code === "active-workbook-identity.unavailable"), true);
    assert.equal(logs.filter((entry) => entry.code === "active-workbook-identity.binding-disabled").length, 1);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

test("document service rejects invalid active workbook identity payloads", () => {
  const service = createDocumentService();

  service.setActiveWorkbookIdentitySnapshot({
    identity: {
      fullName: "",
      isAddin: "no",
      name: "Book1.xlsm"
    },
    observedAt: "not-a-date",
    providerKind: "excel-active-workbook",
    state: "available",
    version: 1
  });

  const state = service.analyzeText(
    "file:///C:/temp/Module1.bas",
    "vba",
    1,
    `Attribute VB_Name = "Module1"
Option Explicit`
  );

  assert.equal(state.activeWorkbookIdentity?.state, "invalid");
  assert.equal(state.activeWorkbookIdentity?.issues.some((issue) => issue.path === "$.identity.fullName"), true);
});

test("document service rejects available snapshot and manifest when unsaved / add-in values are mixed in", () => {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-active-workbook-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "bundle");
  const logs = [];

  mkdirSync(bundleRoot, { recursive: true });
  writeWorkbookBindingManifest(bundleRoot, {
    artifact: "workbook-binding-manifest",
    bindingKind: "active-workbook-fullname",
    version: 1,
    workbook: {
      fullName: "Addin.xlam",
      isAddIn: true,
      name: "Addin.xlam",
      path: "",
      sourceKind: "openxml-package"
    }
  });

  try {
    const service = createDocumentService({
      logger: (entry) => logs.push(entry),
      workspaceRoots: [workspaceRoot]
    });
    const uri = pathToFileURL(path.join(bundleRoot, "Module1.bas")).href;

    service.setActiveWorkbookIdentitySnapshot({
      identity: {
        fullName: "C:\\Work\\Addin.xlam",
        isAddin: true,
        name: "Addin.xlam",
        path: ""
      },
      observedAt: "2026-03-21T00:00:00.000Z",
      providerKind: "excel-active-workbook",
      state: "available",
      version: 1
    });

    const state = service.analyzeText(
      uri,
      "vba",
      1,
      `Attribute VB_Name = "Module1"
Option Explicit`
    );

    assert.equal(state.activeWorkbookIdentity?.state, "invalid");
    assert.equal(state.workbookBindingManifest?.status, "ignored");
    assert.equal(logs.some((entry) => entry.code === "active-workbook-identity.match"), false);
    assert.equal(logs.some((entry) => entry.code === "active-workbook-identity.invalid-payload"), true);
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});

function assertSemanticToken(text, tokens, lineIndex, identifier, expected, occurrence = 0) {
  const lines = text.split("\n");
  const line = lines[lineIndex];
  let startCharacter = -1;
  let searchOffset = 0;

  assert.notEqual(line, undefined, `line ${lineIndex} must exist`);

  for (let index = 0; index <= occurrence; index += 1) {
    startCharacter = line.indexOf(identifier, searchOffset);
    searchOffset = startCharacter + identifier.length;
  }

  assert.notEqual(startCharacter, -1, `identifier '${identifier}' must exist on line ${lineIndex}`);

  const token = tokens.find(
    (entry) =>
      entry.range.start.line === lineIndex &&
      entry.range.start.character === startCharacter &&
      entry.range.end.line === lineIndex &&
      entry.range.end.character === startCharacter + identifier.length
  );

  assert.ok(token, `semantic token '${identifier}' must exist at ${lineIndex}:${startCharacter}`);
  assert.equal(token.type, expected.type);
  assert.deepEqual([...token.modifiers].sort(), [...expected.modifiers].sort());
}

function assertNoSemanticToken(text, tokens, lineIndex, identifier, occurrence = 0) {
  const lines = text.split("\n");
  const line = lines[lineIndex];
  let startCharacter = -1;
  let searchOffset = 0;

  assert.notEqual(line, undefined, `line ${lineIndex} must exist`);

  for (let index = 0; index <= occurrence; index += 1) {
    startCharacter = line.indexOf(identifier, searchOffset);
    searchOffset = startCharacter + identifier.length;
  }

  assert.notEqual(startCharacter, -1, `identifier '${identifier}' must exist on line ${lineIndex}`);
  assert.equal(
    tokens.some(
      (entry) =>
        entry.range.start.line === lineIndex &&
        entry.range.start.character === startCharacter &&
        entry.range.end.line === lineIndex &&
        entry.range.end.character === startCharacter + identifier.length
    ),
    false,
    `semantic token '${identifier}' must not exist at ${lineIndex}:${startCharacter}`
  );
}

function createWorksheetBroadRootFixture(text) {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "book1");
  const moduleDirectory = path.join(bundleRoot, "modules");
  const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;

  try {
    mkdirSync(moduleDirectory, { recursive: true });
    writeDefaultWorksheetBroadRootArtifacts(bundleRoot);

    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });
    service.analyzeText(uri, "vba", 1, text);

    return {
      service,
      uri,
      cleanup() {
        rmSync(temporaryDirectory, { force: true, recursive: true });
      }
    };
  } catch (error) {
    rmSync(temporaryDirectory, { force: true, recursive: true });
    throw error;
  }
}

function createWorkbookQualifiedWorksheetRootFixture(text) {
  const temporaryDirectory = mkdtempSync(path.join(os.tmpdir(), "vba-server-sidecar-"));
  const workspaceRoot = path.join(temporaryDirectory, "workspace");
  const bundleRoot = path.join(workspaceRoot, "book1");
  const moduleDirectory = path.join(bundleRoot, "modules");
  const uri = pathToFileURL(path.join(moduleDirectory, "Module1.bas")).href;
  const thisWorkbookUri = pathToFileURL(path.join(bundleRoot, "ThisWorkbook.cls")).href;

  try {
    mkdirSync(moduleDirectory, { recursive: true });
    writeDefaultWorksheetBroadRootArtifacts(bundleRoot);

    const service = createDocumentService({ workspaceRoots: [workspaceRoot] });
    service.analyzeText(
      thisWorkbookUri,
      "vba",
      1,
      `Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_PredeclaredId = True
Option Explicit`
    );
    service.analyzeText(uri, "vba", 1, text);

    return {
      service,
      uri,
      cleanup() {
        rmSync(temporaryDirectory, { force: true, recursive: true });
      }
    };
  } catch (error) {
    rmSync(temporaryDirectory, { force: true, recursive: true });
    throw error;
  }
}

function writeDefaultWorksheetBroadRootArtifacts(bundleRoot) {
  writeWorksheetControlMetadataSidecar(bundleRoot, {
    artifact: "worksheet-control-metadata-sidecar",
    owners: [
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
        sheetName: "Sheet One",
        status: "supported"
      }
    ],
    version: 1,
    workbook: {
      name: "book1.xlsm",
      sourceKind: "openxml-package"
    }
  });
  writeWorkbookBindingManifest(bundleRoot, {
    artifact: "workbook-binding-manifest",
    bindingKind: "active-workbook-fullname",
    version: 1,
    workbook: {
      fullName: "C:\\Fixtures\\book1.xlsm",
      isAddIn: false,
      name: "book1.xlsm",
      path: "C:\\Fixtures",
      sourceKind: "openxml-package"
    }
  });
}

function createMatchedActiveWorkbookIdentitySnapshot() {
  return {
    identity: {
      fullName: "c:/fixtures/BOOK1.xlsm",
      isAddin: false,
      name: "book1.xlsm",
      path: "c:/fixtures"
    },
    observedAt: "2026-03-21T00:00:00.000Z",
    providerKind: "excel-active-workbook",
    state: "available",
    version: 1
  };
}

function createMismatchedActiveWorkbookIdentitySnapshot() {
  return {
    identity: {
      fullName: "C:\\Fixtures\\OtherBook.xlsm",
      isAddin: false,
      name: "OtherBook.xlsm",
      path: "C:\\Fixtures"
    },
    observedAt: "2026-03-21T00:00:30.000Z",
    providerKind: "excel-active-workbook",
    state: "available",
    version: 1
  };
}

function createUnavailableActiveWorkbookIdentitySnapshot() {
  return {
    observedAt: "2026-03-21T00:01:00.000Z",
    providerKind: "excel-active-workbook",
    reason: "no-active-workbook",
    state: "unavailable",
    version: 1
  };
}

function getCompletionSymbolsAfterToken(service, uri, text, token) {
  return service.getCompletionSymbols(uri, findPositionAfterTokenInText(text, token));
}

function hasCompletionSymbolAfterToken(service, uri, text, token, symbolName) {
  return getCompletionSymbolsAfterToken(service, uri, text, token).some(
    (resolution) => resolution.symbol.name === symbolName
  );
}

function getHoverAfterToken(service, uri, text, token) {
  return service.getHover(uri, findPositionAfterTokenInText(text, token));
}

function getSignatureHelpAfterToken(service, uri, text, token) {
  return service.getSignatureHelp(uri, findPositionAfterTokenInText(text, token));
}

function writeWorksheetControlMetadataSidecar(bundleRoot, metadata) {
  const sidecarDirectory = path.join(bundleRoot, ".vba");
  mkdirSync(sidecarDirectory, { recursive: true });
  writeFileSync(path.join(sidecarDirectory, "worksheet-control-metadata.json"), `${JSON.stringify(metadata, null, 2)}\n`);
}

function writeWorkbookBindingManifest(bundleRoot, manifest) {
  const manifestDirectory = path.join(bundleRoot, ".vba");
  mkdirSync(manifestDirectory, { recursive: true });
  writeFileSync(path.join(manifestDirectory, "workbook-binding.json"), `${JSON.stringify(manifest, null, 2)}\n`);
}

function applyTextEdit(text, edit) {
  const normalizedText = text.replace(/\r\n?/g, "\n");
  const startOffset = toOffset(normalizedText, edit.range.start);
  const endOffset = toOffset(normalizedText, edit.range.end);
  return normalizedText.slice(0, startOffset) + edit.newText + normalizedText.slice(endOffset);
}

function findPositionAfterTokenInText(text, token, offsetFromEnd = 0) {
  const normalizedText = text.replace(/\r\n?/g, "\n");
  const startOffset = normalizedText.indexOf(token);

  assert.notEqual(startOffset, -1, `token not found in text: ${token}`);

  return toPosition(normalizedText, startOffset + token.length + offsetFromEnd);
}

function toOffset(text, position) {
  const lines = text.split("\n");
  let offset = 0;

  for (let index = 0; index < position.line; index += 1) {
    offset += (lines[index]?.length ?? 0) + 1;
  }

  return offset + position.character;
}

function toPosition(text, offset) {
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
