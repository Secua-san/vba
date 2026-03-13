import assert from "node:assert/strict";
import path from "node:path";
import * as vscode from "vscode";

export async function run(): Promise<void> {
  const extension = vscode.extensions.getExtension("tagi0.vba-extension");
  assert.ok(extension, "extension must be discoverable");

  await extension.activate();
  await vscode.workspace.getConfiguration("editor").update("snippetSuggestions", "top", vscode.ConfigurationTarget.Global);
  await vscode.workspace.getConfiguration("editor").update("insertSpaces", true, vscode.ConfigurationTarget.Global);
  await vscode.workspace.getConfiguration("editor").update("tabSize", 4, vscode.ConfigurationTarget.Global);

  const fixturesPath = path.resolve(__dirname, "..", "..", "test", "fixtures");
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
  const namedDialogSheetCompletionItems = await waitForCompletions(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, 'DialogSheets("Dialog1").'),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Activate")
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
  const dialogSheetEvaluateSignatureHelp = await waitForSignatureHelp(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1).Evaluate("),
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
  const dialogSheetHover = await waitForHover(
    dialogSheetBuiltInDocument,
    findPositionAfterToken(dialogSheetBuiltInDocument, "DialogSheets(1).SaveA"),
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
  const dialogSheetSaveAsCompletion = dialogSheetItemCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SaveAs"
  );
  const dialogSheetEvaluateCompletion = dialogSheetItemCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Evaluate"
  );
  const namedDialogSheetActivateCompletion = namedDialogSheetCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Activate"
  );
  const groupedDialogSheetsCountCompletion = groupedDialogSheetsCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Count"
  );
  const itemDialogSheetSaveAsCompletion = itemDialogSheetCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "SaveAs"
  );
  const dialogSheetHoverText = getHoverContentsText(dialogSheetHover[0]);

  assert.ok(dialogSheetsCountCompletion?.detail?.includes("Excel DialogSheets property"));
  assert.ok(dialogSheetSaveAsCompletion?.detail?.includes("Excel DialogSheet method"));
  assert.ok(dialogSheetEvaluateCompletion?.detail?.includes("Excel DialogSheet method"));
  assert.ok(namedDialogSheetActivateCompletion?.detail?.includes("Excel DialogSheet method"));
  assert.ok(groupedDialogSheetsCountCompletion?.detail?.includes("Excel DialogSheets property"));
  assert.ok(itemDialogSheetSaveAsCompletion?.detail?.includes("Excel DialogSheet method"));
  assert.equal(
    groupedDialogSheetsCompletionItems.some((item) => getCompletionItemLabel(item) === "SaveAs"),
    false,
    "grouped DialogSheets selector should stay on the DialogSheets collection"
  );
  assert.equal(dialogSheetEvaluateSignatureHelp.signatures[0]?.label, "Evaluate(Name) As Object");
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
  assert.equal(
    dialogSheetExportSignatureHelp.signatures[0]?.label,
    "ExportAsFixedFormat(Type, Filename, Quality, ..., FixedFormatExtClassPtr)"
  );
  assert.equal(groupedDialogSheetSaveAsSuppressed, true);
  assert.ok(dialogSheetHoverText.includes("SaveAs(Filename, FileFormat, Password, ..., Local)"));
  assert.ok(dialogSheetHoverText.includes("microsoft.office.interop.excel.dialogsheet.saveas"));
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
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const signatureHelp = await vscode.commands.executeCommand<vscode.SignatureHelp>(
      "vscode.executeSignatureHelpProvider",
      document.uri,
      position
    );

    if (signatureHelp) {
      return false;
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  return true;
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

function findPositionAfterToken(document: vscode.TextDocument, token: string, offsetFromEnd = 0): vscode.Position {
  const source = document.getText();
  const startIndex = source.indexOf(token);
  assert.notEqual(startIndex, -1, `token not found in document: ${token}`);
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
  expected: { modifiers: string[]; type: string },
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

function getCodeAction(actions: readonly vscode.CodeAction[], title: string): vscode.CodeAction | undefined {
  return actions.find((action) => action.title === title);
}

function hasCodeAction(actions: readonly vscode.CodeAction[], title: string): boolean {
  return getCodeAction(actions, title) !== undefined;
}

function isCodeAction(action: vscode.CodeAction | vscode.Command): action is vscode.CodeAction {
  return "title" in action && "edit" in action;
}
