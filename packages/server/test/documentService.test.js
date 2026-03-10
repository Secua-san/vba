const assert = require("node:assert/strict");
const test = require("node:test");
const { createDocumentService } = require("../dist/index.js");

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

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "BuiltInMemberCompletion"
Option Explicit

Public Sub Demo()
    Debug.Print Application.
    Debug.Print WorksheetFunction.Su
    Debug.Print Application.WorksheetFunction.Su
End Sub`
  );

  const applicationMembers = service.getCompletionSymbols(uri, { character: 28, line: 4 });
  const worksheetFunctionMembers = service.getCompletionSymbols(uri, { character: 36, line: 5 });
  const chainedWorksheetFunctionMembers = service.getCompletionSymbols(uri, { character: 48, line: 6 });

  const worksheetFunctionProperty = applicationMembers.find((resolution) => resolution.symbol.name === "WorksheetFunction");
  const activeCellProperty = applicationMembers.find((resolution) => resolution.symbol.name === "ActiveCell");
  const worksheetFunctionSum = worksheetFunctionMembers.find((resolution) => resolution.symbol.name === "Sum");
  const chainedWorksheetFunctionSum = chainedWorksheetFunctionMembers.find((resolution) => resolution.symbol.name === "Sum");

  assert.equal(worksheetFunctionProperty?.isBuiltIn, true);
  assert.equal(worksheetFunctionProperty?.moduleName, "Excel Application property");
  assert.equal(worksheetFunctionProperty?.typeName, "WorksheetFunction");
  assert.equal(worksheetFunctionProperty?.documentation?.includes("excel.application.worksheetfunction"), true);
  assert.equal(activeCellProperty?.moduleName, "Excel Application property");
  assert.equal(worksheetFunctionSum?.moduleName, "Excel WorksheetFunction method");
  assert.equal(worksheetFunctionSum?.documentation?.includes("excel.worksheetfunction.sum"), true);
  assert.equal(chainedWorksheetFunctionSum?.moduleName, "Excel WorksheetFunction method");
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
  const text = `Attribute VB_Name = "BuiltInSignature"
Option Explicit

Public Sub Demo()
    Dim transposedResult As Variant
    Debug.Print WorksheetFunction.Sum(1, 2)
    Debug.Print Application.WorksheetFunction.Sum(1, 2)
    Debug.Print Application.WorksheetFunction.Power(2, 3)
    Debug.Print WorksheetFunction.Average(1, 2, 3)
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
    Call Application.CalculateFull()
    Application.OnTime(Now, "BuiltInSignature.Demo")
    Call Application.WorksheetFunction()
    Call Application.AfterCalculate()
    Call Application.ActiveCell()
    Call Application.NewWorkbook()
    Debug.Print Application.Calculate
End Sub`;

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

  service.analyzeText(
    uri,
    "vba",
    1,
    `Attribute VB_Name = "BuiltInSignatureShadowed"
Option Explicit

Public Sub Demo()
    Dim WorksheetFunction As String
    Debug.Print WorksheetFunction.Sum(1, 2)
End Sub`
  );

  assert.equal(service.getSignatureHelp(uri, { character: 38, line: 5 }), undefined);
  assert.equal(service.getHover(uri, { character: 35, line: 5 }), undefined);
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
  const text = `Attribute VB_Name = "BuiltInSemantic"
Option Explicit

Public Sub Demo()
    Beep
    MsgBox xlAll
    Debug.Print Application.Name
    Debug.Print Application.WorksheetFunction.Sum(1, 2)
End Sub`;

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
