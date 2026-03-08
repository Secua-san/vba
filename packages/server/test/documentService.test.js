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
