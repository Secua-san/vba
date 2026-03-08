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
