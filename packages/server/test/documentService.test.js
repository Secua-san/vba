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
  assert.ok(service.getCompletionSymbols(state.uri, { character: 4, line: 5 }).some((resolution) => resolution.symbol.name === "message"));
  assert.equal(service.getDefinition(state.uri, { character: 5, line: 5 })?.symbol.name, "message");
  assert.ok(service.getDocumentSymbols(state.uri)[0]?.children?.some((symbol) => symbol.name === "Demo"));
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

  assert.ok(completions.some((resolution) => resolution.uri === libraryUri && resolution.symbol.name === "PublicMessage"));
  assert.equal(definition?.uri, libraryUri);
  assert.equal(definition?.moduleName, "PublicApi");
  assert.equal(definition?.symbol.name, "PublicMessage");
  assert.equal(diagnostics.some((diagnostic) => diagnostic.code === "undeclared-variable"), false);
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
});
