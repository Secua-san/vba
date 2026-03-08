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
  assert.ok(service.getCompletionSymbols(state.uri, { character: 4, line: 5 }).some((symbol) => symbol.name === "message"));
  assert.equal(service.getDefinition(state.uri, { character: 5, line: 5 })?.name, "message");
  assert.ok(service.getDocumentSymbols(state.uri)[0]?.children?.some((symbol) => symbol.name === "Demo"));
});
