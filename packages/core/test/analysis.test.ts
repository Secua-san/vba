import assert from "node:assert/strict";
import test from "node:test";
import {
  analyzeModule,
  findDefinition,
  getCompletionSymbols,
  getDocumentOutline,
  getSymbolTypeName,
  lexDocument,
  parseModule
} from "../dist/index.js";

test("lexDocument tokenizes VBA literals, comments, and attributes", () => {
  const tokens = lexDocument(
    `Attribute VB_Name = "Module1"
Option Explicit
Dim value$ As String
message = "a""b"
stamp = #2024-01-01#
Rem comment`
  );

  assert.ok(tokens.some((token) => token.kind === "attribute"));
  assert.ok(tokens.some((token) => token.kind === "stringLiteral" && token.text === "\"a\"\"b\""));
  assert.ok(tokens.some((token) => token.kind === "dateLiteral"));
  assert.ok(tokens.some((token) => token.kind === "comment"));
});

test("parseModule recovers from broken procedure blocks", () => {
  const result = parseModule(`Option Explicit

Public Sub Broken()
    If True Then
        value = 1
End Sub`, { fileName: "Broken.bas" });

  assert.ok(result.module.members.some((member) => member.kind === "procedureDeclaration"));
  assert.ok(result.diagnostics.some((diagnostic) => diagnostic.code === "syntax-error"));
});

test("analyzeModule reports undeclared identifiers and missing PtrSafe", () => {
  const result = analyzeModule(`Option Explicit

Private Declare Function MessageBoxA Lib "user32" (ByVal hwnd As LongPtr) As Long

Public Sub Demo()
    Dim knownValue As Long
    knownValue = missingValue + 1
End Sub`, { fileName: "Demo.bas" });

  assert.ok(result.diagnostics.some((diagnostic) => diagnostic.code === "declare-missing-ptrsafe"));
  assert.ok(result.diagnostics.some((diagnostic) => diagnostic.code === "undeclared-variable"));
});

test("analyzeModule skips frm designer text and exposes navigation symbols", () => {
  const result = analyzeModule(`VERSION 5.00
Begin VB.Form UserForm1
    Caption = "Sample"
End
Attribute VB_Name = "UserForm1"
Option Explicit

Public Sub ShowMessage()
    Dim message As String
    message = "Hello"
    MsgBox message
End Sub`, { fileName: "UserForm1.frm" });

  const completions = getCompletionSymbols(result, { character: 4, line: 8 });
  const definition = findDefinition(result, { character: 11, line: 8 });
  const outline = getDocumentOutline(result);

  assert.equal(result.module.name, "UserForm1");
  assert.ok(completions.some((symbol) => symbol.name === "message"));
  assert.equal(definition?.name, "message");
  assert.ok(outline[0]?.children?.some((symbol) => symbol.name === "ShowMessage"));
});

test("findDefinition prefers the declaration under the cursor when names are shadowed", () => {
  const result = analyzeModule(`Attribute VB_Name = "Shadowing"
Option Explicit

Public Const SharedValue As Long = 1

Public Sub Demo()
    Dim SharedValue As Long
    SharedValue = 2
End Sub`, { fileName: "Shadowing.bas" });

  const declarationDefinition = findDefinition(result, { character: 8, line: 6 });
  const usageDefinition = findDefinition(result, { character: 8, line: 7 });

  assert.equal(declarationDefinition?.kind, "variable");
  assert.equal(declarationDefinition?.scope, "procedure");
  assert.equal(usageDefinition?.kind, "variable");
  assert.equal(usageDefinition?.scope, "procedure");
});

test("analyzeModule infers types from literals, simple assignments, and function return values", () => {
  const result = analyzeModule(`Attribute VB_Name = "Inference"
Option Explicit

Public Function BuildMessage()
    BuildMessage = "Hello"
End Function

Public Sub Demo()
    Dim message
    Dim copiedMessage As String
    message = BuildMessage()
    copiedMessage = message
End Sub`, { fileName: "Inference.bas" });

  const procedureSymbol = result.symbols.moduleSymbols.find((symbol) => symbol.kind === "procedure" && symbol.name === "BuildMessage");
  const messageSymbol = result.symbols.procedureScopes
    .flatMap((scope) => scope.symbols)
    .find((symbol) => symbol.kind === "variable" && symbol.name === "message");

  assert.equal(getSymbolTypeName(result, procedureSymbol), "String");
  assert.equal(getSymbolTypeName(result, messageSymbol), "String");
  assert.equal(result.diagnostics.some((diagnostic) => diagnostic.code === "type-mismatch"), false);
});

test("analyzeModule reports simple type mismatches", () => {
  const result = analyzeModule(`Attribute VB_Name = "Mismatch"
Option Explicit

Public Function CountItems() As Long
    CountItems = "wrong"
End Function

Public Sub Demo()
    Dim title As String
    title = True
End Sub`, { fileName: "Mismatch.bas" });

  const mismatchDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "type-mismatch");

  assert.equal(mismatchDiagnostics.length, 2);
  assert.ok(mismatchDiagnostics.every((diagnostic) => diagnostic.severity === "warning"));
});

test("analyzeModule expands type mismatch diagnostics for compound expressions and Set assignments", () => {
  const result = analyzeModule(`Attribute VB_Name = "ExpandedMismatch"
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
End Sub`, { fileName: "ExpandedMismatch.bas" });

  const mismatchDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "type-mismatch");

  assert.equal(mismatchDiagnostics.length, 3);
  assert.deepEqual(
    mismatchDiagnostics.map((diagnostic) => diagnostic.message),
    [
      "Type mismatch: cannot assign Long to String.",
      "Type mismatch: cannot assign String to Long.",
      "Type mismatch: cannot assign Long to Object."
    ]
  );
});

test("analyzeModule warns on risky ByRef arguments while allowing safe array elements", () => {
  const result = analyzeModule(`Attribute VB_Name = "ByRefRisks"
Option Explicit

Private Sub UpdateCount(ByRef count As Long)
End Sub

Private Sub UpdateLabel(ByVal label As String)
End Sub

Public Sub Demo()
    Dim count As Long
    Dim wrongCount As String
    Dim values() As Long
    UpdateCount count + 1
    UpdateCount wrongCount
    UpdateCount values(0)
    UpdateLabel wrongCount & "!"
End Sub`, { fileName: "ByRefRisks.bas" });

  const byRefDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code.startsWith("byref-"));

  assert.equal(byRefDiagnostics.length, 2);
  assert.deepEqual(
    byRefDiagnostics.map((diagnostic) => diagnostic.message),
    [
      "ByRef parameter 'count' in UpdateCount receives an expression. Introduce a temporary variable before the call.",
      "ByRef parameter 'count' in UpdateCount expects Long but receives String. VBA may raise a ByRef argument type mismatch."
    ]
  );
});

test("analyzeModule warns when object assignments omit Set", () => {
  const result = analyzeModule(`Attribute VB_Name = "SetRequired"
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
End Sub`, { fileName: "SetRequired.bas" });

  const setDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "set-required");
  const typeMismatchDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "type-mismatch");

  assert.equal(setDiagnostics.length, 3);
  assert.deepEqual(
    setDiagnostics.map((diagnostic) => diagnostic.message),
    [
      "Set is required to assign Collection to Collection.",
      "Set is required to assign Collection to Object.",
      "Set is required to assign Nothing to Collection."
    ]
  );
  assert.equal(typeMismatchDiagnostics.length, 0);
});
