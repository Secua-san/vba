import assert from "node:assert/strict";
import test from "node:test";
import {
  analyzeModule,
  findDefinition,
  getCompletionSymbols,
  getDocumentOutline,
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
