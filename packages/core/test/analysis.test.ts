import assert from "node:assert/strict";
import test from "node:test";
import {
  analyzeModule,
  findDefinition,
  formatModuleIndentation,
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

test("analyzeModule reports type mismatches across continued assignments", () => {
  const result = analyzeModule(`Attribute VB_Name = "ContinuedMismatch"
Option Explicit

Public Sub Demo()
    Dim title As String
    title = _
        True
End Sub`, { fileName: "ContinuedMismatch.bas" });

  const mismatchDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "type-mismatch");

  assert.equal(mismatchDiagnostics.length, 1);
  assert.equal(mismatchDiagnostics[0]?.message, "Type mismatch: cannot assign Boolean to String.");
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

test("analyzeModule does not require Set for user-defined types", () => {
  const result = analyzeModule(`Attribute VB_Name = "UserTypes"
Option Explicit

Private Type ItemInfo
    Value As Long
End Type

Public Sub Demo()
    Dim leftValue As ItemInfo
    Dim rightValue As ItemInfo
    leftValue = rightValue
End Sub`, { fileName: "UserTypes.bas" });

  assert.equal(result.diagnostics.some((diagnostic) => diagnostic.code === "set-required"), false);
  assert.equal(result.diagnostics.some((diagnostic) => diagnostic.code === "type-mismatch"), false);
});

test("analyzeModule reports duplicate definitions in module and procedure scopes", () => {
  const result = analyzeModule(`Attribute VB_Name = "Duplicates"
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
End Sub`, { fileName: "Duplicates.bas" });

  const duplicateDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "duplicate-definition");

  assert.equal(duplicateDiagnostics.length, 5);
  assert.deepEqual(
    duplicateDiagnostics.map((diagnostic) => diagnostic.message),
    [
      "Duplicate definition 'CustomerRecord' in module scope.",
      "Duplicate definition 'StatusKind' in module scope.",
      "Duplicate definition 'SharedName' in module scope.",
      "Duplicate definition 'value' in procedure 'Demo'.",
      "Duplicate definition 'title' in procedure 'Demo'."
    ]
  );
  assert.ok(duplicateDiagnostics.every((diagnostic) => diagnostic.severity === "error"));
});

test("analyzeModule warns on unreachable code after unconditional exits", () => {
  const result = analyzeModule(`Attribute VB_Name = "Unreachable"
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
End Sub`, { fileName: "Unreachable.bas" });

  const unreachableDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unreachable-code");

  assert.equal(unreachableDiagnostics.length, 3);
  assert.deepEqual(
    unreachableDiagnostics.map((diagnostic) => diagnostic.message),
    [
      "Unreachable code after Exit Sub.",
      "Unreachable code after Exit Sub.",
      "Unreachable code after End."
    ]
  );
  assert.ok(unreachableDiagnostics.every((diagnostic) => diagnostic.severity === "warning"));
});

test("analyzeModule warns on unused local variables and parameters", () => {
  const result = analyzeModule(`Attribute VB_Name = "UnusedLocals"
Option Explicit

Public Sub Demo(ByVal usedArg As Long, ByVal unusedArg As Long)
    Dim usedValue As Long
    Dim writeOnlyValue As Long
    Dim unusedValue As String
    usedValue = usedArg
    writeOnlyValue = 1
    Debug.Print usedValue
End Sub`, { fileName: "UnusedLocals.bas" });

  const unusedDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unused-variable");

  assert.equal(unusedDiagnostics.length, 2);
  assert.deepEqual(
    unusedDiagnostics.map((diagnostic) => diagnostic.message),
    [
      "Unused local declaration 'unusedArg'.",
      "Unused local declaration 'unusedValue'."
    ]
  );
  assert.ok(unusedDiagnostics.every((diagnostic) => diagnostic.severity === "warning"));
});

test("analyzeModule warns on write-only local variables without duplicating unused-variable", () => {
  const result = analyzeModule(`Attribute VB_Name = "WriteOnlyLocals"
Option Explicit

Public Sub Demo()
    Dim readValue As Long
    Dim writeOnlyValue As Long
    Dim objectHolder As Collection
    readValue = 1
    writeOnlyValue = readValue
    Set objectHolder = New Collection
    Debug.Print readValue
End Sub`, { fileName: "WriteOnlyLocals.bas" });

  const writeOnlyDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "write-only-variable");
  const unusedDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unused-variable");

  assert.equal(writeOnlyDiagnostics.length, 2);
  assert.deepEqual(
    writeOnlyDiagnostics.map((diagnostic) => diagnostic.message),
    [
      "Write-only local variable 'writeOnlyValue'.",
      "Write-only local variable 'objectHolder'."
    ]
  );
  assert.equal(unusedDiagnostics.length, 0);
  assert.ok(writeOnlyDiagnostics.every((diagnostic) => diagnostic.severity === "warning"));
});

test("formatModuleIndentation indents nested VBA blocks and continued lines", () => {
  const formatted = formatModuleIndentation(`Attribute VB_Name = "Formatter"
Option Explicit

Public Property Get Message() As String
Dim value As String
If True Then
value = _
"Hello"
Else
Select Case Len(value)
Case 0
With Application
.StatusBar = value
End With
Case Else
For index = 1 To 2
Do While index < 2
index = index + 1
Loop
Next index
End Select
End If
Message = value
End Property`, { fileName: "Formatter.bas", indentSize: 4, insertSpaces: true });

  assert.equal(
    formatted,
    `Attribute VB_Name = "Formatter"
Option Explicit

Public Property Get Message() As String
    Dim value As String
    If True Then
        value = _
            "Hello"
    Else
        Select Case Len(value)
            Case 0
                With Application
                    .StatusBar = value
                End With
            Case Else
                For index = 1 To 2
                    Do While index < 2
                        index = index + 1
                    Loop
                Next index
        End Select
    End If
    Message = value
End Property`
  );
});

test("formatModuleIndentation preserves frm designer text and outdents labels", () => {
  const formatted = formatModuleIndentation(`VERSION 5.00
Begin VB.Form SampleForm
    Caption = "Sample"
End
Attribute VB_Name = "SampleForm"
Option Explicit

Public Sub Demo()
If True Then
GoHere:
Debug.Print "ready"
End If
End Sub`, { fileName: "SampleForm.frm", indentSize: 4, insertSpaces: true });

  assert.equal(
    formatted,
    `VERSION 5.00
Begin VB.Form SampleForm
    Caption = "Sample"
End
Attribute VB_Name = "SampleForm"
Option Explicit

Public Sub Demo()
    If True Then
GoHere:
        Debug.Print "ready"
    End If
End Sub`
  );
});

test("formatModuleIndentation stabilizes hanging indent for assignments, arguments, and method chains", () => {
  const formatted = formatModuleIndentation(`Attribute VB_Name = "ContinuationFormatting"
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
End Sub`, { fileName: "ContinuationFormatting.bas", indentSize: 4, insertSpaces: true });

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

test("formatModuleIndentation expands compressed block boundaries while keeping ordinary colon statements", () => {
  const formatted = formatModuleIndentation(`Attribute VB_Name = "BlockLayoutFormatting"
Option Explicit

Public Sub Demo()
Dim value As Long: value = 0
If value = 0 Then: Debug.Print "zero": ElseIf value = 1 Then: Debug.Print "one": Else: Debug.Print "other": End If
Select Case value: Case 0: Debug.Print "case zero": Case Else: With Application: .StatusBar = "fallback": End With: End Select
#If VBA7 Then: value = value + 1: #Else: value = value - 1: #End If
End Sub`, { fileName: "BlockLayoutFormatting.bas", indentSize: 4, insertSpaces: true });

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

test("formatModuleIndentation aligns declaration blocks conservatively", () => {
  const formatted = formatModuleIndentation(`Attribute VB_Name = "DeclarationAlignment"
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
End Sub`, { fileName: "DeclarationAlignment.bas", indentSize: 4, insertSpaces: true });

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

test("formatModuleIndentation normalizes comment spacing while preserving comment placement", () => {
  const formatted = formatModuleIndentation(`Attribute VB_Name = "CommentFormatting"
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
End Sub`, { fileName: "CommentFormatting.bas", indentSize: 4, insertSpaces: true });

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
