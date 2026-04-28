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
total = 1# + _
    value%
Rem comment`
  );

  assert.ok(tokens.some((token) => token.kind === "attribute"));
  assert.ok(tokens.some((token) => token.kind === "lineContinuation" && token.text === "_"));
  assert.ok(tokens.some((token) => token.kind === "stringLiteral" && token.text === "\"a\"\"b\""));
  assert.ok(tokens.some((token) => token.kind === "dateLiteral"));
  assert.ok(tokens.some((token) => token.kind === "comment"));
  assert.ok(tokens.some((token) => token.kind === "numberLiteral" && token.text === "1#"));
  assert.ok(tokens.some((token) => token.kind === "identifier" && token.text === "value%"));

  const continuationBeforeCommentTokens = lexDocument(`value = 1 _ Rem note
value = 1 _ ' note`);
  assert.equal(
    continuationBeforeCommentTokens.filter((token) => token.kind === "lineContinuation").length,
    2
  );

  const nonContinuationTokens = lexDocument(`value = 1 _remote
value = _Rem note`);
  assert.equal(nonContinuationTokens.some((token) => token.kind === "lineContinuation"), false);
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

test("parseModule reports invalid ElseIf and Case ordering inside structured blocks", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    If ready Then
    Else
    ElseIf fallback Then
    End If

    Select Case value
        Case Else
        Case 1
    End Select
End Sub`, { fileName: "InvalidClauseOrder.bas" });

  assert.deepEqual(
    result.diagnostics
      .filter((diagnostic) => diagnostic.code === "syntax-error")
      .map((diagnostic) => ({
        message: diagnostic.message,
        start: `${diagnostic.range.start.line}:${diagnostic.range.start.character}`
      })),
    [
      {
        message: "Unexpected block clause in Demo.",
        start: "5:0"
      },
      {
        message: "Unexpected block clause in Demo.",
        start: "10:0"
      }
    ]
  );
});

test("parseModule reports mismatched Next counters for structured For blocks", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    Dim items As Collection

    For index = 1 To 2
    Next otherIndex

    For Each item In items
    Next otherItem
End Sub`, { fileName: "InvalidNextCounters.bas" });

  assert.deepEqual(
    result.diagnostics
      .filter((diagnostic) => diagnostic.code === "syntax-error")
      .map((diagnostic) => ({
        message: diagnostic.message,
        start: `${diagnostic.range.start.line}:${diagnostic.range.start.character}`
      })),
    [
      {
        message: "Next counter 'otherIndex' does not match active loop variable 'index' in Demo.",
        start: "6:0"
      },
      {
        message: "Next counter 'otherItem' does not match active loop variable 'item' in Demo.",
        start: "9:0"
      }
    ]
  );
});

test("parseModule keeps labeled procedure declarations structured for Const and Dim statements", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
Label1: Const localValue As Long = 1
Label2: Dim totalCount As Long
End Sub`, { fileName: "LabeledDeclarations.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const constStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[0] : undefined;
  const declarationStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[1] : undefined;

  assert.ok(procedure && procedure.kind === "procedureDeclaration");
  assert.equal(constStatement?.kind, "constStatement");
  assert.deepEqual(
    constStatement?.declaredConstants.map((constant) => constant.name),
    ["localValue"]
  );
  assert.deepEqual(
    constStatement?.declaredConstants.map((constant) => constant.valueText),
    ["1"]
  );
  assert.equal(declarationStatement?.kind, "declarationStatement");
  assert.deepEqual(
    declarationStatement?.declaredVariables.map((variable) => variable.name),
    ["totalCount"]
  );
});

test("parseModule preserves module-level multiline Const value ranges", () => {
  const result = parseModule(`Option Explicit
Private Const moduleValue As Long = _
    1`, { fileName: "ModuleConst.bas" });
  const constant = result.module.members.find((member) => member.kind === "constDeclaration");

  assert.ok(constant && constant.kind === "constDeclaration");
  assert.equal(constant.valueText, "1");
  assert.deepEqual(
    {
      end: `${constant.valueRange?.end.line}:${constant.valueRange?.end.character}`,
      start: `${constant.valueRange?.start.line}:${constant.valueRange?.start.character}`
    },
    {
      end: "2:5",
      start: "2:4"
    }
  );
});

test("parseModule structures simple assignment and call statements in procedure bodies", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    Set holder = CreateObject("Scripting.Dictionary")
    Call UpdateCount(holder.Count)
    UpdateCount holder.Count
    Debug.Print holder.Count
    Property.Get holder.Count
End Sub`, { fileName: "StructuredStatements.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const assignmentStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[0] : undefined;
  const callKeywordStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[1] : undefined;
  const bareCallStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[2] : undefined;
  const memberCallStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[3] : undefined;
  const propertyCallStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[4] : undefined;

  assert.ok(procedure && procedure.kind === "procedureDeclaration");
  assert.ok(assignmentStatement?.kind === "assignmentStatement");
  assert.equal(assignmentStatement.assignmentKind, "set");
  assert.equal(assignmentStatement.targetText, "holder");
  assert.equal(assignmentStatement.expressionText, 'CreateObject("Scripting.Dictionary")');
  assert.ok(callKeywordStatement?.kind === "callStatement");
  assert.equal(callKeywordStatement.callStyle, "call");
  assert.equal(callKeywordStatement.name, "UpdateCount");
  assert.deepEqual(callKeywordStatement.arguments.map((argument) => argument.text), ["holder.Count"]);
  assert.ok(bareCallStatement?.kind === "callStatement");
  assert.equal(bareCallStatement.callStyle, "bare");
  assert.equal(bareCallStatement.name, "UpdateCount");
  assert.ok(memberCallStatement?.kind === "callStatement");
  assert.equal(memberCallStatement.callStyle, "bare");
  assert.equal(memberCallStatement.name, "Debug.Print");
  assert.deepEqual(memberCallStatement.arguments.map((argument) => argument.text), ["holder.Count"]);
  assert.ok(propertyCallStatement?.kind === "callStatement");
  assert.equal(propertyCallStatement.callStyle, "bare");
  assert.equal(propertyCallStatement.name, "Property.Get");
  assert.deepEqual(propertyCallStatement.arguments.map((argument) => argument.text), ["holder.Count"]);
});

test("parseModule structures call statements that span multiple physical lines", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    Call UpdateCount( _
        holder.Count _
    )
    UpdateCount _
        holder.Count
End Sub`, { fileName: "StructuredMultilineCalls.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const callKeywordStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[0] : undefined;
  const bareCallStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[1] : undefined;

  assert.ok(procedure && procedure.kind === "procedureDeclaration");
  assert.ok(callKeywordStatement?.kind === "callStatement");
  assert.equal(callKeywordStatement.callStyle, "call");
  assert.equal(callKeywordStatement.name, "UpdateCount");
  assert.deepEqual(callKeywordStatement.nameRange, {
    end: { character: 20, line: 3 },
    start: { character: 9, line: 3 }
  });
  assert.deepEqual(callKeywordStatement.arguments.map((argument) => argument.text), ["holder.Count"]);
  assert.deepEqual(callKeywordStatement.arguments[0]?.range, {
    end: { character: 20, line: 4 },
    start: { character: 8, line: 4 }
  });
  assert.ok(bareCallStatement?.kind === "callStatement");
  assert.equal(bareCallStatement.callStyle, "bare");
  assert.equal(bareCallStatement.name, "UpdateCount");
  assert.deepEqual(bareCallStatement.nameRange, {
    end: { character: 15, line: 6 },
    start: { character: 4, line: 6 }
  });
  assert.deepEqual(bareCallStatement.arguments.map((argument) => argument.text), ["holder.Count"]);
  assert.deepEqual(bareCallStatement.arguments[0]?.range, {
    end: { character: 20, line: 7 },
    start: { character: 8, line: 7 }
  });
});

test("parseModule preserves physical-line ranges for multiline assignment statements", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    title = _
        True
End Sub`, { fileName: "StructuredMultilineAssignment.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const assignmentStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[0] : undefined;

  assert.ok(procedure && procedure.kind === "procedureDeclaration");
  assert.equal(assignmentStatement?.kind, "assignmentStatement");
  assert.equal(assignmentStatement?.targetText, "title");
  assert.equal(assignmentStatement?.expressionText, "True");
  assert.deepEqual(assignmentStatement?.targetRange, {
    end: { character: 9, line: 3 },
    start: { character: 4, line: 3 }
  });
  assert.deepEqual(assignmentStatement?.expressionRange, {
    end: { character: 12, line: 4 },
    start: { character: 8, line: 4 }
  });
});

test("parseModule structures labeled assignment and call statements in procedure bodies", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
Label1: title = 1
Label2: UpdateCount title
End Sub`, { fileName: "StructuredLabeledStatements.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const assignmentStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[0] : undefined;
  const callStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[1] : undefined;

  assert.ok(procedure && procedure.kind === "procedureDeclaration");
  assert.equal(assignmentStatement?.kind, "assignmentStatement");
  assert.equal(assignmentStatement?.text, "Label1: title = 1");
  assert.equal(assignmentStatement?.leadingLabel?.text, "Label1");
  assert.deepEqual(assignmentStatement?.leadingLabel?.range, {
    end: { character: 6, line: 3 },
    start: { character: 0, line: 3 }
  });
  assert.equal(assignmentStatement?.targetText, "title");
  assert.deepEqual(assignmentStatement?.targetRange, {
    end: { character: 13, line: 3 },
    start: { character: 8, line: 3 }
  });
  assert.equal(callStatement?.kind, "callStatement");
  assert.equal(callStatement?.text, "Label2: UpdateCount title");
  assert.equal(callStatement?.leadingLabel?.text, "Label2");
  assert.equal(callStatement?.name, "UpdateCount");
  assert.deepEqual(callStatement?.nameRange, {
    end: { character: 19, line: 4 },
    start: { character: 8, line: 4 }
  });
});

test("parseModule preserves leading labels on structured block boundaries", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
LabelIf: If ready Then
LabelElse: Else
LabelEnd: End If
End Sub`, { fileName: "StructuredBoundaryLabels.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const ifStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[0] : undefined;
  const elseStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[1] : undefined;
  const endIfStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[2] : undefined;

  assert.equal(ifStatement?.kind, "ifBlockStatement");
  assert.equal(ifStatement?.leadingLabel?.text, "LabelIf");
  assert.equal(elseStatement?.kind, "elseClauseStatement");
  assert.equal(elseStatement?.leadingLabel?.text, "LabelElse");
  assert.equal(endIfStatement?.kind, "endIfStatement");
  assert.equal(endIfStatement?.leadingLabel?.text, "LabelEnd");
});

test("parseModule structures block If, Select Case, For, and For Each statements in procedure bodies", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    If ready Then
    ElseIf fallback Then
    Else
    End If
    Select Case value
        Case 0
        Case Else
    End Select
    For index = 1 To limit Step 2
    Next index
    For Each item In items
    Next item
End Sub`, { fileName: "StructuredBlocks.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");

  assert.ok(procedure && procedure.kind === "procedureDeclaration");
  const ifStatement = procedure.body[0];
  const elseIfStatement = procedure.body[1];
  const elseStatement = procedure.body[2];
  const endIfStatement = procedure.body[3];
  const selectStatement = procedure.body[4];
  const caseValueStatement = procedure.body[5];
  const caseElseStatement = procedure.body[6];
  const endSelectStatement = procedure.body[7];
  const forStatement = procedure.body[8];
  const nextIndexStatement = procedure.body[9];
  const forEachStatement = procedure.body[10];
  const nextItemStatement = procedure.body[11];

  assert.equal(ifStatement?.kind, "ifBlockStatement");
  assert.equal(ifStatement && ifStatement.kind === "ifBlockStatement" ? ifStatement.conditionText : "", "ready");
  assert.equal(elseIfStatement?.kind, "elseIfClauseStatement");
  assert.equal(
    elseIfStatement && elseIfStatement.kind === "elseIfClauseStatement" ? elseIfStatement.conditionText : "",
    "fallback"
  );
  assert.equal(elseStatement?.kind, "elseClauseStatement");
  assert.equal(endIfStatement?.kind, "endIfStatement");
  assert.equal(selectStatement?.kind, "selectCaseStatement");
  assert.equal(selectStatement && selectStatement.kind === "selectCaseStatement" ? selectStatement.expressionText : "", "value");
  assert.equal(caseValueStatement?.kind, "caseClauseStatement");
  assert.equal(
    caseValueStatement && caseValueStatement.kind === "caseClauseStatement" ? caseValueStatement.conditionText : "",
    "0"
  );
  assert.equal(caseElseStatement?.kind, "caseClauseStatement");
  assert.equal(caseElseStatement && caseElseStatement.kind === "caseClauseStatement" ? caseElseStatement.caseKind : "", "else");
  assert.equal(endSelectStatement?.kind, "endSelectStatement");
  assert.equal(forStatement?.kind, "forStatement");
  assert.equal(forStatement && forStatement.kind === "forStatement" ? forStatement.counterText : "", "index");
  assert.equal(forStatement && forStatement.kind === "forStatement" ? forStatement.startExpressionText : "", "1");
  assert.equal(forStatement && forStatement.kind === "forStatement" ? forStatement.endExpressionText : "", "limit");
  assert.equal(forStatement && forStatement.kind === "forStatement" ? forStatement.stepExpressionText : "", "2");
  assert.equal(nextIndexStatement?.kind, "nextStatement");
  assert.equal(nextIndexStatement && nextIndexStatement.kind === "nextStatement" ? nextIndexStatement.counterText : "", "index");
  assert.equal(forEachStatement?.kind, "forEachStatement");
  assert.equal(forEachStatement && forEachStatement.kind === "forEachStatement" ? forEachStatement.itemText : "", "item");
  assert.equal(
    forEachStatement && forEachStatement.kind === "forEachStatement" ? forEachStatement.collectionText : "",
    "items"
  );
  assert.equal(nextItemStatement?.kind, "nextStatement");
  assert.equal(nextItemStatement && nextItemStatement.kind === "nextStatement" ? nextItemStatement.counterText : "", "item");
});

test("parseModule keeps block If and ElseIf statements structured when literals contain colons", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    If Format$(Now, "hh:mm") = "00:00" Then
    ElseIf stamp = #12:34:56 AM# Then
    End If
End Sub`, { fileName: "StructuredLiteralColonBlocks.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const ifStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[0] : undefined;
  const elseIfStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[1] : undefined;
  const endIfStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[2] : undefined;

  assert.ok(procedure && procedure.kind === "procedureDeclaration");
  assert.equal(ifStatement?.kind, "ifBlockStatement");
  assert.equal(
    ifStatement && ifStatement.kind === "ifBlockStatement" ? ifStatement.conditionText : "",
    'Format$(Now, "hh:mm") = "00:00"'
  );
  assert.equal(elseIfStatement?.kind, "elseIfClauseStatement");
  assert.equal(
    elseIfStatement && elseIfStatement.kind === "elseIfClauseStatement" ? elseIfStatement.conditionText : "",
    "stamp = #12:34:56 AM#"
  );
  assert.equal(endIfStatement?.kind, "endIfStatement");
});

test("parseModule structures Do, While, With, and On Error statements in procedure bodies", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    Do While keepRunning
    Loop Until finished
    While ready
    Wend
    With Application
    End With
    On Error Resume Next
    On Error GoTo Handler
Handler:
End Sub`, { fileName: "StructuredControlBlocks.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");

  assert.ok(procedure && procedure.kind === "procedureDeclaration");
  const doStatement = procedure.body[0];
  const loopStatement = procedure.body[1];
  const whileStatement = procedure.body[2];
  const wendStatement = procedure.body[3];
  const withStatement = procedure.body[4];
  const endWithStatement = procedure.body[5];
  const onErrorResumeStatement = procedure.body[6];
  const onErrorGotoStatement = procedure.body[7];

  assert.equal(doStatement?.kind, "doBlockStatement");
  assert.equal(doStatement && doStatement.kind === "doBlockStatement" ? doStatement.clauseKind : "", "while");
  assert.equal(doStatement && doStatement.kind === "doBlockStatement" ? doStatement.conditionText : "", "keepRunning");
  assert.equal(loopStatement?.kind, "loopStatement");
  assert.equal(loopStatement && loopStatement.kind === "loopStatement" ? loopStatement.clauseKind : "", "until");
  assert.equal(loopStatement && loopStatement.kind === "loopStatement" ? loopStatement.conditionText : "", "finished");
  assert.equal(whileStatement?.kind, "whileStatement");
  assert.equal(whileStatement && whileStatement.kind === "whileStatement" ? whileStatement.conditionText : "", "ready");
  assert.equal(wendStatement?.kind, "wendStatement");
  assert.equal(withStatement?.kind, "withBlockStatement");
  assert.equal(withStatement && withStatement.kind === "withBlockStatement" ? withStatement.targetText : "", "Application");
  assert.equal(endWithStatement?.kind, "endWithStatement");
  assert.equal(onErrorResumeStatement?.kind, "onErrorStatement");
  if (onErrorResumeStatement?.kind === "onErrorStatement") {
    assert.equal(onErrorResumeStatement.actionKind, "resumeNext");
  }
  assert.equal(onErrorGotoStatement?.kind, "onErrorStatement");
  if (onErrorGotoStatement?.kind === "onErrorStatement") {
    assert.equal(onErrorGotoStatement.actionKind, "goto");
    assert.equal(onErrorGotoStatement.targetText, "Handler");
  }
});

test("parseModule preserves single-character sub-range positions for structured block statements", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    For o = 1 To 10
    Next x
    Do While d
    Loop Until f
    While w
    With h
    On Error GoTo o
End Sub`, { fileName: "StructuredBlockRanges.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");

  assert.ok(procedure && procedure.kind === "procedureDeclaration");
  const forStatement = procedure.body[0];
  const nextStatement = procedure.body[1];
  const doStatement = procedure.body[2];
  const loopStatement = procedure.body[3];
  const whileStatement = procedure.body[4];
  const withStatement = procedure.body[5];
  const onErrorStatement = procedure.body[6];

  assert.equal(forStatement?.kind, "forStatement");
  if (forStatement?.kind === "forStatement") {
    assert.equal(forStatement.counterRange.start.character, 8);
    assert.equal(forStatement.startExpressionRange.start.character, 12);
    assert.equal(forStatement.endExpressionRange.start.character, 17);
  }

  assert.equal(nextStatement?.kind, "nextStatement");
  if (nextStatement?.kind === "nextStatement") {
    assert.equal(nextStatement.counterRange?.start.character, 9);
  }

  assert.equal(doStatement?.kind, "doBlockStatement");
  if (doStatement?.kind === "doBlockStatement") {
    assert.equal(doStatement.conditionRange?.start.character, 13);
  }

  assert.equal(loopStatement?.kind, "loopStatement");
  if (loopStatement?.kind === "loopStatement") {
    assert.equal(loopStatement.conditionRange?.start.character, 15);
  }

  assert.equal(whileStatement?.kind, "whileStatement");
  if (whileStatement?.kind === "whileStatement") {
    assert.equal(whileStatement.conditionRange.start.character, 10);
  }

  assert.equal(withStatement?.kind, "withBlockStatement");
  if (withStatement?.kind === "withBlockStatement") {
    assert.equal(withStatement.targetRange.start.character, 9);
  }

  assert.equal(onErrorStatement?.kind, "onErrorStatement");
  if (onErrorStatement?.kind === "onErrorStatement") {
    assert.equal(onErrorStatement.targetRange?.start.character, 18);
  }
});

test("parseModule structures GoTo, GoSub, and Resume statements in procedure bodies", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    GoTo Handler
    GoSub RetryPoint
    Resume
    Resume Next
    Resume Handler
Handler:
RetryPoint:
End Sub`, { fileName: "StructuredLabelTargets.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");

  assert.ok(procedure && procedure.kind === "procedureDeclaration");
  const goToStatement = procedure.body[0];
  const goSubStatement = procedure.body[1];
  const resumeStatement = procedure.body[2];
  const resumeNextStatement = procedure.body[3];
  const resumeTargetStatement = procedure.body[4];

  assert.equal(goToStatement?.kind, "goToStatement");
  if (goToStatement?.kind === "goToStatement") {
    assert.equal(goToStatement.actionKind, "goTo");
    assert.equal(goToStatement.targetText, "Handler");
    assert.equal(goToStatement.targetRange.start.character, 9);
  }

  assert.equal(goSubStatement?.kind, "goToStatement");
  if (goSubStatement?.kind === "goToStatement") {
    assert.equal(goSubStatement.actionKind, "goSub");
    assert.equal(goSubStatement.targetText, "RetryPoint");
    assert.equal(goSubStatement.targetRange.start.character, 10);
  }

  assert.equal(resumeStatement?.kind, "resumeStatement");
  if (resumeStatement?.kind === "resumeStatement") {
    assert.equal(resumeStatement.actionKind, "implicit");
    assert.equal(resumeStatement.targetRange, undefined);
  }

  assert.equal(resumeNextStatement?.kind, "resumeStatement");
  if (resumeNextStatement?.kind === "resumeStatement") {
    assert.equal(resumeNextStatement.actionKind, "next");
    assert.equal(resumeNextStatement.targetRange, undefined);
  }

  assert.equal(resumeTargetStatement?.kind, "resumeStatement");
  if (resumeTargetStatement?.kind === "resumeStatement") {
    assert.equal(resumeTargetStatement.actionKind, "target");
    assert.equal(resumeTargetStatement.targetText, "Handler");
    assert.equal(resumeTargetStatement.targetRange?.start.character, 11);
  }
});

test("parseModule structures Exit and End statements in procedure bodies", () => {
  const result = parseModule(`Option Explicit

Public Sub Demo()
    Exit Sub
    End
End Sub

Public Function Build() As Long
    Exit Function
End Function

Public Property Get Name() As String
    Exit Property
End Property`, { fileName: "StructuredTerminationStatements.bas" });
  const demo = result.module.members.find((member) => member.kind === "procedureDeclaration" && member.name === "Demo");
  const build = result.module.members.find((member) => member.kind === "procedureDeclaration" && member.name === "Build");
  const name = result.module.members.find((member) => member.kind === "procedureDeclaration" && member.name === "Name");
  const exitSubStatement = demo && demo.kind === "procedureDeclaration" ? demo.body[0] : undefined;
  const endStatement = demo && demo.kind === "procedureDeclaration" ? demo.body[1] : undefined;
  const exitFunctionStatement = build && build.kind === "procedureDeclaration" ? build.body[0] : undefined;
  const exitPropertyStatement = name && name.kind === "procedureDeclaration" ? name.body[0] : undefined;

  assert.equal(exitSubStatement?.kind, "exitStatement");
  if (exitSubStatement?.kind === "exitStatement") {
    assert.equal(exitSubStatement.exitKind, "Sub");
  }

  assert.equal(endStatement?.kind, "endStatement");
  assert.equal(exitFunctionStatement?.kind, "exitStatement");
  if (exitFunctionStatement?.kind === "exitStatement") {
    assert.equal(exitFunctionStatement.exitKind, "Function");
  }

  assert.equal(exitPropertyStatement?.kind, "exitStatement");
  if (exitPropertyStatement?.kind === "exitStatement") {
    assert.equal(exitPropertyStatement.exitKind, "Property");
  }
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

test("analyzeModule ignores label targets in GoTo, GoSub, Resume, and On Error statements", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredLabelTargets"
Option Explicit

Public Sub Demo()
    On Error GoTo Handler
    GoTo Handler
    GoSub RetryPoint
    Resume Next
    Resume Handler
Handler:
RetryPoint:
End Sub`, { fileName: "StructuredLabelTargets.bas" });
  const undeclaredDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "undeclared-variable");

  assert.deepEqual(undeclaredDiagnostics, []);
});

test("analyzeModule suppresses undeclared diagnostics for reserved and built-in reference data", () => {
  const result = analyzeModule(`Attribute VB_Name = "BuiltInReferences"
Option Explicit

Public Sub Demo()
    Beep
    Debug.Print Application.Name
    MsgBox xlAll
End Sub`, { fileName: "BuiltInReferences.bas" });

  assert.equal(result.diagnostics.some((diagnostic) => diagnostic.code === "undeclared-variable"), false);
});

test("analyzeModule reports undeclared callable names for structured call statements", () => {
  const result = analyzeModule(`Attribute VB_Name = "MissingCallable"
Option Explicit

Public Sub Demo()
    Dim value As Long
    MissingHandler value
End Sub`, { fileName: "MissingCallable.bas" });

  assert.deepEqual(
    result.diagnostics
      .filter((diagnostic) => diagnostic.code === "undeclared-variable")
      .map((diagnostic) => ({
        message: diagnostic.message,
        start: `${diagnostic.range.start.line}:${diagnostic.range.start.character}`
      })),
    [
      {
        message: "Undeclared identifier 'MissingHandler'.",
        start: "5:4"
      }
    ]
  );
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

test("symbol resolution lets procedure variables shadow module variables of the same kind", () => {
  const result = analyzeModule(`Attribute VB_Name = "SameKindShadowing"
Option Explicit

Private SharedValue As String

Public Sub Demo()
    Dim SharedValue
    SharedValue = 1
    Debug.Print SharedValue
End Sub`, { fileName: "SameKindShadowing.bas" });

  const declarationDefinition = findDefinition(result, { character: 8, line: 6 });
  const usageDefinition = findDefinition(result, { character: 4, line: 7 });
  const completionVariables = getCompletionSymbols(result, { character: 4, line: 8 }).filter(
    (symbol) => symbol.kind === "variable" && symbol.name === "SharedValue"
  );
  const localSymbol = result.symbols.procedureScopes
    .flatMap((scope) => scope.symbols)
    .find((symbol) => symbol.kind === "variable" && symbol.name === "SharedValue");

  assert.equal(declarationDefinition?.scope, "procedure");
  assert.equal(usageDefinition?.scope, "procedure");
  assert.deepEqual(
    completionVariables.map((symbol) => symbol.scope),
    ["procedure"]
  );
  assert.equal(getSymbolTypeName(result, localSymbol), "Long");
  assert.equal(result.diagnostics.some((diagnostic) => diagnostic.code === "type-mismatch"), false);
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

test("analyzeModule infers known CreateObject ProgID result types", () => {
  const result = analyzeModule(`Attribute VB_Name = "KnownProgIdInference"
Option Explicit

Public Sub Demo()
    Dim shell
    Set shell = CreateObject("WScript.Shell")
End Sub`, { fileName: "KnownProgIdInference.bas" });

  const shellSymbol = result.symbols.procedureScopes
    .flatMap((scope) => scope.symbols)
    .find((symbol) => symbol.kind === "variable" && symbol.name === "shell");

  assert.equal(getSymbolTypeName(result, shellSymbol), "WshShell");
  assert.equal(result.diagnostics.some((diagnostic) => diagnostic.code === "set-required"), false);
  assert.equal(result.diagnostics.some((diagnostic) => diagnostic.code === "type-mismatch"), false);
});

test("analyzeModule keeps GetObject pathname arguments as Object", () => {
  const result = analyzeModule(`Attribute VB_Name = "GetObjectInference"
Option Explicit

Public Sub Demo()
    Dim shell
    Set shell = GetObject("WScript.Shell")
End Sub`, { fileName: "GetObjectInference.bas" });

  const shellSymbol = result.symbols.procedureScopes
    .flatMap((scope) => scope.symbols)
    .find((symbol) => symbol.kind === "variable" && symbol.name === "shell");

  assert.equal(getSymbolTypeName(result, shellSymbol), "Object");
  assert.equal(result.diagnostics.some((diagnostic) => diagnostic.code === "set-required"), false);
  assert.equal(result.diagnostics.some((diagnostic) => diagnostic.code === "type-mismatch"), false);
});

test("analyzeModule keeps dynamic CreateObject ProgID expressions as Object", () => {
  const result = analyzeModule(`Attribute VB_Name = "DynamicProgIdInference"
Option Explicit

Public Sub Demo()
    Dim shell
    Set shell = CreateObject("WScript." & "Shell")
End Sub`, { fileName: "DynamicProgIdInference.bas" });

  const shellSymbol = result.symbols.procedureScopes
    .flatMap((scope) => scope.symbols)
    .find((symbol) => symbol.kind === "variable" && symbol.name === "shell");

  assert.equal(getSymbolTypeName(result, shellSymbol), "Object");
  assert.equal(result.diagnostics.some((diagnostic) => diagnostic.code === "set-required"), false);
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

test("analyzeModule reports type mismatches on the physical expression line for continued assignments", () => {
  const result = analyzeModule(`Attribute VB_Name = "ContinuedMismatchRange"
Option Explicit

Public Sub Demo()
    Dim title As String
    title = _
        True
End Sub`, { fileName: "ContinuedMismatchRange.bas" });

  const mismatchDiagnostic = result.diagnostics.find((diagnostic) => diagnostic.code === "type-mismatch");

  assert.equal(mismatchDiagnostic?.message, "Type mismatch: cannot assign Boolean to String.");
  assert.deepEqual(mismatchDiagnostic?.range, {
    end: { character: 12, line: 6 },
    start: { character: 8, line: 6 }
  });
});

test("analyzeModule reports type mismatches for labeled assignments", () => {
  const result = analyzeModule(`Attribute VB_Name = "LabeledMismatch"
Option Explicit

Public Sub Demo()
    Dim title As String
Label1: title = 1
End Sub`, { fileName: "LabeledMismatch.bas" });

  const mismatchDiagnostic = result.diagnostics.find((diagnostic) => diagnostic.code === "type-mismatch");

  assert.equal(mismatchDiagnostic?.message, "Type mismatch: cannot assign Long to String.");
  assert.deepEqual(mismatchDiagnostic?.range, {
    end: { character: 17, line: 5 },
    start: { character: 16, line: 5 }
  });
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

test("analyzeModule ignores omitted optional ByRef arguments", () => {
  const result = analyzeModule(`Attribute VB_Name = "ByRefOptional"
Option Explicit

Private Sub UpdateCount(Optional count As Variant, ByRef nextCount As Long)
End Sub

Public Sub Demo()
    Dim nextCount As Long
    UpdateCount , nextCount
End Sub`, { fileName: "ByRefOptional.bas" });

  const byRefDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code.startsWith("byref-"));

  assert.equal(byRefDiagnostics.length, 0);
});

test("analyzeModule keeps block header reads and ByRef checks after structured control statement parsing", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredControlDiagnostics"
Option Explicit

Private Function AcceptCount(ByRef count As Long) As Boolean
    AcceptCount = True
End Function

Public Sub Demo()
    Dim ready As Boolean
    Dim fallback As Boolean
    Dim count As Long
    Dim item As Variant
    Dim items As Collection
    Set items = New Collection

    If AcceptCount(count + 1) Then
        ready = True
    ElseIf fallback Then
        ready = False
    End If

    For Each item In items
        Debug.Print ready, item
    Next item
  End Sub`, { fileName: "StructuredControlDiagnostics.bas" });
  const unusedDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unused-variable");
  const byRefDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code.startsWith("byref-"));

  assert.deepEqual(
    unusedDiagnostics.map((diagnostic) => diagnostic.message),
    ["Unused local declaration 'count'."]
  );
  assert.equal(byRefDiagnostics.length, 1);
  assert.equal(
    byRefDiagnostics[0]?.message,
    "ByRef parameter 'count' in AcceptCount receives an expression. Introduce a temporary variable before the call."
  );
});

test("analyzeModule avoids false block syntax errors when If headers contain literal colons", () => {
  const result = analyzeModule(`Attribute VB_Name = "LiteralColonBlocks"
Option Explicit

Public Sub Demo()
    Dim stamp As Date
    stamp = #12:00:00 AM#

    If Format$(stamp, "hh:mm") = "00:00" Then
        Debug.Print "midnight"
    ElseIf stamp = #12:34:56 AM# Then
        Debug.Print "fallback"
    End If
End Sub`, { fileName: "LiteralColonBlocks.bas" });

  assert.deepEqual(
    result.diagnostics.filter((diagnostic) => diagnostic.code === "syntax-error"),
    []
  );
});

test("analyzeModule keeps ByRef checks in structured For headers after AST-based invocation scanning", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredForByRef"
Option Explicit

Private Function ReadLimit(ByRef count As Long) As Long
    ReadLimit = count
End Function

Public Sub Demo()
    Dim count As Long
    Dim index As Long

    For index = 1 To ReadLimit(count + 1)
    Next index
End Sub`, { fileName: "StructuredForByRef.bas" });
  const byRefDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code.startsWith("byref-"));

  assert.equal(byRefDiagnostics.length, 1);
  assert.equal(
    byRefDiagnostics[0]?.message,
    "ByRef parameter 'count' in ReadLimit receives an expression. Introduce a temporary variable before the call."
  );
});

test("analyzeModule keeps ByRef checks in multiline structured If headers when invocation stays on one line", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredMultilineIfByRef"
Option Explicit

Private Function AcceptCount(ByRef count As Long) As Boolean
    AcceptCount = True
End Function

Public Sub Demo()
    Dim count As Long
    Dim ready As Boolean

    If AcceptCount(count + 1) And _
        ready Then
        Debug.Print ready
    End If
End Sub`, { fileName: "StructuredMultilineIfByRef.bas" });
  const byRefDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code.startsWith("byref-"));

  assert.equal(byRefDiagnostics.length, 1);
  assert.equal(
    byRefDiagnostics[0]?.message,
    "ByRef parameter 'count' in AcceptCount receives an expression. Introduce a temporary variable before the call."
  );
});

test("analyzeModule keeps ByRef checks when call invocations span multiple physical lines", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredMultilineCallByRef"
Option Explicit

Private Sub UpdateCount(ByRef count As Long)
End Sub

Public Sub Demo()
    Dim count As Long

    Call UpdateCount( _
        count + 1 _
    )
End Sub`, { fileName: "StructuredMultilineCallByRef.bas" });
  const byRefDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code.startsWith("byref-"));

  assert.equal(byRefDiagnostics.length, 1);
  assert.equal(
    byRefDiagnostics[0]?.message,
    "ByRef parameter 'count' in UpdateCount receives an expression. Introduce a temporary variable before the call."
  );
});

test("analyzeModule keeps ByRef checks for labeled call statements", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredLabeledCallByRef"
Option Explicit

Private Sub UpdateCount(ByRef count As Long)
End Sub

Public Sub Demo()
    Dim wrongCount As String

Label1: UpdateCount wrongCount
End Sub`, { fileName: "StructuredLabeledCallByRef.bas" });
  const byRefDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code.startsWith("byref-"));

  assert.equal(byRefDiagnostics.length, 1);
  assert.equal(
    byRefDiagnostics[0]?.message,
    "ByRef parameter 'count' in UpdateCount expects Long but receives String. VBA may raise a ByRef argument type mismatch."
  );
});

test("analyzeModule keeps Do, While, and With header reads after structured control statement parsing", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredLoopReads"
Option Explicit

Public Sub Demo()
    Dim keepRunning As Boolean
    Dim finished As Boolean
    Dim ready As Boolean
    Dim holder As Collection
    Set holder = New Collection

    Do While keepRunning
    Loop Until finished

    While ready
    Wend

    With holder
        Debug.Print .Count
    End With
End Sub`, { fileName: "StructuredLoopReads.bas" });
  const unusedDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unused-variable");

  assert.equal(unusedDiagnostics.length, 0);
});

test("analyzeModule reports undeclared identifiers in structured control headers", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredHeaderUndeclared"
Option Explicit

Public Sub Demo()
    Dim index As Long
    Dim item As Variant

    If ready Then
    ElseIf fallback Then
    End If

    Select Case selector
        Case caseValue
    End Select

    For index = 1 To limit Step stepSize
    Next index

    For Each item In items
    Next item

    Do While keepGoing
    Loop Until finished

    While active
    Wend

    With holder
    End With
End Sub`, { fileName: "StructuredHeaderUndeclared.bas" });
  const undeclaredDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "undeclared-variable");

  assert.deepEqual(
    undeclaredDiagnostics.map((diagnostic) => diagnostic.message),
    [
      "Undeclared identifier 'ready'.",
      "Undeclared identifier 'fallback'.",
      "Undeclared identifier 'selector'.",
      "Undeclared identifier 'caseValue'.",
      "Undeclared identifier 'limit'.",
      "Undeclared identifier 'stepSize'.",
      "Undeclared identifier 'items'.",
      "Undeclared identifier 'keepGoing'.",
      "Undeclared identifier 'finished'.",
      "Undeclared identifier 'active'.",
      "Undeclared identifier 'holder'."
    ]
  );
});

test("analyzeModule reports undeclared identifiers on the physical line inside multiline structured headers", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredMultilineHeaderUndeclared"
Option Explicit

Public Sub Demo()
    If ready And _
        fallback Then
        Debug.Print "x"
    End If
End Sub`, { fileName: "StructuredMultilineHeaderUndeclared.bas" });
  const undeclaredDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "undeclared-variable");

  assert.deepEqual(
    undeclaredDiagnostics.map((diagnostic) => ({
      message: diagnostic.message,
      start: `${diagnostic.range.start.line}:${diagnostic.range.start.character}`
    })),
    [
      {
        message: "Undeclared identifier 'ready'.",
        start: "4:7"
      },
      {
        message: "Undeclared identifier 'fallback'.",
        start: "5:8"
      }
    ]
  );
});

test("analyzeModule uses structured Const initializer references in diagnostics", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredConstInitializerReferences"
Option Explicit

Public Sub Demo()
    Const baseValue As Long = 1
    Const usedValue As Long = baseValue
    Const missingValue As Long = missingConst
    Debug.Print usedValue
End Sub`, { fileName: "StructuredConstInitializerReferences.bas" });

  const undeclaredDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "undeclared-variable");
  const unusedDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unused-variable");

  assert.deepEqual(
    undeclaredDiagnostics.map((diagnostic) => ({
      message: diagnostic.message,
      start: `${diagnostic.range.start.line}:${diagnostic.range.start.character}`
    })),
    [
      {
        message: "Undeclared identifier 'missingConst'.",
        start: "6:33"
      }
    ]
  );
  assert.deepEqual(
    unusedDiagnostics.map((diagnostic) => diagnostic.message),
    ["Unused local declaration 'missingValue'."]
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

test("analyzeModule uses structured multiline assignments for type inference", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredAssignmentInference"
Option Explicit

Public Sub Demo()
    Dim items As Collection
Label1: items = _
        New Collection
End Sub`, { fileName: "StructuredAssignmentInference.bas" });

  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const assignment = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[1] : undefined;
  const setDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "set-required");

  assert.equal(assignment?.kind, "assignmentStatement");
  if (assignment?.kind === "assignmentStatement") {
    assert.equal(assignment.assignmentKind, "implicit");
    assert.equal(assignment.targetText, "items");
    assert.equal(assignment.expressionText, "New Collection");
  }
  assert.deepEqual(
    setDiagnostics.map((diagnostic) => ({
      message: diagnostic.message,
      start: `${diagnostic.range.start.line}:${diagnostic.range.start.character}`
    })),
    [
      {
        message: "Set is required to assign Collection to Collection.",
        start: "5:8"
      }
    ]
  );
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

test("analyzeModule uses structured termination statements for unreachable code", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredTerminationUnreachable"
Option Explicit

Public Function Build() As Long
    Dim marker As Long
    Exit Function
    marker = 1
End Function`, { fileName: "StructuredTerminationUnreachable.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const exitStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[1] : undefined;
  const unreachableDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unreachable-code");

  assert.equal(exitStatement?.kind, "exitStatement");
  assert.deepEqual(
    unreachableDiagnostics.map((diagnostic) => diagnostic.message),
    ["Unreachable code after Exit Function."]
  );
});

test("analyzeModule uses labeled structured property exits for unreachable code", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredLabeledPropertyExit"
Option Explicit

Public Property Get Name() As String
LabelExit: Exit Property
    Name = "fallback"
End Property`, { fileName: "StructuredLabeledPropertyExit.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const exitStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[0] : undefined;
  const unreachableDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unreachable-code");

  assert.equal(exitStatement?.kind, "exitStatement");
  assert.deepEqual(
    unreachableDiagnostics.map((diagnostic) => diagnostic.message),
    ["Unreachable code after Exit Property."]
  );
});

test("analyzeModule uses structured leading labels to clear unreachable state", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredLabeledReachable"
Option Explicit

Public Sub Demo()
    Dim marker As Long
    Exit Sub
Label1: marker = 1
End Sub`, { fileName: "StructuredLabeledReachable.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const assignmentStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[2] : undefined;
  const unreachableDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unreachable-code");

  assert.equal(assignmentStatement?.kind, "assignmentStatement");
  assert.equal(assignmentStatement?.leadingLabel?.text, "Label1");
  assert.deepEqual(unreachableDiagnostics, []);
});

test("analyzeModule uses executable label fallback to clear unreachable state", () => {
  const result = analyzeModule(`Attribute VB_Name = "ExecutableLabelReachable"
Option Explicit

Public Sub Demo()
    Dim marker As Long
    Exit Sub
Label1:
    marker = 1
End Sub`, { fileName: "ExecutableLabelReachable.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const labelStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[2] : undefined;
  const assignmentStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[3] : undefined;
  const unreachableDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unreachable-code");

  assert.equal(labelStatement?.kind, "executableStatement");
  assert.equal(labelStatement?.text, "Label1:");
  assert.equal(assignmentStatement?.kind, "assignmentStatement");
  assert.deepEqual(unreachableDiagnostics, []);
});

test("analyzeModule ignores mismatched structured exit kind for unreachable code", () => {
  const result = analyzeModule(`Attribute VB_Name = "MismatchedStructuredExit"
Option Explicit

Public Function Build() As Long
    Dim marker As Long
    Exit Sub
    marker = 1
End Function`, { fileName: "MismatchedStructuredExit.bas" });
  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const exitStatement = procedure && procedure.kind === "procedureDeclaration" ? procedure.body[1] : undefined;
  const unreachableDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unreachable-code");

  assert.equal(exitStatement?.kind, "exitStatement");
  assert.deepEqual(unreachableDiagnostics, []);
});

test("analyzeModule clears unreachable state at structured Case and Next boundaries", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredUnreachableBoundaries"
Option Explicit

Public Sub Demo()
    Dim value As Long
    Dim index As Long

    Select Case value
        Case 0
            Exit Sub
            value = 1
        Case Else
            value = 2
    End Select

    For index = 1 To 2
        End
        index = index + 1
    Next index
End Sub`, { fileName: "StructuredUnreachableBoundaries.bas" });

  const unreachableDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unreachable-code");

  assert.equal(unreachableDiagnostics.length, 2);
  assert.deepEqual(
    unreachableDiagnostics.map((diagnostic) => diagnostic.message),
    [
      "Unreachable code after Exit Sub.",
      "Unreachable code after End."
    ]
  );
});

test("analyzeModule uses structured Else, Loop, Wend, and End With boundaries for unreachable code", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredUnreachableControlBoundaries"
Option Explicit

Public Sub Demo()
    Dim ready As Boolean
    Dim keepRunning As Boolean
    Dim marker As Long

    If ready Then
        Exit Sub
        marker = 1
    Else
        marker = 2
    End If

    Do While keepRunning
        End
        marker = 3
    Loop

    While ready
        End
        marker = 4
    Wend

    With Application
        End
    End With

    marker = 5
End Sub`, { fileName: "StructuredUnreachableControlBoundaries.bas" });

  const unreachableDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unreachable-code");

  assert.equal(unreachableDiagnostics.length, 4);
  assert.deepEqual(
    unreachableDiagnostics.map((diagnostic) => diagnostic.message),
    [
      "Unreachable code after Exit Sub.",
      "Unreachable code after End.",
      "Unreachable code after End.",
      "Unreachable code after End."
    ]
  );
});

test("analyzeModule keeps structured ElseIf boundaries quiet after unreachable exits", () => {
  const result = analyzeModule(`Attribute VB_Name = "StructuredElseIfUnreachableBoundaries"
Option Explicit

Public Sub Demo()
    Dim ready As Boolean
    Dim fallback As Boolean
    Dim marker As Long

    If ready Then
        Exit Sub
        marker = 1
    ElseIf Format$(Now, "hh:mm") = "12:34" And fallback Then
        marker = 2
    End If

    marker = 3
End Sub`, { fileName: "StructuredElseIfUnreachableBoundaries.bas" });

  const procedure = result.module.members.find((member) => member.kind === "procedureDeclaration");
  const unreachableDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unreachable-code");
  const unreachableStatement =
    procedure && procedure.kind === "procedureDeclaration"
      ? procedure.body.find(
          (statement) =>
            statement.kind === "assignmentStatement" &&
            statement.targetText === "marker" &&
            statement.expressionText === "1"
        )
      : undefined;

  assert.ok(unreachableStatement?.kind === "assignmentStatement");
  assert.deepEqual(
    unreachableDiagnostics.map((diagnostic) => ({
      message: diagnostic.message,
      start: diagnostic.range.start
    })),
    [
      {
        message: "Unreachable code after Exit Sub.",
        start: unreachableStatement.range.start
      }
    ]
  );
});

test("analyzeModule keeps outer unreachable state across nested inner If boundaries", () => {
  const result = analyzeModule(`Attribute VB_Name = "NestedStructuredIfUnreachableBoundaries"
Option Explicit

Public Sub Demo()
    Dim ready As Boolean
    Dim innerReady As Boolean
    Dim marker As Long

    If ready Then
        Exit Sub
        If innerReady Then
        End If
        marker = 1
    Else
        marker = 2
    End If
End Sub`, { fileName: "NestedStructuredIfUnreachableBoundaries.bas" });

  const unreachableDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unreachable-code");

  assert.deepEqual(
    unreachableDiagnostics.map((diagnostic) => ({
      message: diagnostic.message,
      start: `${diagnostic.range.start.line}:${diagnostic.range.start.character}`
    })),
    [
      {
        message: "Unreachable code after Exit Sub.",
        start: "10:0"
      },
      {
        message: "Unreachable code after Exit Sub.",
        start: "12:0"
      }
    ]
  );
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

test("analyzeModule keeps nested invocation reads when statements contain parenthesized calls", () => {
  const result = analyzeModule(`Attribute VB_Name = "NestedInvocationReads"
Option Explicit

Public Sub Demo()
    Dim leftValue As String
    Dim rightValue As String
    leftValue = "A"
    rightValue = "B"
    Debug.Print leftValue & Len(rightValue)
End Sub`, { fileName: "NestedInvocationReads.bas" });

  const unusedDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unused-variable");
  const writeOnlyDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "write-only-variable");

  assert.equal(unusedDiagnostics.length, 0);
  assert.equal(writeOnlyDiagnostics.length, 0);
});

test("analyzeModule keeps dotted call receivers as local reads", () => {
  const result = analyzeModule(`Attribute VB_Name = "DottedCallReceiverReads"
Option Explicit

Public Sub Demo()
    Dim items As Collection
    Set items = New Collection
    items.Add 1
End Sub`, { fileName: "DottedCallReceiverReads.bas" });

  const unusedDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unused-variable");
  const writeOnlyDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "write-only-variable");

  assert.equal(unusedDiagnostics.length, 0);
  assert.equal(writeOnlyDiagnostics.length, 0);
});

test("analyzeModule keeps indexed assignment writes after structured assignment parsing", () => {
  const result = analyzeModule(`Attribute VB_Name = "IndexedAssignmentWrites"
Option Explicit

Public Sub Demo()
    Dim values(1 To 2) As Long
    values(1) = 42
End Sub`, { fileName: "IndexedAssignmentWrites.bas" });

  const unusedDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "unused-variable");
  const writeOnlyDiagnostics = result.diagnostics.filter((diagnostic) => diagnostic.code === "write-only-variable");

  assert.equal(unusedDiagnostics.length, 0);
  assert.deepEqual(
    writeOnlyDiagnostics.map((diagnostic) => diagnostic.message),
    ["Write-only local variable 'values'."]
  );
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

test("formatModuleIndentation keeps block If headers with literal colons aligned as blocks", () => {
  const formatted = formatModuleIndentation(`Attribute VB_Name = "LiteralColonFormatting"
Option Explicit

Public Sub Demo()
Dim stamp As Date
stamp = #12:00:00 AM#
If Format$(stamp, "hh:mm") = "00:00" Then
Debug.Print "midnight"
ElseIf stamp = #12:34:56 AM# Then
Debug.Print "fallback"
End If
End Sub`, { fileName: "LiteralColonFormatting.bas", indentSize: 4, insertSpaces: true });

  assert.equal(
    formatted,
    `Attribute VB_Name = "LiteralColonFormatting"
Option Explicit

Public Sub Demo()
    Dim stamp As Date
    stamp = #12:00:00 AM#
    If Format$(stamp, "hh:mm") = "00:00" Then
        Debug.Print "midnight"
    ElseIf stamp = #12:34:56 AM# Then
        Debug.Print "fallback"
    End If
End Sub`
  );
});

test("formatModuleIndentation keeps labeled block headers aligned with structured block bodies", () => {
  const formatted = formatModuleIndentation(`Attribute VB_Name = "LabeledBlockFormatting"
Option Explicit

Public Sub Demo()
Label1: If True Then
Debug.Print "ready"
Else
Debug.Print "fallback"
End If
End Sub`, { fileName: "LabeledBlockFormatting.bas", indentSize: 4, insertSpaces: true });

  assert.equal(
    formatted,
    `Attribute VB_Name = "LabeledBlockFormatting"
Option Explicit

Public Sub Demo()
    Label1: If True Then
        Debug.Print "ready"
    Else
        Debug.Print "fallback"
    End If
End Sub`
  );
});

test("formatModuleIndentation keeps structured non-block statements at the active block indentation", () => {
  const formatted = formatModuleIndentation(`Attribute VB_Name = "StructuredNonBlockFormatting"
Option Explicit

Public Sub Demo()
If True Then
value = 1
Call Trace(value)
On Error GoTo HandleError
GoTo Done
Resume Next
Exit Sub
End
End If
Done:
Debug.Print value
HandleError:
Resume Done
End Sub`, { fileName: "StructuredNonBlockFormatting.bas", indentSize: 4, insertSpaces: true });

  assert.equal(
    formatted,
    `Attribute VB_Name = "StructuredNonBlockFormatting"
Option Explicit

Public Sub Demo()
    If True Then
        value = 1
        Call Trace(value)
        On Error GoTo HandleError
        GoTo Done
        Resume Next
        Exit Sub
        End
    End If
Done:
    Debug.Print value
HandleError:
    Resume Done
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
