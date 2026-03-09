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

  const builtInMemberCompletionDocument = await vscode.workspace.openTextDocument(
    path.resolve(fixturesPath, "BuiltInMemberCompletion.bas")
  );
  await vscode.window.showTextDocument(builtInMemberCompletionDocument);

  const applicationMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    new vscode.Position(4, 28),
    (items) => items.some((item) => getCompletionItemLabel(item) === "WorksheetFunction")
  );
  const worksheetFunctionMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    new vscode.Position(5, 36),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Sum")
  );
  const chainedWorksheetFunctionMemberCompletionItems = await waitForCompletions(
    builtInMemberCompletionDocument,
    new vscode.Position(6, 48),
    (items) => items.some((item) => getCompletionItemLabel(item) === "Sum")
  );
  const worksheetFunctionPropertyCompletion = applicationMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "WorksheetFunction"
  );
  const worksheetFunctionSumCompletion = worksheetFunctionMemberCompletionItems.find(
    (item) => getCompletionItemLabel(item) === "Sum"
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

function getSignatureDocumentation(
  documentation: vscode.MarkdownString | vscode.ParameterInformation["documentation"] | vscode.SignatureInformation["documentation"] | undefined
): string {
  if (!documentation) {
    return "";
  }

  return typeof documentation === "string" ? documentation : documentation.value;
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

function getCodeAction(actions: readonly vscode.CodeAction[], title: string): vscode.CodeAction | undefined {
  return actions.find((action) => action.title === title);
}

function hasCodeAction(actions: readonly vscode.CodeAction[], title: string): boolean {
  return getCodeAction(actions, title) !== undefined;
}

function isCodeAction(action: vscode.CodeAction | vscode.Command): action is vscode.CodeAction {
  return "title" in action && "edit" in action;
}
