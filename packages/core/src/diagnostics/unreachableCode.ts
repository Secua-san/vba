import type { Diagnostic, ParseResult, ProcedureDeclarationNode, ProcedureKind } from "../types/model";
import { hasStatementSeparatorColon } from "../parser/text";

type BlockKind = "do" | "for" | "if" | "select" | "while" | "with";
type BarrierKind = "do" | "for" | "if" | "procedure" | "select" | "while";

interface UnreachableState {
  barrierKind: BarrierKind;
  reason: string;
}

export function collectUnreachableCodeDiagnostics(parseResult: ParseResult): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];

  for (const member of parseResult.module.members) {
    if (member.kind !== "procedureDeclaration") {
      continue;
    }

    diagnostics.push(...collectProcedureUnreachableDiagnostics(member));
  }

  return diagnostics;
}

function collectProcedureUnreachableDiagnostics(procedure: ProcedureDeclarationNode): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];
  const blockStack: BlockKind[] = [];
  let unreachableState: UnreachableState | undefined;

  for (const statement of procedure.body) {
    const rawText = statement.text.trim();

    if (rawText.length === 0 || /^#/i.test(rawText)) {
      continue;
    }

    const hasLabel = hasLeadingLabel(rawText);
    const controlText = stripLeadingLabel(rawText);

    if (hasLabel) {
      unreachableState = undefined;
    }

    if (unreachableState && clearsUnreachableState(statement, controlText, unreachableState)) {
      unreachableState = undefined;
    }

    if (unreachableState && shouldReportUnreachableStatement(statement, controlText)) {
      diagnostics.push({
        code: "unreachable-code",
        message: `Unreachable code after ${unreachableState.reason}.`,
        range: statement.range,
        severity: "warning"
      });
    }

    if (isUnconditionalProcedureExit(controlText, procedure.procedureKind)) {
      unreachableState = {
        barrierKind: getBarrierKind(blockStack),
        reason: normalizeTerminationReason(controlText)
      };
    }

    applyBlockTransition(statement, controlText, blockStack);
  }

  return diagnostics;
}

function applyBlockTransition(statement: ProcedureDeclarationNode["body"][number], text: string, blockStack: BlockKind[]): void {
  const transition = getStatementBlockTransition(statement, text);

  if (transition?.push) {
    blockStack.push(transition.push);
    return;
  }

  if (transition?.pop) {
    popLastBlockOfKind(blockStack, transition.pop);
    return;
  }
}

function clearsUnreachableState(
  statement: ProcedureDeclarationNode["body"][number],
  text: string,
  unreachableState: UnreachableState
): boolean {
  switch (unreachableState.barrierKind) {
    case "if":
      return isIfBoundaryStatement(statement, text);
    case "select":
      return isSelectBoundaryStatement(statement, text);
    case "for":
      return isForBoundaryStatement(statement, text);
    case "do":
      return isDoBoundaryStatement(statement, text);
    case "while":
      return isWhileBoundaryStatement(statement, text);
    default:
      return false;
  }
}

function getBarrierKind(blockStack: BlockKind[]): BarrierKind {
  for (let index = blockStack.length - 1; index >= 0; index -= 1) {
    const blockKind = blockStack[index];

    if (blockKind !== "with") {
      return blockKind;
    }
  }

  return "procedure";
}

function hasLeadingLabel(text: string): boolean {
  return /^(?:[A-Za-z_][A-Za-z0-9_]*|\d+):/u.test(text);
}

function isIfBlockStart(text: string): boolean {
  return !/^ElseIf\b/i.test(text) && /^If\b.*\bThen\s*$/i.test(text) && !hasStatementSeparatorColon(text);
}

function isUnconditionalProcedureExit(text: string, procedureKind: ProcedureKind): boolean {
  if (/^End$/i.test(text)) {
    return true;
  }

  switch (procedureKind) {
    case "Function":
      return /^Exit\s+Function$/i.test(text);
    case "PropertyGet":
    case "PropertyLet":
    case "PropertySet":
      return /^Exit\s+Property$/i.test(text);
    default:
      return /^Exit\s+Sub$/i.test(text);
  }
}

function normalizeTerminationReason(text: string): string {
  return /^End$/i.test(text) ? "End" : text.replace(/\s+/g, " ").trim();
}

function popLastBlockOfKind(blockStack: BlockKind[], blockKind: BlockKind): void {
  for (let index = blockStack.length - 1; index >= 0; index -= 1) {
    if (blockStack[index] === blockKind) {
      blockStack.splice(index, 1);
      return;
    }
  }
}

function shouldReportUnreachableStatement(statement: ProcedureDeclarationNode["body"][number], text: string): boolean {
  if (text.length === 0) {
    return false;
  }

  if (isStructuredBoundaryStatement(statement)) {
    return false;
  }

  if (statement.kind !== "executableStatement") {
    return true;
  }

  return !isExecutableBoundaryStatement(text);
}

function getStatementBlockTransition(
  statement: ProcedureDeclarationNode["body"][number],
  text: string
): { pop?: BlockKind; push?: BlockKind } | undefined {
  return getStructuredBlockTransition(statement) ?? getExecutableBlockTransition(statement, text);
}

function getStructuredBlockTransition(
  statement: ProcedureDeclarationNode["body"][number]
): { pop?: BlockKind; push?: BlockKind } | undefined {
  switch (statement.kind) {
    case "ifBlockStatement":
      return { push: "if" };
    case "selectCaseStatement":
      return { push: "select" };
    case "forStatement":
    case "forEachStatement":
      return { push: "for" };
    case "doBlockStatement":
      return { push: "do" };
    case "whileStatement":
      return { push: "while" };
    case "withBlockStatement":
      return { push: "with" };
    case "endIfStatement":
      return { pop: "if" };
    case "endSelectStatement":
      return { pop: "select" };
    case "nextStatement":
      return { pop: "for" };
    case "loopStatement":
      return { pop: "do" };
    case "wendStatement":
      return { pop: "while" };
    case "endWithStatement":
      return { pop: "with" };
    default:
      return undefined;
  }
}

function getExecutableBlockTransition(
  statement: ProcedureDeclarationNode["body"][number],
  text: string
): { pop?: BlockKind; push?: BlockKind } | undefined {
  if (statement.kind !== "executableStatement") {
    return undefined;
  }

  if (isIfBlockStart(text)) {
    return { push: "if" };
  }

  if (/^Select\s+Case\b/i.test(text)) {
    return { push: "select" };
  }

  if (/^For\b/i.test(text)) {
    return { push: "for" };
  }

  if (/^Do\b/i.test(text)) {
    return { push: "do" };
  }

  if (/^While\b/i.test(text)) {
    return { push: "while" };
  }

  if (/^With\b/i.test(text)) {
    return { push: "with" };
  }

  if (/^End\s+If\b/i.test(text)) {
    return { pop: "if" };
  }

  if (/^End\s+Select\b/i.test(text)) {
    return { pop: "select" };
  }

  if (/^Next\b/i.test(text)) {
    return { pop: "for" };
  }

  if (/^Loop\b/i.test(text)) {
    return { pop: "do" };
  }

  if (/^Wend\b/i.test(text)) {
    return { pop: "while" };
  }

  if (/^End\s+With\b/i.test(text)) {
    return { pop: "with" };
  }

  return undefined;
}

function isStructuredBoundaryStatement(statement: ProcedureDeclarationNode["body"][number]): boolean {
  return (
    statement.kind === "elseIfClauseStatement" ||
    statement.kind === "elseClauseStatement" ||
    statement.kind === "caseClauseStatement" ||
    statement.kind === "endIfStatement" ||
    statement.kind === "endSelectStatement" ||
    statement.kind === "nextStatement" ||
    statement.kind === "loopStatement" ||
    statement.kind === "wendStatement" ||
    statement.kind === "endWithStatement"
  );
}

function isExecutableBoundaryStatement(text: string): boolean {
  return (
    isExecutableIfBoundary(text) ||
    isExecutableSelectBoundary(text) ||
    isExecutableForBoundary(text) ||
    isExecutableDoBoundary(text) ||
    isExecutableWhileBoundary(text) ||
    isExecutableWithBoundary(text)
  );
}

function isIfBoundaryStatement(statement: ProcedureDeclarationNode["body"][number], text: string): boolean {
  return (
    statement.kind === "elseIfClauseStatement" ||
    statement.kind === "elseClauseStatement" ||
    statement.kind === "endIfStatement" ||
    (statement.kind === "executableStatement" && isExecutableIfBoundary(text))
  );
}

function isSelectBoundaryStatement(statement: ProcedureDeclarationNode["body"][number], text: string): boolean {
  return (
    statement.kind === "caseClauseStatement" ||
    statement.kind === "endSelectStatement" ||
    (statement.kind === "executableStatement" && isExecutableSelectBoundary(text))
  );
}

function isForBoundaryStatement(statement: ProcedureDeclarationNode["body"][number], text: string): boolean {
  return statement.kind === "nextStatement" || (statement.kind === "executableStatement" && isExecutableForBoundary(text));
}

function isDoBoundaryStatement(statement: ProcedureDeclarationNode["body"][number], text: string): boolean {
  return statement.kind === "loopStatement" || (statement.kind === "executableStatement" && isExecutableDoBoundary(text));
}

function isWhileBoundaryStatement(statement: ProcedureDeclarationNode["body"][number], text: string): boolean {
  return statement.kind === "wendStatement" || (statement.kind === "executableStatement" && isExecutableWhileBoundary(text));
}

function isExecutableIfBoundary(text: string): boolean {
  return /^Else(?:If\b|$)/i.test(text) || /^End\s+If\b/i.test(text);
}

function isExecutableSelectBoundary(text: string): boolean {
  return /^Case\b/i.test(text) || /^End\s+Select\b/i.test(text);
}

function isExecutableForBoundary(text: string): boolean {
  return /^Next\b/i.test(text);
}

function isExecutableDoBoundary(text: string): boolean {
  return /^Loop\b/i.test(text);
}

function isExecutableWhileBoundary(text: string): boolean {
  return /^Wend\b/i.test(text);
}

function isExecutableWithBoundary(text: string): boolean {
  return /^End\s+With\b/i.test(text);
}

function stripLeadingLabel(text: string): string {
  return text.replace(/^(?:[A-Za-z_][A-Za-z0-9_]*|\d+):\s*/u, "").trim();
}
