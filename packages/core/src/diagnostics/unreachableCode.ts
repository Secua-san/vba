import type { Diagnostic, ParseResult, ProcedureDeclarationNode, ProcedureKind } from "../types/model";
import { hasStatementSeparatorColon } from "../parser/text";

type BlockKind = "do" | "for" | "if" | "select" | "while" | "with";
type BarrierKind = "do" | "for" | "if" | "procedure" | "select" | "while";

interface UnreachableState {
  barrierKind: BarrierKind;
  reason: string;
}

interface StatementControlMetadata {
  boundaryKind?: BlockKind;
  pop?: BlockKind;
  push?: BlockKind;
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
  const transition = getStatementControlMetadata(statement, text);

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
  return getStatementBoundaryKind(statement, text) === unreachableState.barrierKind;
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

  if (getStatementBoundaryKind(statement, text)) {
    return false;
  }

  return true;
}

function getStatementControlMetadata(
  statement: ProcedureDeclarationNode["body"][number],
  text: string
): StatementControlMetadata | undefined {
  return getStructuredControlMetadata(statement) ?? getExecutableControlMetadata(statement, text);
}

function getStatementBoundaryKind(
  statement: ProcedureDeclarationNode["body"][number],
  text: string
): BlockKind | undefined {
  return getStatementControlMetadata(statement, text)?.boundaryKind;
}

function getStructuredControlMetadata(statement: ProcedureDeclarationNode["body"][number]): StatementControlMetadata | undefined {
  switch (statement.kind) {
    case "ifBlockStatement":
      return { push: "if" };
    case "elseIfClauseStatement":
    case "elseClauseStatement":
      return { boundaryKind: "if" };
    case "selectCaseStatement":
      return { push: "select" };
    case "caseClauseStatement":
      return { boundaryKind: "select" };
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
      return { boundaryKind: "if", pop: "if" };
    case "endSelectStatement":
      return { boundaryKind: "select", pop: "select" };
    case "nextStatement":
      return { boundaryKind: "for", pop: "for" };
    case "loopStatement":
      return { boundaryKind: "do", pop: "do" };
    case "wendStatement":
      return { boundaryKind: "while", pop: "while" };
    case "endWithStatement":
      return { boundaryKind: "with", pop: "with" };
    default:
      return undefined;
  }
}

function getExecutableControlMetadata(
  statement: ProcedureDeclarationNode["body"][number],
  text: string
): StatementControlMetadata | undefined {
  if (statement.kind !== "executableStatement") {
    return undefined;
  }

  if (isExecutableIfBlockStart(text)) {
    return { push: "if" };
  }

  if (/^Else(?:If\b|$)/i.test(text)) {
    return { boundaryKind: "if" };
  }

  if (/^Select\s+Case\b/i.test(text)) {
    return { push: "select" };
  }

  if (/^Case\b/i.test(text)) {
    return { boundaryKind: "select" };
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
    return { boundaryKind: "if", pop: "if" };
  }

  if (/^End\s+Select\b/i.test(text)) {
    return { boundaryKind: "select", pop: "select" };
  }

  if (/^Next\b/i.test(text)) {
    return { boundaryKind: "for", pop: "for" };
  }

  if (/^Loop\b/i.test(text)) {
    return { boundaryKind: "do", pop: "do" };
  }

  if (/^Wend\b/i.test(text)) {
    return { boundaryKind: "while", pop: "while" };
  }

  if (/^End\s+With\b/i.test(text)) {
    return { boundaryKind: "with", pop: "with" };
  }

  return undefined;
}

function isExecutableIfBlockStart(text: string): boolean {
  return !/^ElseIf\b/i.test(text) && /^If\b.*\bThen\s*$/i.test(text) && !hasStatementSeparatorColon(text);
}

function stripLeadingLabel(text: string): string {
  return text.replace(/^(?:[A-Za-z_][A-Za-z0-9_]*|\d+):\s*/u, "").trim();
}
