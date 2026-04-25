import type { Diagnostic, ParseResult, ProcedureDeclarationNode, ProcedureKind } from "../types/model";
import { hasStatementSeparatorColon } from "../parser/text";

type BlockKind = "do" | "for" | "if" | "select" | "while" | "with";
type BarrierKind = "do" | "for" | "if" | "procedure" | "select" | "while";

interface UnreachableState {
  barrierKind: BarrierKind;
  barrierIndex?: number;
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

    if (hasLeadingLabel(rawText)) {
      unreachableState = undefined;
    }

    const legacyControlText = statement.kind === "executableStatement" ? stripLeadingLabel(rawText) : undefined;
    const controlMetadata = getStatementControlMetadata(statement, legacyControlText);

    if (unreachableState && clearsUnreachableState(controlMetadata, blockStack, unreachableState)) {
      unreachableState = undefined;
    }

    if (unreachableState && shouldReportUnreachableStatement(controlMetadata)) {
      diagnostics.push({
        code: "unreachable-code",
        message: `Unreachable code after ${unreachableState.reason}.`,
        range: statement.range,
        severity: "warning"
      });
    }

    const terminationReason = getTerminationReason(statement, procedure.procedureKind, legacyControlText);

    if (terminationReason) {
      const barrierKind = getBarrierKind(blockStack);
      unreachableState = {
        barrierIndex: getBarrierIndex(blockStack, barrierKind),
        barrierKind,
        reason: terminationReason
      };
    }

    applyBlockTransition(controlMetadata, blockStack);
  }

  return diagnostics;
}

function applyBlockTransition(controlMetadata: StatementControlMetadata | undefined, blockStack: BlockKind[]): void {
  if (controlMetadata?.push) {
    blockStack.push(controlMetadata.push);
    return;
  }

  if (controlMetadata?.pop) {
    popLastBlockOfKind(blockStack, controlMetadata.pop);
    return;
  }
}

function clearsUnreachableState(
  controlMetadata: StatementControlMetadata | undefined,
  blockStack: BlockKind[],
  unreachableState: UnreachableState
): boolean {
  const boundaryKind = controlMetadata?.boundaryKind;

  if (!boundaryKind || boundaryKind !== unreachableState.barrierKind) {
    return false;
  }

  return getBarrierIndex(blockStack, boundaryKind) === unreachableState.barrierIndex;
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

function getBarrierIndex(blockStack: BlockKind[], barrierKind: BlockKind | BarrierKind): number | undefined {
  if (barrierKind === "procedure") {
    return undefined;
  }

  const barrierIndex = blockStack.lastIndexOf(barrierKind);
  return barrierIndex >= 0 ? barrierIndex : undefined;
}

function hasLeadingLabel(text: string): boolean {
  return /^(?:[A-Za-z_][A-Za-z0-9_]*|\d+):/u.test(text);
}

function getTerminationReason(
  statement: ProcedureDeclarationNode["body"][number],
  procedureKind: ProcedureKind,
  legacyControlText: string | undefined
): string | undefined {
  return getStructuredTerminationReason(statement, procedureKind) ?? getTextTerminationReason(legacyControlText, procedureKind);
}

function getStructuredTerminationReason(
  statement: ProcedureDeclarationNode["body"][number],
  procedureKind: ProcedureKind
): string | undefined {
  if (statement.kind === "endStatement") {
    return "End";
  }

  if (statement.kind !== "exitStatement" || !isExitKindForProcedure(statement.exitKind, procedureKind)) {
    return undefined;
  }

  return `Exit ${statement.exitKind}`;
}

function getTextTerminationReason(text: string | undefined, procedureKind: ProcedureKind): string | undefined {
  if (!text) {
    return undefined;
  }

  const normalizedText = text.trim();

  if (/^End$/i.test(normalizedText)) {
    return "End";
  }

  const exitKind = getTextExitKind(normalizedText);

  if (!exitKind || !isExitKindForProcedure(exitKind, procedureKind)) {
    return undefined;
  }

  return `Exit ${exitKind}`;
}

function getTextExitKind(text: string): "Function" | "Property" | "Sub" | undefined {
  const match = /^Exit\s+(Function|Property|Sub)$/iu.exec(text);

  if (!match?.[1]) {
    return undefined;
  }

  const normalizedExitKind = match[1].toLowerCase();
  return normalizedExitKind === "function" ? "Function" : normalizedExitKind === "property" ? "Property" : "Sub";
}

function isExitKindForProcedure(exitKind: "Function" | "Property" | "Sub", procedureKind: ProcedureKind): boolean {
  switch (procedureKind) {
    case "Function":
      return exitKind === "Function";
    case "PropertyGet":
    case "PropertyLet":
    case "PropertySet":
      return exitKind === "Property";
    case "Sub":
      return exitKind === "Sub";
    default:
      return assertNever(procedureKind);
  }
}

function assertNever(value: never): never {
  throw new Error(`Unexpected procedure kind: ${String(value)}`);
}

function popLastBlockOfKind(blockStack: BlockKind[], blockKind: BlockKind): void {
  for (let index = blockStack.length - 1; index >= 0; index -= 1) {
    if (blockStack[index] === blockKind) {
      blockStack.splice(index, 1);
      return;
    }
  }
}

function shouldReportUnreachableStatement(controlMetadata: StatementControlMetadata | undefined): boolean {
  if (controlMetadata?.boundaryKind) {
    return false;
  }

  return true;
}

function getStatementControlMetadata(
  statement: ProcedureDeclarationNode["body"][number],
  legacyControlText: string | undefined
): StatementControlMetadata | undefined {
  return getStructuredControlMetadata(statement) ?? getExecutableControlMetadata(legacyControlText);
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

function getExecutableControlMetadata(text: string | undefined): StatementControlMetadata | undefined {
  if (!text) {
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
