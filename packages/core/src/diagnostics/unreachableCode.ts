import type { Diagnostic, ParseResult, ProcedureDeclarationNode, ProcedureKind } from "../types/model";

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

    if (unreachableState && clearsUnreachableState(controlText, unreachableState)) {
      unreachableState = undefined;
    }

    if (unreachableState && shouldReportUnreachableStatement(controlText)) {
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

    applyBlockTransition(controlText, blockStack);
  }

  return diagnostics;
}

function applyBlockTransition(text: string, blockStack: BlockKind[]): void {
  if (isIfBlockStart(text)) {
    blockStack.push("if");
    return;
  }

  if (/^Select\s+Case\b/i.test(text)) {
    blockStack.push("select");
    return;
  }

  if (/^For\b/i.test(text)) {
    blockStack.push("for");
    return;
  }

  if (/^Do\b/i.test(text)) {
    blockStack.push("do");
    return;
  }

  if (/^While\b/i.test(text)) {
    blockStack.push("while");
    return;
  }

  if (/^With\b/i.test(text)) {
    blockStack.push("with");
    return;
  }

  if (/^End\s+If\b/i.test(text)) {
    popLastBlockOfKind(blockStack, "if");
    return;
  }

  if (/^End\s+Select\b/i.test(text)) {
    popLastBlockOfKind(blockStack, "select");
    return;
  }

  if (/^Next\b/i.test(text)) {
    popLastBlockOfKind(blockStack, "for");
    return;
  }

  if (/^Loop\b/i.test(text)) {
    popLastBlockOfKind(blockStack, "do");
    return;
  }

  if (/^Wend\b/i.test(text)) {
    popLastBlockOfKind(blockStack, "while");
    return;
  }

  if (/^End\s+With\b/i.test(text)) {
    popLastBlockOfKind(blockStack, "with");
  }
}

function clearsUnreachableState(text: string, unreachableState: UnreachableState): boolean {
  switch (unreachableState.barrierKind) {
    case "if":
      return /^Else(?:If\b|$)/i.test(text) || /^End\s+If\b/i.test(text);
    case "select":
      return /^Case\b/i.test(text) || /^End\s+Select\b/i.test(text);
    case "for":
      return /^Next\b/i.test(text);
    case "do":
      return /^Loop\b/i.test(text);
    case "while":
      return /^Wend\b/i.test(text);
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
  return !/^ElseIf\b/i.test(text) && /^If\b.*\bThen\s*$/i.test(text) && !/:/.test(text);
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

function shouldReportUnreachableStatement(text: string): boolean {
  if (text.length === 0) {
    return false;
  }

  return !(
    /^Else(?:If\b|$)/i.test(text) ||
    /^Case\b/i.test(text) ||
    /^End\s+If\b/i.test(text) ||
    /^End\s+Select\b/i.test(text) ||
    /^Next\b/i.test(text) ||
    /^Loop\b/i.test(text) ||
    /^Wend\b/i.test(text) ||
    /^End\s+With\b/i.test(text)
  );
}

function stripLeadingLabel(text: string): string {
  return text.replace(/^(?:[A-Za-z_][A-Za-z0-9_]*|\d+):\s*/u, "").trim();
}
