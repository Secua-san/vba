import { analyzeProcedureLocalUsage } from "./localVariableUsage";
import type { Diagnostic, ParseResult, ProcedureDeclarationNode } from "../types/model";

export function collectUnusedVariableDiagnostics(parseResult: ParseResult): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];

  for (const member of parseResult.module.members) {
    if (member.kind !== "procedureDeclaration") {
      continue;
    }

    diagnostics.push(...collectProcedureUnusedVariableDiagnostics(member));
  }

  return diagnostics;
}

function collectProcedureUnusedVariableDiagnostics(procedure: ProcedureDeclarationNode): Diagnostic[] {
  const usage = analyzeProcedureLocalUsage(procedure);

  return usage.declarations
    .filter((entry) => !usage.readNames.has(entry.normalizedName) && !usage.writtenNames.has(entry.normalizedName))
    .map((entry) => ({
      code: "unused-variable",
      message: `Unused local declaration '${entry.name}'.`,
      range: entry.range,
      severity: "warning" as const
    }));
}
