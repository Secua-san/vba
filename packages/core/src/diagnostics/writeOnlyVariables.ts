import { analyzeProcedureLocalUsage } from "./localVariableUsage";
import type { Diagnostic, ParseResult } from "../types/model";

export function collectWriteOnlyVariableDiagnostics(parseResult: ParseResult): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];

  for (const member of parseResult.module.members) {
    if (member.kind !== "procedureDeclaration") {
      continue;
    }

    const usage = analyzeProcedureLocalUsage(member);

    diagnostics.push(
      ...usage.declarations
        .filter(
          (entry) =>
            entry.declarationKind === "variable" &&
            usage.writtenNames.has(entry.normalizedName) &&
            !usage.readNames.has(entry.normalizedName)
        )
        .map((entry) => ({
          code: "write-only-variable",
          message: `Write-only local variable '${entry.name}'.`,
          range: entry.range,
          severity: "warning" as const
        }))
    );
  }

  return diagnostics;
}
