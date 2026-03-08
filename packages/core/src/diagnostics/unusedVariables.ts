import { removeStringAndDateLiterals, splitCodeAndComment } from "../parser/text";
import { normalizeIdentifier } from "../types/helpers";
import type { Diagnostic, ParameterNode, ParseResult, ProcedureDeclarationNode, VariableDeclaratorNode } from "../types/model";

interface LocalDeclarationEntry {
  name: string;
  normalizedName: string;
  range: Diagnostic["range"];
}

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
  const declarations = new Map<string, LocalDeclarationEntry>();
  const usedNames = new Set<string>();

  for (const parameter of procedure.parameters) {
    addDeclaration(declarations, parameter.name, parameter.range);
  }

  for (const statement of procedure.body) {
    if (statement.declaredVariables) {
      for (const variable of statement.declaredVariables) {
        addDeclaration(declarations, variable.name, variable.range);
      }
    }

    if (statement.kind !== "executableStatement") {
      continue;
    }

    const { code } = splitCodeAndComment(statement.text);
    const scrubbedText = removeStringAndDateLiterals(stripLeadingLabel(code));

    for (const match of scrubbedText.matchAll(/[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?/g)) {
      usedNames.add(normalizeIdentifier(match[0].replace(/[$%&!#@]$/, "")));
    }
  }

  return [...declarations.values()]
    .filter((entry) => !usedNames.has(entry.normalizedName))
    .map((entry) => ({
      code: "unused-variable",
      message: `Unused local declaration '${entry.name}'.`,
      range: entry.range,
      severity: "warning" as const
    }));
}

function addDeclaration(
  declarations: Map<string, LocalDeclarationEntry>,
  name: ParameterNode["name"] | VariableDeclaratorNode["name"],
  range: Diagnostic["range"]
): void {
  const normalizedName = normalizeIdentifier(name);

  if (!declarations.has(normalizedName)) {
    declarations.set(normalizedName, {
      name,
      normalizedName,
      range
    });
  }
}

function stripLeadingLabel(text: string): string {
  return text.replace(/^\s*(?:[A-Za-z_][A-Za-z0-9_]*|\d+):\s*/u, "");
}
