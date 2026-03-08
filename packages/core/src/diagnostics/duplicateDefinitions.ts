import { normalizeIdentifier } from "../types/helpers";
import type {
  ConstDeclarationNode,
  Diagnostic,
  ParseResult,
  SourceRange,
  VariableDeclaratorNode
} from "../types/model";

interface DefinitionEntry {
  name: string;
  range: SourceRange;
}

export function collectDuplicateDefinitionDiagnostics(parseResult: ParseResult): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];
  const moduleDefinitions: DefinitionEntry[] = [];

  for (const member of parseResult.module.members) {
    switch (member.kind) {
      case "constDeclaration":
      case "declareStatement":
      case "enumDeclaration":
      case "procedureDeclaration":
      case "typeDeclaration":
        moduleDefinitions.push({
          name: member.name,
          range: member.kind === "procedureDeclaration" ? member.headerRange : member.range
        });
        break;
      case "variableDeclaration":
        moduleDefinitions.push(...member.declarators.map(toDefinitionEntry));
        break;
      default:
        break;
    }
  }

  diagnostics.push(...collectScopeDuplicates(moduleDefinitions, "module scope"));

  for (const member of parseResult.module.members) {
    if (member.kind !== "procedureDeclaration") {
      continue;
    }

    const procedureDefinitions: DefinitionEntry[] = member.parameters.map((parameter) => ({
      name: parameter.name,
      range: parameter.range
    }));

    if (member.procedureKind !== "Sub") {
      procedureDefinitions.push({
        name: member.name,
        range: member.headerRange
      });
    }

    for (const statement of member.body) {
      if (statement.declaredVariables) {
        procedureDefinitions.push(...statement.declaredVariables.map(toDefinitionEntry));
      }

      if (statement.declaredConstants) {
        procedureDefinitions.push(...statement.declaredConstants.map((constant) => ({
          name: constant.name,
          range: constant.range
        })));
      }
    }

    diagnostics.push(...collectScopeDuplicates(procedureDefinitions, `procedure '${member.name}'`));
  }

  return diagnostics;
}

function collectScopeDuplicates(entries: DefinitionEntry[], scopeLabel: string): Diagnostic[] {
  const diagnostics: Diagnostic[] = [];
  const seenDefinitions = new Map<string, DefinitionEntry>();

  for (const entry of entries) {
    const normalizedName = normalizeIdentifier(entry.name);

    if (!seenDefinitions.has(normalizedName)) {
      seenDefinitions.set(normalizedName, entry);
      continue;
    }

    diagnostics.push({
      code: "duplicate-definition",
      message: `Duplicate definition '${entry.name}' in ${scopeLabel}.`,
      range: entry.range,
      severity: "error"
    });
  }

  return diagnostics;
}

function toDefinitionEntry(entry: ConstDeclarationNode | VariableDeclaratorNode): DefinitionEntry {
  return {
    name: entry.name,
    range: entry.range
  };
}
