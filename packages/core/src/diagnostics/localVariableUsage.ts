import { removeStringAndDateLiterals, splitCodeAndComment } from "../parser/text";
import { normalizeIdentifier } from "../types/helpers";
import type { Diagnostic, ProcedureDeclarationNode } from "../types/model";

export interface LocalDeclarationEntry {
  declarationKind: "parameter" | "variable";
  name: string;
  normalizedName: string;
  range: Diagnostic["range"];
}

export interface LocalVariableUsageAnalysis {
  declarations: LocalDeclarationEntry[];
  readNames: Set<string>;
  writtenNames: Set<string>;
}

export function analyzeProcedureLocalUsage(procedure: ProcedureDeclarationNode): LocalVariableUsageAnalysis {
  const declarations = new Map<string, LocalDeclarationEntry>();
  const readNames = new Set<string>();
  const writtenNames = new Set<string>();

  for (const parameter of procedure.parameters) {
    addDeclaration(declarations, "parameter", parameter.name, parameter.range);
  }

  for (const statement of procedure.body) {
    if (statement.declaredVariables) {
      for (const variable of statement.declaredVariables) {
        addDeclaration(declarations, "variable", variable.name, variable.range);
      }
    }

    if (statement.kind !== "executableStatement") {
      continue;
    }

    const { code } = splitCodeAndComment(statement.text);
    const scrubbedText = removeStringAndDateLiterals(stripLeadingLabel(code));
    const assignment = parseAssignment(scrubbedText);

    if (assignment?.targetName && declarations.has(assignment.targetName)) {
      writtenNames.add(assignment.targetName);
    }

    collectReads(assignment?.expressionText ?? scrubbedText, declarations, readNames);
  }

  return {
    declarations: [...declarations.values()],
    readNames,
    writtenNames
  };
}

function addDeclaration(
  declarations: Map<string, LocalDeclarationEntry>,
  declarationKind: LocalDeclarationEntry["declarationKind"],
  name: string,
  range: Diagnostic["range"]
): void {
  const normalizedName = normalizeIdentifier(name);

  if (!declarations.has(normalizedName)) {
    declarations.set(normalizedName, {
      declarationKind,
      name,
      normalizedName,
      range
    });
  }
}

function collectReads(
  text: string,
  declarations: Map<string, LocalDeclarationEntry>,
  readNames: Set<string>
): void {
  for (const match of text.matchAll(/[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?/g)) {
    const rawIdentifier = match[0];
    const normalizedName = normalizeIdentifier(rawIdentifier.replace(/[$%&!#@]$/, ""));
    const startIndex = match.index ?? 0;
    const previousCharacter = text[startIndex - 1] ?? "";

    if (previousCharacter === "." || previousCharacter === ":") {
      continue;
    }

    if (declarations.has(normalizedName)) {
      readNames.add(normalizedName);
    }
  }
}

function parseAssignment(text: string): { expressionText: string; targetName?: string } | undefined {
  const equalsIndex = findAssignmentOperatorIndex(text);

  if (equalsIndex < 0) {
    return undefined;
  }

  const leftText = text.slice(0, equalsIndex);
  const rightText = text.slice(equalsIndex + 1);
  const match = /^\s*(?:(?:Set|Let)\s+)?([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)(?:\s*\(.*\))?\s*$/iu.exec(leftText);

  return {
    expressionText: rightText.trim(),
    targetName: match?.[1] ? normalizeIdentifier(match[1].replace(/[$%&!#@]$/, "")) : undefined
  };
}

function findAssignmentOperatorIndex(text: string): number {
  let depth = 0;
  let index = 0;

  while (index < text.length) {
    const currentCharacter = text[index];

    if (currentCharacter === "(") {
      depth += 1;
      index += 1;
      continue;
    }

    if (currentCharacter === ")") {
      depth = Math.max(0, depth - 1);
      index += 1;
      continue;
    }

    if (depth === 0 && currentCharacter === "=") {
      const previousNonWhitespaceCharacter = getAdjacentNonWhitespaceCharacter(text, index, -1);

      if (previousNonWhitespaceCharacter !== "<" && previousNonWhitespaceCharacter !== ">") {
        return index;
      }
    }

    index += 1;
  }

  return -1;
}

function getAdjacentNonWhitespaceCharacter(text: string, startIndex: number, direction: -1 | 1): string | undefined {
  let index = startIndex + direction;

  while (index >= 0 && index < text.length) {
    if (!/\s/u.test(text[index] ?? "")) {
      return text[index];
    }

    index += direction;
  }

  return undefined;
}

function stripLeadingLabel(text: string): string {
  return text.replace(/^\s*(?:[A-Za-z_][A-Za-z0-9_]*|\d+):\s*/u, "");
}
