import type { ProcedureStatementNode, SourceRange } from "../types/model";

export type ProcedureStatementReferenceRole = "read" | "readWrite" | "write";

export interface ProcedureStatementReferenceSegment {
  identifierName?: string;
  range: SourceRange;
  role: ProcedureStatementReferenceRole;
  text: string;
}

export function getProcedureStatementReferenceSegments(
  statement: ProcedureStatementNode
): ProcedureStatementReferenceSegment[] | undefined {
  switch (statement.kind) {
    case "assignmentStatement":
      return [
        {
          identifierName: statement.targetName ?? extractAssignableIdentifierName(statement.targetText),
          range: statement.targetRange,
          role: "write",
          text: statement.targetText
        },
        {
          range: statement.expressionRange,
          role: "read",
          text: statement.expressionText
        }
      ];
    case "callStatement":
      return statement.arguments.map((argument) => ({
        range: argument.range,
        role: "read",
        text: argument.text
      }));
    case "ifBlockStatement":
    case "elseIfClauseStatement":
      return [
        {
          range: statement.conditionRange,
          role: "read",
          text: statement.conditionText
        }
      ];
    case "selectCaseStatement":
      return [
        {
          range: statement.expressionRange,
          role: "read",
          text: statement.expressionText
        }
      ];
    case "caseClauseStatement":
      return statement.caseKind === "value" && statement.conditionRange && statement.conditionText
        ? [
            {
              range: statement.conditionRange,
              role: "read",
              text: statement.conditionText
            }
          ]
        : [];
    case "forStatement":
      return [
        {
          identifierName: statement.counterName ?? extractAssignableIdentifierName(statement.counterText),
          range: statement.counterRange,
          role: "readWrite",
          text: statement.counterText
        },
        {
          range: statement.startExpressionRange,
          role: "read",
          text: statement.startExpressionText
        },
        {
          range: statement.endExpressionRange,
          role: "read",
          text: statement.endExpressionText
        },
        ...(statement.stepExpressionRange && statement.stepExpressionText
          ? [
              {
                range: statement.stepExpressionRange,
                role: "read",
                text: statement.stepExpressionText
              } satisfies ProcedureStatementReferenceSegment
            ]
          : [])
      ];
    case "forEachStatement":
      return [
        {
          identifierName: statement.itemName ?? extractAssignableIdentifierName(statement.itemText),
          range: statement.itemRange,
          role: "readWrite",
          text: statement.itemText
        },
        {
          range: statement.collectionRange,
          role: "read",
          text: statement.collectionText
        }
      ];
    case "nextStatement":
      return statement.counterRange && statement.counterText
        ? [
            {
              identifierName: statement.counterName ?? extractAssignableIdentifierName(statement.counterText),
              range: statement.counterRange,
              role: "read",
              text: statement.counterText
            }
          ]
        : [];
    case "doBlockStatement":
    case "loopStatement":
      return statement.conditionRange && statement.conditionText
        ? [
            {
              range: statement.conditionRange,
              role: "read",
              text: statement.conditionText
            }
          ]
        : [];
    case "whileStatement":
      return [
        {
          range: statement.conditionRange,
          role: "read",
          text: statement.conditionText
        }
      ];
    case "withBlockStatement":
      return [
        {
          range: statement.targetRange,
          role: "read",
          text: statement.targetText
        }
      ];
    case "constStatement":
    case "declarationStatement":
    case "elseClauseStatement":
    case "endIfStatement":
    case "endSelectStatement":
    case "endStatement":
    case "endWithStatement":
    case "exitStatement":
    case "goToStatement":
    case "onErrorStatement":
    case "resumeStatement":
    case "wendStatement":
      return [];
    case "executableStatement":
      return undefined;
  }
}

function extractAssignableIdentifierName(text: string): string | undefined {
  const match = /^\s*([A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)(?:\s*\(.*\))?\s*$/u.exec(text);
  return match?.[1] ? match[1].replace(/[$%&!#@]$/, "") : undefined;
}
