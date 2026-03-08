import { normalizeIdentifier } from "../types/helpers";
import { ParseResult, ProcedureDeclarationNode, ProcedureScope, SymbolInfo, SymbolTable } from "../types/model";

export function buildModuleSymbols(parseResult: ParseResult): SymbolTable {
  const moduleSymbol: SymbolInfo = {
    kind: "module",
    name: parseResult.module.name,
    normalizedName: normalizeIdentifier(parseResult.module.name),
    range: parseResult.module.range,
    scope: "module",
    selectionRange: parseResult.module.range
  };
  const moduleSymbols: SymbolInfo[] = [];
  const procedureScopes: ProcedureScope[] = [];

  for (const member of parseResult.module.members) {
    switch (member.kind) {
      case "constDeclaration":
        moduleSymbols.push(createModuleSymbol("constant", member.name, member.range, member.range, member.typeName));
        break;
      case "declareStatement":
        moduleSymbols.push({
          kind: "declare",
          name: member.name,
          normalizedName: normalizeIdentifier(member.name),
          procedureKind: member.procedureKind,
          range: member.range,
          scope: "module",
          selectionRange: member.range,
          typeName: member.returnType
        });
        break;
      case "enumDeclaration":
        moduleSymbols.push(createModuleSymbol("enum", member.name, member.range, member.range));

        for (const enumMember of member.members) {
          moduleSymbols.push({
            containerName: member.name,
            kind: "enumMember",
            name: enumMember.name,
            normalizedName: normalizeIdentifier(enumMember.name),
            range: enumMember.range,
            scope: "module",
            selectionRange: enumMember.range
          });
        }
        break;
      case "procedureDeclaration":
        moduleSymbols.push({
          kind: "procedure",
          name: member.name,
          normalizedName: normalizeIdentifier(member.name),
          procedureKind: member.procedureKind,
          range: member.range,
          scope: "module",
          selectionRange: member.headerRange,
          typeName: member.returnType
        });
        procedureScopes.push(buildProcedureScope(member));
        break;
      case "typeDeclaration":
        moduleSymbols.push(createModuleSymbol("type", member.name, member.range, member.range));

        for (const typeMember of member.members) {
          moduleSymbols.push({
            containerName: member.name,
            kind: "typeMember",
            name: typeMember.name,
            normalizedName: normalizeIdentifier(typeMember.name),
            range: typeMember.range,
            scope: "module",
            selectionRange: typeMember.range,
            typeName: typeMember.typeName
          });
        }
        break;
      case "variableDeclaration":
        for (const declarator of member.declarators) {
          moduleSymbols.push(
            createModuleSymbol("variable", declarator.name, declarator.range, declarator.range, declarator.typeName, declarator.arraySuffix)
          );
        }
        break;
      default:
        break;
    }
  }

  return {
    allSymbols: [moduleSymbol, ...moduleSymbols, ...procedureScopes.flatMap((scope) => scope.symbols)],
    moduleName: parseResult.module.name,
    moduleSymbol,
    moduleSymbols,
    procedureScopes
  };
}

export function getAccessibleSymbolsAtLine(symbolTable: SymbolTable, line: number): SymbolInfo[] {
  const scope = symbolTable.procedureScopes.find(
    (item) => line >= item.procedure.range.start.line && line <= item.procedure.range.end.line
  );

  if (!scope) {
    return [symbolTable.moduleSymbol, ...symbolTable.moduleSymbols];
  }

  const combinedSymbols = [symbolTable.moduleSymbol, ...symbolTable.moduleSymbols, ...scope.symbols];
  const uniqueSymbols = new Map<string, SymbolInfo>();

  for (const symbol of combinedSymbols) {
    const key = `${symbol.kind}:${symbol.normalizedName}`;

    if (!uniqueSymbols.has(key)) {
      uniqueSymbols.set(key, symbol);
    }
  }

  return [...uniqueSymbols.values()];
}

function buildProcedureScope(procedure: ProcedureDeclarationNode): ProcedureScope {
  const symbols: SymbolInfo[] = procedure.parameters.map((parameter) => ({
    isArray: parameter.arraySuffix,
    kind: "parameter",
    name: parameter.name,
    normalizedName: normalizeIdentifier(parameter.name),
    range: parameter.range,
    scope: "procedure",
    selectionRange: parameter.range,
    typeName: parameter.typeName
  }));

  if (procedure.procedureKind !== "Sub") {
    symbols.push({
      kind: "variable",
      name: procedure.name,
      normalizedName: normalizeIdentifier(procedure.name),
      range: procedure.headerRange,
      scope: "procedure",
      selectionRange: procedure.headerRange,
      typeName: procedure.returnType
    });
  }

  for (const statement of procedure.body) {
    if (statement.declaredVariables) {
      for (const variable of statement.declaredVariables) {
        symbols.push({
          isArray: variable.arraySuffix,
          kind: "variable",
          name: variable.name,
          normalizedName: normalizeIdentifier(variable.name),
          range: variable.range,
          scope: "procedure",
          selectionRange: variable.range,
          typeName: variable.typeName
        });
      }
    }

    if (statement.declaredConstants) {
      for (const constant of statement.declaredConstants) {
        symbols.push({
          kind: "constant",
          name: constant.name,
          normalizedName: normalizeIdentifier(constant.name),
          range: constant.range,
          scope: "procedure",
          selectionRange: constant.range,
          typeName: constant.typeName
        });
      }
    }
  }

  return {
    procedure,
    symbols
  };
}

function createModuleSymbol(
  kind: SymbolInfo["kind"],
  name: string,
  range: SymbolInfo["range"],
  selectionRange: SymbolInfo["selectionRange"],
  typeName?: string,
  isArray?: boolean
): SymbolInfo {
  return {
    isArray,
    kind,
    name,
    normalizedName: normalizeIdentifier(name),
    range,
    scope: "module",
    selectionRange,
    typeName
  };
}
