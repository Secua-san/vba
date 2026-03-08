export { analyzeModule, findDefinition, getCompletionSymbols, getDocumentOutline } from "./diagnostics/analyzeModule";
export { areTypesCompatible, getSymbolTypeName, inferExpressionTypeAtLine, inferModuleTypes } from "./inference/inferModuleTypes";
export { lexDocument } from "./lexer/lexDocument";
export { parseModule } from "./parser/parseModule";
export { extractIdentifierAtPosition, removeStringAndDateLiterals, splitCodeAndComment } from "./parser/text";
export { buildModuleSymbols, getAccessibleSymbolsAtLine } from "./symbol/buildModuleSymbols";
export { normalizeIdentifier } from "./types/helpers";
export type {
  AnalysisResult,
  AnalyzeModuleOptions,
  Diagnostic,
  InferredSymbolType,
  LinePosition,
  ModuleNode,
  OutlineSymbol,
  ParseResult,
  SourceRange,
  SymbolInfo,
  SymbolTable,
  Token,
  TokenKind,
  TypeInferenceResult,
  TypeInferenceSource
} from "./types/model";
