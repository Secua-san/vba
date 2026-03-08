export { analyzeModule, findDefinition, getCompletionSymbols, getDocumentOutline } from "./diagnostics/analyzeModule";
export { lexDocument } from "./lexer/lexDocument";
export { parseModule } from "./parser/parseModule";
export { buildModuleSymbols, getAccessibleSymbolsAtLine } from "./symbol/buildModuleSymbols";
export type {
  AnalysisResult,
  AnalyzeModuleOptions,
  Diagnostic,
  LinePosition,
  ModuleNode,
  OutlineSymbol,
  ParseResult,
  SourceRange,
  SymbolInfo,
  SymbolTable,
  Token,
  TokenKind
} from "./types/model";
