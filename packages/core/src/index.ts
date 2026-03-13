export { analyzeModule, findDefinition, getCompletionSymbols, getDocumentOutline } from "./diagnostics/analyzeModule";
export { collectByRefArgumentDiagnostics } from "./diagnostics/byRefDiagnostics";
export type { ResolvedCallable } from "./diagnostics/byRefDiagnostics";
export { formatModuleIndentation } from "./format/formatModuleIndentation";
export type { FormatModuleIndentationOptions } from "./format/formatModuleIndentation";
export { areTypesCompatible, getSymbolTypeName, inferExpressionTypeAtLine, inferModuleTypes } from "./inference/inferModuleTypes";
export { lexDocument } from "./lexer/lexDocument";
export { parseModule } from "./parser/parseModule";
export { extractIdentifierAtPosition, removeStringAndDateLiterals, splitCodeAndComment } from "./parser/text";
export {
  BUILTIN_IDENTIFIERS,
  BUILTIN_REFERENCE_ITEMS,
  getBuiltinCompletionItems,
  getBuiltinMemberCompletionItems,
  getBuiltinMemberReferenceItem,
  getBuiltinMemberSignature,
  getBuiltinReferenceItem,
  isReservedOrBuiltinIdentifier,
  resolveBuiltinMemberOwner,
  stripIndexedAccessMarker,
  RESERVED_IDENTIFIERS
} from "./reference/builtinReference";
export type {
  BuiltinCallableSignature,
  BuiltinCompletionKind,
  BuiltinMemberKind,
  BuiltinMemberReferenceItem,
  BuiltinReferenceItem,
  BuiltinSemanticModifier,
  BuiltinSemanticType,
  BuiltinSignatureParameter
} from "./reference/builtinReference";
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
