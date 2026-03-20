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
  getIndexedAccessKind,
  getBuiltinCompletionItems,
  getBuiltinMemberCompletionItems,
  getBuiltinMemberReferenceItem,
  getBuiltinMemberSignature,
  getBuiltinReferenceItem,
  isReservedOrBuiltinIdentifier,
  markIndexedAccessPathSegment,
  resolveBuiltinMemberOwner,
  resolveBuiltinMemberOwnerFromRootType,
  stripIndexedAccessMarker,
  RESERVED_IDENTIFIERS
} from "./reference/builtinReference";
export type {
  BuiltinCallableSignature,
  BuiltinCompletionKind,
  BuiltinMemberKind,
  BuiltinMemberReferenceItem,
  BuiltinReferenceItem,
  IndexedAccessKind,
  BuiltinSemanticModifier,
  BuiltinSemanticType,
  BuiltinSignatureParameter
} from "./reference/builtinReference";
export {
  ACTIVE_WORKBOOK_IDENTITY_NOTIFICATION_METHOD,
  ACTIVE_WORKBOOK_IDENTITY_TEST_STATE_REQUEST_METHOD,
  ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND,
  ACTIVE_WORKBOOK_IDENTITY_VERSION,
  normalizeWorkbookFullNameForComparison,
  parseActiveWorkbookIdentitySnapshot
} from "./reference/activeWorkbookIdentity";
export type {
  ActiveWorkbookIdentityAvailableSnapshot,
  ActiveWorkbookIdentityFields,
  ActiveWorkbookIdentityParseResult,
  ActiveWorkbookIdentityProtectedViewSnapshot,
  ActiveWorkbookIdentitySnapshot,
  ActiveWorkbookIdentityUnavailableReason,
  ActiveWorkbookIdentityUnavailableSnapshot,
  ActiveWorkbookIdentityUnsupportedReason,
  ActiveWorkbookIdentityUnsupportedSnapshot,
  ActiveWorkbookIdentityValidationIssue,
  ActiveWorkbookProtectedViewFields
} from "./reference/activeWorkbookIdentity";
export {
  buildWorksheetControlMetadataSidecarPath,
  findNearestWorksheetControlMetadataSidecar,
  getSupportedWorksheetControlMetadataOwners,
  parseWorksheetControlMetadataSidecar,
  WORKSHEET_CONTROL_METADATA_SIDECAR_ARTIFACT,
  WORKSHEET_CONTROL_METADATA_SIDECAR_DIRECTORY_NAME,
  WORKSHEET_CONTROL_METADATA_SIDECAR_FILE_NAME,
  WORKSHEET_CONTROL_METADATA_SIDECAR_VERSION
} from "./reference/worksheetControlMetadataSidecar";
export type {
  WorksheetControlMetadataSidecar,
  WorksheetControlMetadataSidecarControl,
  WorksheetControlMetadataSidecarLocation,
  WorksheetControlMetadataSidecarLookupOptions,
  WorksheetControlMetadataSidecarOwner,
  WorksheetControlMetadataSidecarParseResult,
  WorksheetControlMetadataSidecarWorkbook,
  WorksheetControlMetadataSupportedOwner,
  WorksheetControlMetadataUnsupportedOwner,
  WorksheetControlMetadataValidationIssue
} from "./reference/worksheetControlMetadataSidecar";
export {
  buildWorkbookBindingManifestPath,
  findNearestWorkbookBindingManifest,
  parseWorkbookBindingManifest,
  WORKBOOK_BINDING_MANIFEST_ARTIFACT,
  WORKBOOK_BINDING_MANIFEST_BINDING_KIND,
  WORKBOOK_BINDING_MANIFEST_DIRECTORY_NAME,
  WORKBOOK_BINDING_MANIFEST_FILE_NAME,
  WORKBOOK_BINDING_MANIFEST_VERSION
} from "./reference/workbookBindingManifest";
export type {
  WorkbookBindingManifest,
  WorkbookBindingManifestLocation,
  WorkbookBindingManifestLookupOptions,
  WorkbookBindingManifestParseResult,
  WorkbookBindingManifestValidationIssue,
  WorkbookBindingManifestWorkbook
} from "./reference/workbookBindingManifest";
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
