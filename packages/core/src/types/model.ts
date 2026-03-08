export interface LinePosition {
  line: number;
  character: number;
}

export interface SourceRange {
  start: LinePosition;
  end: LinePosition;
}

export type DiagnosticSeverity = "error" | "warning";

export interface Diagnostic {
  code: string;
  message: string;
  severity: DiagnosticSeverity;
  range: SourceRange;
}

export type TokenKind =
  | "attribute"
  | "comment"
  | "dateLiteral"
  | "directive"
  | "eof"
  | "identifier"
  | "keyword"
  | "newline"
  | "numberLiteral"
  | "operator"
  | "punctuation"
  | "stringLiteral";

export interface Token {
  kind: TokenKind;
  text: string;
  range: SourceRange;
}

export type ModuleKind = "class" | "form" | "standard";

export interface AnalyzeModuleOptions {
  fileName?: string;
  moduleName?: string;
}

export interface SourceDocument {
  fileName?: string;
  lineMap: number[];
  moduleKind: ModuleKind;
  moduleName: string;
  normalizedLines: string[];
  normalizedText: string;
  originalLines: string[];
  originalText: string;
}

export interface VariableDeclaratorNode {
  arraySuffix: boolean;
  kind: "variableDeclarator";
  name: string;
  range: SourceRange;
  typeName?: string;
}

export interface ParameterNode {
  arraySuffix: boolean;
  direction: "byRef" | "byVal";
  isOptional: boolean;
  isParamArray: boolean;
  kind: "parameter";
  name: string;
  range: SourceRange;
  typeName?: string;
}

export interface AttributeLineNode {
  kind: "attributeLine";
  name: string;
  range: SourceRange;
  text: string;
  value?: string;
}

export interface OptionStatementNode {
  kind: "optionStatement";
  name: string;
  range: SourceRange;
  text: string;
}

export interface VariableDeclarationNode {
  declarators: VariableDeclaratorNode[];
  kind: "variableDeclaration";
  modifier?: string;
  range: SourceRange;
  text: string;
}

export interface ConstDeclarationNode {
  kind: "constDeclaration";
  modifier?: string;
  name: string;
  range: SourceRange;
  text: string;
  typeName?: string;
}

export interface EnumMemberNode {
  kind: "enumMember";
  name: string;
  range: SourceRange;
}

export interface EnumDeclarationNode {
  kind: "enumDeclaration";
  members: EnumMemberNode[];
  modifier?: string;
  name: string;
  range: SourceRange;
  text: string;
}

export interface TypeMemberNode {
  kind: "typeMember";
  name: string;
  range: SourceRange;
  typeName?: string;
}

export interface TypeDeclarationNode {
  kind: "typeDeclaration";
  members: TypeMemberNode[];
  modifier?: string;
  name: string;
  range: SourceRange;
  text: string;
}

export type ProcedureKind =
  | "Function"
  | "PropertyGet"
  | "PropertyLet"
  | "PropertySet"
  | "Sub";

export interface DeclareStatementNode {
  isPtrSafe: boolean;
  kind: "declareStatement";
  modifier?: string;
  name: string;
  parameters: ParameterNode[];
  procedureKind: "Function" | "Sub";
  range: SourceRange;
  returnType?: string;
  text: string;
}

export interface ProcedureStatementNode {
  declaredConstants?: ConstDeclarationNode[];
  declaredVariables?: VariableDeclaratorNode[];
  kind: "constStatement" | "declarationStatement" | "executableStatement";
  range: SourceRange;
  text: string;
}

export interface ProcedureDeclarationNode {
  body: ProcedureStatementNode[];
  headerRange: SourceRange;
  isStatic: boolean;
  kind: "procedureDeclaration";
  modifier?: string;
  name: string;
  parameters: ParameterNode[];
  procedureKind: ProcedureKind;
  range: SourceRange;
  returnType?: string;
}

export interface DirectiveNode {
  kind: "directive";
  range: SourceRange;
  text: string;
}

export interface UnknownStatementNode {
  kind: "unknownStatement";
  range: SourceRange;
  text: string;
}

export type ModuleMemberNode =
  | AttributeLineNode
  | ConstDeclarationNode
  | DeclareStatementNode
  | DirectiveNode
  | EnumDeclarationNode
  | OptionStatementNode
  | ProcedureDeclarationNode
  | TypeDeclarationNode
  | UnknownStatementNode
  | VariableDeclarationNode;

export interface ModuleNode {
  kind: "module";
  members: ModuleMemberNode[];
  name: string;
  range: SourceRange;
}

export type SymbolKind =
  | "constant"
  | "declare"
  | "enum"
  | "enumMember"
  | "module"
  | "parameter"
  | "procedure"
  | "type"
  | "typeMember"
  | "variable";

export interface SymbolInfo {
  containerName?: string;
  isArray?: boolean;
  kind: SymbolKind;
  name: string;
  normalizedName: string;
  procedureKind?: ProcedureKind;
  range: SourceRange;
  scope: "module" | "procedure";
  selectionRange: SourceRange;
  typeName?: string;
}

export interface ProcedureScope {
  procedure: ProcedureDeclarationNode;
  symbols: SymbolInfo[];
}

export interface SymbolTable {
  allSymbols: SymbolInfo[];
  moduleName: string;
  moduleSymbol: SymbolInfo;
  moduleSymbols: SymbolInfo[];
  procedureScopes: ProcedureScope[];
}

export type TypeInferenceSource = "assignment" | "explicit" | "return";

export interface InferredSymbolType {
  source: TypeInferenceSource;
  symbol: SymbolInfo;
  typeName: string;
}

export interface TypeInferenceResult {
  diagnostics: Diagnostic[];
  symbolTypes: InferredSymbolType[];
}

export interface OutlineSymbol {
  children?: OutlineSymbol[];
  kind: SymbolKind;
  name: string;
  range: SourceRange;
  selectionRange: SourceRange;
}

export interface ParseResult {
  diagnostics: Diagnostic[];
  module: ModuleNode;
  source: SourceDocument;
  tokens: Token[];
}

export interface AnalysisResult extends ParseResult {
  diagnostics: Diagnostic[];
  symbols: SymbolTable;
  typeInference: TypeInferenceResult;
}
