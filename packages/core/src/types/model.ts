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

export interface ConstStatementNode {
  declaredConstants: ConstDeclarationNode[];
  kind: "constStatement";
  range: SourceRange;
  text: string;
}

export interface DeclarationStatementNode {
  declaredVariables: VariableDeclaratorNode[];
  kind: "declarationStatement";
  range: SourceRange;
  text: string;
}

export interface AssignmentStatementNode {
  assignmentKind: "implicit" | "let" | "set";
  expressionRange: SourceRange;
  expressionText: string;
  kind: "assignmentStatement";
  range: SourceRange;
  targetName?: string;
  targetRange: SourceRange;
  targetText: string;
  text: string;
}

export interface InvocationArgumentNode {
  range: SourceRange;
  text: string;
}

export interface CallStatementNode {
  arguments: InvocationArgumentNode[];
  callStyle: "bare" | "call" | "parenthesized";
  kind: "callStatement";
  name: string;
  nameRange: SourceRange;
  range: SourceRange;
  text: string;
}

export interface IfBlockStatementNode {
  conditionRange: SourceRange;
  conditionText: string;
  kind: "ifBlockStatement";
  range: SourceRange;
  text: string;
}

export interface ElseIfClauseStatementNode {
  conditionRange: SourceRange;
  conditionText: string;
  kind: "elseIfClauseStatement";
  range: SourceRange;
  text: string;
}

export interface ElseClauseStatementNode {
  kind: "elseClauseStatement";
  range: SourceRange;
  text: string;
}

export interface EndIfStatementNode {
  kind: "endIfStatement";
  range: SourceRange;
  text: string;
}

export interface SelectCaseStatementNode {
  expressionRange: SourceRange;
  expressionText: string;
  kind: "selectCaseStatement";
  range: SourceRange;
  text: string;
}

export interface CaseClauseStatementNode {
  caseKind: "else" | "value";
  conditionRange?: SourceRange;
  conditionText?: string;
  kind: "caseClauseStatement";
  range: SourceRange;
  text: string;
}

export interface EndSelectStatementNode {
  kind: "endSelectStatement";
  range: SourceRange;
  text: string;
}

export interface ForStatementNode {
  counterName?: string;
  counterRange: SourceRange;
  counterText: string;
  endExpressionRange: SourceRange;
  endExpressionText: string;
  kind: "forStatement";
  range: SourceRange;
  startExpressionRange: SourceRange;
  startExpressionText: string;
  stepExpressionRange?: SourceRange;
  stepExpressionText?: string;
  text: string;
}

export interface ForEachStatementNode {
  collectionRange: SourceRange;
  collectionText: string;
  itemName?: string;
  itemRange: SourceRange;
  itemText: string;
  kind: "forEachStatement";
  range: SourceRange;
  text: string;
}

export interface NextStatementNode {
  counterName?: string;
  counterRange?: SourceRange;
  counterText?: string;
  kind: "nextStatement";
  range: SourceRange;
  text: string;
}

export interface DoBlockStatementNode {
  clauseKind: "none" | "until" | "while";
  conditionRange?: SourceRange;
  conditionText?: string;
  kind: "doBlockStatement";
  range: SourceRange;
  text: string;
}

export interface LoopStatementNode {
  clauseKind: "none" | "until" | "while";
  conditionRange?: SourceRange;
  conditionText?: string;
  kind: "loopStatement";
  range: SourceRange;
  text: string;
}

export interface WhileStatementNode {
  conditionRange: SourceRange;
  conditionText: string;
  kind: "whileStatement";
  range: SourceRange;
  text: string;
}

export interface WendStatementNode {
  kind: "wendStatement";
  range: SourceRange;
  text: string;
}

export interface WithBlockStatementNode {
  targetRange: SourceRange;
  targetText: string;
  kind: "withBlockStatement";
  range: SourceRange;
  text: string;
}

export interface EndWithStatementNode {
  kind: "endWithStatement";
  range: SourceRange;
  text: string;
}

export interface OnErrorStatementNode {
  actionKind: "goto" | "resumeNext";
  kind: "onErrorStatement";
  range: SourceRange;
  targetRange?: SourceRange;
  targetText?: string;
  text: string;
}

export interface GoToStatementNode {
  actionKind: "goSub" | "goTo";
  kind: "goToStatement";
  range: SourceRange;
  targetRange: SourceRange;
  targetText: string;
  text: string;
}

export interface ResumeStatementNode {
  actionKind: "implicit" | "next" | "target";
  kind: "resumeStatement";
  range: SourceRange;
  targetRange?: SourceRange;
  targetText?: string;
  text: string;
}

export interface ExitStatementNode {
  exitKind: "Function" | "Property" | "Sub";
  kind: "exitStatement";
  range: SourceRange;
  text: string;
}

export interface EndStatementNode {
  kind: "endStatement";
  range: SourceRange;
  text: string;
}

export interface ExecutableStatementNode {
  kind: "executableStatement";
  range: SourceRange;
  text: string;
}

export type ProcedureStatementNode =
  | AssignmentStatementNode
  | CaseClauseStatementNode
  | CallStatementNode
  | ConstStatementNode
  | DeclarationStatementNode
  | DoBlockStatementNode
  | EndWithStatementNode
  | ElseClauseStatementNode
  | ElseIfClauseStatementNode
  | EndIfStatementNode
  | EndSelectStatementNode
  | EndStatementNode
  | ExecutableStatementNode
  | ExitStatementNode
  | ForEachStatementNode
  | ForStatementNode
  | GoToStatementNode
  | IfBlockStatementNode
  | LoopStatementNode
  | NextStatementNode
  | OnErrorStatementNode
  | ResumeStatementNode
  | SelectCaseStatementNode
  | WhileStatementNode
  | WendStatementNode
  | WithBlockStatementNode;

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
