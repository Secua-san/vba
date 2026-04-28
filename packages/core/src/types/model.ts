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
  | "lineContinuation"
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
  valueRange?: SourceRange;
  valueText?: string;
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
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface DeclarationStatementNode {
  declaredVariables: VariableDeclaratorNode[];
  kind: "declarationStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface ProcedureStatementLabelNode {
  kind: "procedureStatementLabel";
  range: SourceRange;
  text: string;
}

export interface AssignmentStatementNode {
  assignmentKind: "implicit" | "let" | "set";
  expressionRange: SourceRange;
  expressionText: string;
  kind: "assignmentStatement";
  leadingLabel?: ProcedureStatementLabelNode;
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
  leadingLabel?: ProcedureStatementLabelNode;
  name: string;
  nameRange: SourceRange;
  range: SourceRange;
  text: string;
}

export interface IfBlockStatementNode {
  conditionRange: SourceRange;
  conditionText: string;
  kind: "ifBlockStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface ElseIfClauseStatementNode {
  conditionRange: SourceRange;
  conditionText: string;
  kind: "elseIfClauseStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface ElseClauseStatementNode {
  kind: "elseClauseStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface EndIfStatementNode {
  kind: "endIfStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface SelectCaseStatementNode {
  expressionRange: SourceRange;
  expressionText: string;
  kind: "selectCaseStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface CaseClauseStatementNode {
  caseKind: "else" | "value";
  conditionRange?: SourceRange;
  conditionText?: string;
  kind: "caseClauseStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface EndSelectStatementNode {
  kind: "endSelectStatement";
  leadingLabel?: ProcedureStatementLabelNode;
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
  leadingLabel?: ProcedureStatementLabelNode;
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
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface NextStatementNode {
  counterName?: string;
  counterRange?: SourceRange;
  counterText?: string;
  kind: "nextStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface DoBlockStatementNode {
  clauseKind: "none" | "until" | "while";
  conditionRange?: SourceRange;
  conditionText?: string;
  kind: "doBlockStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface LoopStatementNode {
  clauseKind: "none" | "until" | "while";
  conditionRange?: SourceRange;
  conditionText?: string;
  kind: "loopStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface WhileStatementNode {
  conditionRange: SourceRange;
  conditionText: string;
  kind: "whileStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface WendStatementNode {
  kind: "wendStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface WithBlockStatementNode {
  targetRange: SourceRange;
  targetText: string;
  kind: "withBlockStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface EndWithStatementNode {
  kind: "endWithStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface OnErrorStatementNode {
  actionKind: "goto" | "resumeNext";
  kind: "onErrorStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  targetRange?: SourceRange;
  targetText?: string;
  text: string;
}

export interface GoToStatementNode {
  actionKind: "goSub" | "goTo";
  kind: "goToStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  targetRange: SourceRange;
  targetText: string;
  text: string;
}

export interface ResumeStatementNode {
  actionKind: "implicit" | "next" | "target";
  kind: "resumeStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  targetRange?: SourceRange;
  targetText?: string;
  text: string;
}

export interface ExitStatementNode {
  exitKind: "Function" | "Property" | "Sub";
  kind: "exitStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface EndStatementNode {
  kind: "endStatement";
  leadingLabel?: ProcedureStatementLabelNode;
  range: SourceRange;
  text: string;
}

export interface ExecutableStatementNode {
  kind: "executableStatement";
  leadingLabel?: ProcedureStatementLabelNode;
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
