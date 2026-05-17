# Incremental Analysis Feasibility

## 結論

- 現行の解析は、open / change / restore のたびに document 全文を `DocumentService.analyzeText()` へ渡す全体再解析である。
- server は `TextDocumentSyncKind.Full` を宣言しており、LSP change event から差分 range を解析 pipeline へ渡していない。
- parser / symbol / type inference / diagnostics は `AnalysisResult` を丸ごと作り直す前提でつながっている。
- 既存 cache は document state、workspace index、workbook binding manifest、worksheet control metadata、active workbook identity snapshot の境界であり、token / AST / symbol / type の差分 reuse 境界はまだ無い。
- PR8 では product code を変更せず、差分更新解析の実装は別 roadmap に切り出す。

## 現行 call path

1. `packages/server/src/index.ts` は LSP capability として `TextDocumentSyncKind.Full` を返す。
2. `documents.onDidOpen()` は即時、`documents.onDidChangeContent()` は `analysisDebounceMs` 後に `analyzeAndPublish()` を呼ぶ。
3. `analyzeAndPublish()` は `document.getText()` の全文を `documentService.analyzeText(document.uri, document.languageId, document.version, document.getText())` へ渡す。
4. workspace file restore も file 全文を読み、`documentService.analyzeText(uri, "vba", 0, text)` を呼ぶ。
5. `DocumentService.analyzeText()` は `analyzeModule(text, { fileName })` を呼び、返った `AnalysisResult` を `DocumentState.analysis` として保存する。
6. `analyzeModule()` は `parseModule(text)`、`buildModuleSymbols(parseResult)`、`inferModuleTypes(parseResult, symbols)`、各 diagnostics collector を順に実行する。
7. `parseModule()` は `createSourceDocument(text)` と `lexPreparedDocument(source)` で全文 source / token を作り、`parsePreparedModule(source, tokens)` で module AST を作る。

この経路では、LSP の `version` は `DocumentState.version` に保持されるが、同一 document の前回 `AnalysisResult` との差分判定や partial reuse には使われていない。

## 現行 cache 境界

### DocumentState

- `createDocumentService()` は `documentStates: Map<string, DocumentState>` を URI ごとに持つ。
- `DocumentState` は `analysis`、`text`、`version`、workbook binding manifest state、worksheet control metadata state、active workbook identity state をまとめる最新 snapshot である。
- `analyzeText()` は毎回新しい `AnalysisResult` を作り、URI の `DocumentState` を置き換える。

### Workspace index

- `workspaceIndex` は `createWorkspaceIndex([...documentStates.values()])` で `DocumentState` 群から作る。
- `analyzeText()` と `remove()` は workspace index を全 document state から再構築する。
- definition / reference / diagnostics filtering / completion / semantic token は `documentStates` と `workspaceIndex` を読むため、1 document の差分解析でも cross-document symbol view の invalidation が必要になる。

### File artifact cache

- `workbookBindingManifestCache` は manifest path を key にし、`mtimeMs` と `size` が一致すると cached state を返す。
- `worksheetControlMetadataCache` は sidecar path を key にし、`mtimeMs` と `size` が一致すると cached state を返す。
- `setWorkspaceRoots()` はこの 2 つの cache を clear し、既存 `DocumentState` の manifest / sidecar state を再解決する。
- これらは file artifact の read cache であり、VBA source の token / AST / symbol / type cache ではない。

### Runtime snapshot cache

- `activeWorkbookIdentityState` は extension から受けた runtime snapshot の最新値を持つ。
- `setActiveWorkbookIdentitySnapshot()` は既存 `DocumentState` の active workbook identity と binding state を更新するが、VBA source を再 parse しない。
- この cache は workbook root gating 用であり、差分 source 解析の reuse 境界ではない。

## 差分更新が小修正で済まない理由

- `SourceDocument` は全文から `originalLines`、`normalizedLines`、`lineMap`、`normalizedText` を作る。`.frm` では code 開始行検出も全文前提である。
- lexer は `normalizedLines` 全体を走査し、各 token に絶対 line / character range を付ける。
- parser は logical line 全体から module member と procedure body を作り、node range は全文中の絶対位置を持つ。
- symbol table は module members と procedure body 全体から作られ、procedure scope は range と body symbols に依存する。
- type inference は procedure body の assignment を走査し、explicit type seed と assignment order に依存する。
- diagnostics、semantic token、reference、rename、completion は `AnalysisResult` 全体と workspace index を読む。

したがって、単に changed line だけを parse して差し替えると、range remap、procedure boundary、module-level symbol、type inference、workspace index の整合が崩れる可能性がある。

## 実装前の設計ゲート

差分更新解析を実装する前に、少なくとも以下を別 roadmap で決める。

1. 測定基準
   - どの module size / edit pattern で全体再解析が問題になるか。
   - server 側で計測する場合の log level、個人コード断片を出さない測定項目、既定無効 / 有効の条件。
2. invalidation 単位
   - line、logical line、procedure、module member のどれを最小 dirty unit にするか。
   - `.frm` の code start 変化、line continuation、`:` statement separator、conditional directive をどう扱うか。
3. range / identity contract
   - node / symbol に stable identity を持たせるか。
   - unchanged node の range remap をどの層が担うか。
4. cache boundary
   - token cache、logical line cache、AST cache、symbol cache、type cache をどこまで分けるか。
   - `DocumentState.analysis` の public shape を維持するか、内部 cache を別 object として持つか。
5. fallback policy
   - dirty range が procedure boundary や module-level declaration に触れた場合は全体再解析へ戻す条件。
   - parse error がある document で partial reuse を許すか。
6. workspace invalidation
   - module-level export が変わったときに workspace index と他 document diagnostics をどう再計算するか。
   - procedure body だけの変更で workspace index rebuild を省ける条件。
7. validation
   - full parse と incremental parse の `AnalysisResult` 等価性を比較する fixture 方針。
   - parser / symbol / type / server test をどこまで分けるか。

## 推奨 roadmap

1. 計測のみの PR
   - behavior を変えず、全体再解析の時間と document size を opt-in log で測る。
2. parser 境界設計 PR
   - logical line / procedure boundary / range remap の contract を docs または ADR に固定する。
3. cache prototype PR
   - product path へ入れる前に、full parse と同じ結果になるかを core test で比較する小さい prototype に限定する。
4. server integration PR
   - `TextDocumentSyncKind.Incremental` へ切り替えるか、Full sync のまま内部 cache だけ持つかを決めてから実装する。
5. workspace invalidation PR
   - symbol export と diagnostics filtering の再計算条件を固定する。

## 非目標

- PR8 で parser、symbol、type inference、diagnostics、server LSP runtime を変更しない。
- token / AST / symbol / type cache を推測で追加しない。
- `TextDocumentSyncKind.Incremental` へ切り替えない。
- workspace index の rebuild 方針を product code で変更しない。
- file artifact cache を source analysis cache と混同しない。

## 次段候補

- 大きい `.bas` / `.cls` / `.frm` fixture に対する現行全体再解析時間を opt-in に測る。
- parser の logical line と procedure boundary を、incremental reuse 可能な単位として扱えるかだけを設計する。
- `AnalysisResult` の public shape を変えずに内部 cache を持てるか、または別 API が必要かを ADR で決める。
