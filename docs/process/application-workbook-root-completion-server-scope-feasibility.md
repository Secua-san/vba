# application workbook root family の extension-only completion negative を server scope へ広げる要否

## 結論

- `applicationWorkbookRoot.completion.negative` に残っている extension-only 3 entry は、server scope へ広げる価値がある。
- 結論は「shared spec には残す」ではなく、そのうえで `server-application-ole` / `server-application-shape` の mirror も足す、である。
- 次段は docs ではなく最小実装として、shared case table の `scopes` 更新と server fixture text への anchor 追加を行う。

## 目的

前段で、extension-only completion negative 3 entry は shared spec の canonical anchor source に残す判断を固定した。  
このメモでは、その 3 entry について server unit test でも mirror する価値があるかを整理する。

## 対象

現在 `applicationWorkbookRoot.completion.negative` のうち scope が `["extension"]` だけの entry は次の 3 本である。

- `Application.ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.`
  - `reason: "numeric-selector"`
  - `state: "static"`
- `Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.`
  - `reason: "code-name-selector"`
  - `state: "static"`
- `Application.ActiveWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.`
  - `reason: "numeric-selector"`
  - `state: "matched"`

## 観察結果

### 1. 3 本とも completion anchor 自体は fixture 上に存在し、追加コストが低い

- [ApplicationWorkbookRootBuiltIn.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootBuiltIn.bas) には、上記 3 本に対応する anchor が既に存在する。
- server 側で未 mirror なのは resolver や fixture が不足しているからではなく、inline text 側に同じ completion anchor 行をまだ置いていないためである。
- したがって、server scope 拡張は shared case table の `scopes` 更新と、server fixture text への 3 行追加で済む。

### 2. 3 本のうち 2 本は、同じ prefix を server がまだ completion / interaction / semantic のどこでも踏めていない

- `Application.ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.`
- `Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.`

この 2 本は、server 側では

- direct-route numeric selector の `ThisWorkbook + OLE`
- item-route code-name selector の `ThisWorkbook + Shape`

という組み合わせを同じ prefix でまだ踏めていない。  
extension E2E にだけ残しておくと、route-specific な回帰が server unit test では止まらず、VS Code host 経由まで落ちてこない。

### 3. 残り 1 本は同 prefix の hover / semantic が server にあるが、completion negative としても mirror 価値がある

- `Application.ActiveWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.`

この anchor prefix は server 側に hover / semantic negative が既にある。  
ただし completion negative は

- member list が開かないこと
- broad root gating が completion surface でも閉じること

を直接見る slice であり、hover / semantic の mirror だけでは代替できない。  
completion は `DocumentService.getCompletionSymbols()` の閉じ方そのものを見るため、completion negative を server に足す意味は残る。

### 4. completion negative は server mirror しやすく、adapter 依存が少ない

- extension completion の固有事情は `CompletionItem.detail` fragment と blocked label だが、server 側は `symbol.name` が出ないことだけを見ればよい。
- つまり server mirror で増えるのは host 依存ではなく、`DocumentService` の completion resolver への純粋な coverage である。
- 既存 helper の枠内で収まり、shared spec schema や package-local adapter 境界を増やさない。

### 5. residual slice は 3 本に閉じており、review 負荷より coverage 利得が勝つ

- もし extension-only residual が十数本規模なら、server mirror 追加は matrix の読みにくさを増やしやすい。
- 現状は 3 本だけで、しかも route / reason の穴が明確である。
- この規模なら review 負荷は小さく、shared spec の scope 非対称をさらに減らせる利得の方が大きい。

## 判断

### server scope へ広げる

- 上記 3 本は server scope へ広げる。
- OLE 側 1 本には `server-application-ole` を追加し、Shape 側 2 本には `server-application-shape` を追加する。
- これにより `applicationWorkbookRoot.completion.negative` の completion residual slice を減らす。

### ただし shared spec / adapter の境界は変えない

- shared spec に残すのは引き続き `anchor` / `reason` / `state` / `scopes` までとする。
- `CompletionItem.detail` fragment と blocked label は extension adapter に残す。
- つまり今回広げるのは scope だけであり、schema や helper 契約は変えない。

### 実装は最小に留める

- 新しい helper は追加しない。
- server fixture text に必要な 3 anchor 行だけを足し、shared case table の scope を更新する。
- docs 側ではこの判断を正本として残し、次タスクで実装する。

## 再評価の条件

- 3 本の mirror 追加後も、completion negative の route-specific gap が同 family 内に複数残るとき
- server 側で completion negative を増やした結果、helper や message の分岐が膨らみ、1 画面で主語を追いにくくなったとき
- extension-only completion residual が再び増え、server mirror の費用対効果が悪化したとき

## 推奨方針

### 次段でやること

- `test-support/workbookRootFamilyCaseTables.cjs` の 3 entry に server scope を追加する
- [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の inline text に 3 anchor 行を追加する
- `npm run test --workspace @vba/server` と関連 lint を通し、completion negative の server mirror が閉じ方を壊していないことを確認する

### 今やらないこと

- shared spec schema に completion 専用 field を追加する
- extension completion adapter の `detailFragment` / blocked label を server 側へ持ち込む
- hover / signature / semantic の別残件まで同時に広げる

## 関連文書

- 直前の判断: [application-workbook-root-extension-only-completion-shared-spec-feasibility.md](./application-workbook-root-extension-only-completion-shared-spec-feasibility.md)
- 共通 policy: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
- interaction / semantic 側の判断: [application-workbook-root-extension-only-interaction-shared-spec-feasibility.md](./application-workbook-root-extension-only-interaction-shared-spec-feasibility.md)
