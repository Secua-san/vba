# application workbook root family の extension-only interaction / semantic entry を server scope へ広げない判断

## 結論

- 現時点では `applicationWorkbookRoot` の extension-only `hover` / `signature` / `semantic` entry を server scope へ広げない。
- `test-support/workbookRootFamilyCaseTables.cjs` に `server-application-ole` / `server-application-shape` / `server-application-shadowed` を追加するのは見送る。
- server 側は引き続き、completion negative の mirror、positive / closed / shadowed の interaction coverage、semantic token の conservative coverage を維持し、残る user-facing residual slice は extension E2E に委ねる。

## 目的

前段で `applicationWorkbookRoot.completion.negative` の extension-only 3 entry は server scope へ広げ、completion の route-specific gap を閉じた。  
このメモでは、その後に残る extension-only `hover` / `signature` / `semantic` entry について、同じように server へ mirror する価値がまだあるかを再観測する。

## 現状

### 残っている extension-only entry

- `hover.negative`
  - `Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Valu`
  - `Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu`
- `signature.negative`
  - `Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Select(`
  - `Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(`
  - `Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(`
- `semantic.negative`
  - `Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value`
  - `Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value`
  - いずれも `reason: "shadowed-root"` / `state: "shadowed"`

### server 側で既にある coverage

- [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) は `server-application-ole` / `server-application-shape` scope で completion negative を static / matched 両 state に対して持つ。
- 同じ test 群で、`hover` / `signature` / `semantic` の positive と negative を `server-application-ole` / `server-application-shape` から読んでいる。
- `server-application-shadowed` scope では、shadowed root に対する closed completion と no-hover / no-signature を already covered として持つ。

### extension 側の役割

- [packages/extension/test/suite/index.ts](../../packages/extension/test/suite/index.ts) は shared spec の `scope: "extension"` entry を end-to-end に実行し、hover / signature wait 条件、semantic token provider 経由の最終表示まで確認している。
- `ApplicationWorkbookRootBuiltIn.bas` と `ApplicationWorkbookRootShadowed.bas` は `applicationWorkbookRoot` family の canonical anchor source であり、local helper の都合で追加された anchor ではない。

## 観察結果

### 1. completion で閉じたのは route-specific gap であり、残差は API surface の重複寄りである

- completion を server へ広げた 3 entry は、`Application.ThisWorkbook.Worksheets(1)` や `Application.ThisWorkbook.Worksheets.Item("Sheet1")` のように、server が同じ prefix / route をまだ completion として踏んでいない穴だった。
- 一方で今回残っている interaction entry は、同じ reason の completion negative や semantic negative を server が既に持っている。
- つまり今の残差は「resolver の未踏破」よりも「別 API でも同じ closed 結果を重ねるか」の論点に寄っている。

### 2. hover / signature の残差は `ThisWorkbook` static negative の一部だけで、server で増える説明力が小さい

- `hover.negative` の extension-only 2 本は code-name selector の OLE / Shape。
- `signature.negative` の extension-only 3 本は上記に numeric selector の OLE を足しただけで、いずれも `ThisWorkbook` static negative に閉じている。
- これらは server 側で既に
  - completion negative による root / selector gating
  - semantic negative による conservative token
  - positive anchor に対する hover / signature 成立
  を確認済みで、追加しても「その anchor でも出なかった」を重ねる比重が大きい。

### 3. semantic の extension-only 残差は shadowed-root だけで、branch 追加ではなく surface 追加になる

- `semantic.negative` で extension-only のまま残っているのは shadowed root 2 本だけである。
- server 側は [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) で `server-application-shadowed` scope の closed completion と no-hover / no-signature を既に検証している。
- ここへ no-semantic まで追加しても、`Application` shadowing による built-in root suppression という同じ分岐結果を別 surface で重ねるだけで、route 切り分けの粒度は大きく増えない。

### 4. extension E2E の方が user-facing residual slice を自然に観測できる

- hover / signature は editor command と async wait を含み、semantic は VS Code provider 経由で最終 token を得る。
- これらの残差は server unit test で mirror するより、extension E2E で「実際に表示されない」ことを確認する方が user-facing な観測として素直である。
- package-local adapter に残している wait 条件や token decode は、まさにこの残差を extension 側に置く理由になっている。

### 5. 今 server scope を足すと、shared spec の entry 数だけが増えやすい

- 追加対象は 7 entry で、どれも既存の reason / state 語彙に乗っているため shared spec へは載せやすい。
- ただしその増分は、現時点では `packages/server/test/documentService.test.js` の説明力を大きく増やさず、review では `completion` のときより「なぜ増やしたか」を説明しにくい。
- したがって、completion と同じ基準で mirror を広げず、ここで止める方が shared spec の密度を保ちやすい。

## 判断

### 今回は server scope を追加しない

- `applicationWorkbookRoot.hover.negative` / `signature.negative` / `semantic.negative` に新しい server scope は追加しない。
- OLE / Shape / shadowed の各 server test は、今ある completion / interaction / semantic coverage を維持し、extension-only residual slice は extension E2E の shared entry として残す。

### completion と同じ扱いにはしない

- completion は route-specific gap を閉じるために server mirror を増やした。
- interaction / semantic の残差は、その gap が埋まった後の surface duplication 寄りなので、同じ基準で mirror 拡張しない。
- したがって「shared spec に残す」と「server でも mirror する」を再び分離して扱う。

### 再評価の条件

- `applicationWorkbookRoot` の hover / signature / semantic で、completion では検知できない route-specific regression が server 側で実際に起きたとき
- `server-application-shadowed` に no-semantic を足さないことで原因切り分けが弱いと、review か不具合解析で繰り返し分かったとき
- workbook root family 全体で「どの surface なら server へ mirror するか」の基準を共通 policy として持った方が説明しやすい状況になったとき

## 推奨方針

### 維持するもの

- completion negative の server mirror
- `server-application-ole` / `server-application-shape` の positive / negative semantic coverage
- `server-application-shadowed` の closed completion と no-hover / no-signature coverage
- extension 側の `hover` / `signature` / `semantic` residual slice

### 今やらないこと

- `hover` / `signature` の static negative へ `server-application-ole` / `server-application-shape` を足す
- shadowed semantic negative へ `server-application-shadowed` を足す
- extension-only residual slice を shared spec から外して local table に戻す

## 関連文書

- 直前の判断: [application-workbook-root-completion-server-scope-feasibility.md](./application-workbook-root-completion-server-scope-feasibility.md)
- shared spec 境界: [application-workbook-root-extension-only-interaction-shared-spec-feasibility.md](./application-workbook-root-extension-only-interaction-shared-spec-feasibility.md)
- 共通 policy: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
