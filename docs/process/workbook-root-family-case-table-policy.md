# workbook root family の matrix case table 共通化方針

## 目的

`packages/server/test/documentService.test.js` と `packages/extension/test/suite/index.ts` にある workbook root family 向け matrix は、helper までは寄ったが case table 自体はまだ重複している。  
このメモでは、server の同期 assertion と extension の非同期 assertion を壊さずに、どこまで shared 化するかを固定する。

## 現状の制約

- server test は CommonJS + Node 標準テストで、`DocumentService` を直接たたく同期 assertion が中心
- extension test は TypeScript + VS Code host 経由で、completion / hover / signature help / semantic token を非同期に待つ
- completion case は server 側が `symbol.name` と blocked symbol を見る一方、extension 側は `CompletionItem.detail` も見る
- semantic token は anchor token と `occurrenceIndex` が共有論点だが、decoded token の shape は package ごとに異なる
- fixture は共通でも、triage しやすい failure message は package ごとに持っていた方が読みやすい

## 判断

### 1. assertion helper は package-local のまま維持する

`assertWorkbookRootCompletionCases()` などの helper は、server と extension で同期 / 非同期、入力型、diagnostics の出し方が違うため共有しない。  
共通化対象は helper 関数ではなく、helper に食わせる matrix の正本だけに限定する。

### 2. canonical な case spec は repo root の test support へ寄せる

実装するときの正本は、repo root 配下の `test-support/` に置く。  
候補は `test-support/workbookRootFamilyCaseTables.cjs` とし、Node / CommonJS からそのまま `require()` できる形式を優先する。

この置き方にする理由:

- server test から相対 import しやすい
- extension test も compiled JS から absolute path `require()` で読める
- `resolveJsonModule` や test build への copy step を追加せずに済む
- TypeScript 専用 module にすると、extension 側だけ都合がよく server 側の読み方が崩れやすい

### 3. shared 化するのは anchor / state / expectation kind までに留める

共通 spec に入れる対象:

- fixture 名
- matrix family 名
- state
  - `static`
  - `matched`
  - `closed`
  - `shadowed`
- route
  - `ole-object`
  - `shape-oleformat`
- anchor token
- identifier
- `occurrenceIndex`
- expectation kind
  - completion positive / negative
  - hover positive / negative
  - signature positive / negative
  - semantic positive / negative

現時点の適用範囲:

- `WorksheetBroadRootBuiltIn.bas` / `ApplicationWorkbookRootBuiltIn.bas` / `ApplicationWorkbookRootShadowed.bas` の completion / hover / signature / semantic は shared 化済み
- `ApplicationWorkbookRootShadowed.bas` と server 側 inline shadow fixture は direct anchor topology が 1 対 1 にそろっており、shadow hover / signature も `occurrenceIndex = 0` の shared entry を使う
- package-local に残すのは、`CompletionItem.detail` fragment、package ごとの failure message、decoded token / hover / signature help の最終 assertion shapeだけに留める

package-local adapter に残す対象:

- server 側の `symbol.name`
- extension 側の `CompletionItem.detail` fragment
- package ごとの failure message
- decoded token / hover / signature help の最終 assertion shape

### 4. duplicated string のうち、shared 化対象は「fixture と 1 対 1 に対応する anchor」だけにする

`Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value` のような fixture 上の anchor token は shared 化する価値が高い。  
一方で「hover を出さない」「snapshot 一致後も閉じる」などの文章メッセージは package ごとに持つ。

### 4.5 scope 非対称だけでは shared spec から外さない

shared spec entry は、`server` と `extension` の両方が使うことを必須条件にしない。  
`scopes` は「その entry をどの package / slice が消費するか」を表す契約であり、ある entry が `extension` だけ、または `server` だけを持っていてもよい。

shared spec に残してよい条件:

- anchor / reason / state が family canonical source として意味を持つ
- package 固有事情が helper / adapter 層に隔離されている
- fixture anchor の正本を local file へ戻すより、`test-support/` に残した方が drift を抑えやすい

shared spec から外すべき候補:

- async wait 条件や failure message のように package 固有事情が主体になっている
- canonical anchor ではなく local helper の都合でしか使わない
- scope 非対称 entry が増えすぎて、family table より local file の方が主語を追いやすい

### 5. shadow / duplicate occurrence は canonical spec に必ず明示する

anchor token ベースへ寄せると、同じ fixture 内で duplicate anchor が生じたときに `occurrenceIndex` 抜けが起きやすい。  
workbook root family の shadow case は dedicated fixture 分離で duplicate anchor を解消したが、今後も重複しうる anchor は shared spec 側で明示指定を必須にする。

### 6. per-scope occurrence override は導入しない

現時点で package ごとの occurrence 差分は解消しているが、v1 では shared spec schema に

- `occurrenceIndexByScope`
- `occurrenceIndexByKind`
- `fixtureVariant`

のような override を足さない。理由:

- 現状の workbook root family では override 無しで shared 化できており、schema 追加の必要が無い
- 将来別 family で同種のズレが出ても、まず fixture topology と anchor topology の整理で吸収できるかを先に見るべき
- `test-support/` をミニ DSL 化すると review しづらく、戻しにくい

再評価のトリガー:

- shadow hover / signature と同種の per-scope occurrence 差分が別 family でも 2 箇所以上出たとき
- server inline shadow text と extension dedicated shadow fixture の anchor drift が、shared spec 維持コストの主因になったとき
- shared spec へ残した local case が review 負荷の主因になり、schema 複雑化のコストを上回ると判断できたとき

## やらないこと

- server / extension の helper 関数そのものを 1 つへ統合する
- workbook root family 以外の built-in test まで同時に DSL 化する
- failure message を shared spec へ押し込み、review 時の読みやすさを落とす
- JSON 化のためだけに build step や copy step を追加する

## 次の見直し候補

1. workbook root family 以外の built-in family へ shared case spec を広げるときも、まず dedicated fixture と anchor topology の整理で吸収できるかを確認する
2. workbook root family の shadow text source は [workbook-root-shadow-text-source-canonicalization-feasibility.md](workbook-root-shadow-text-source-canonicalization-feasibility.md) の判断を正本とし、drift が review 主因になった場合だけ再判断する

## 受け入れ条件

- workbook root family の fixture anchor が server / extension で二重管理されない
- `CompletionItem.detail` や async wait 条件のような package 固有事情は shared spec に漏れ出さない
- review 時に「どの anchor を shared 正本で持ち、どの期待値を package-local で持つか」が 1 画面で追える
- dedicated shadow fixture 分離後も scope override 無しで shared spec を維持できる理由と、再評価トリガーがこの文書だけで説明できる
