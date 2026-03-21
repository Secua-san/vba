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

package-local adapter に残す対象:

- server 側の `symbol.name`
- extension 側の `CompletionItem.detail` fragment
- package ごとの failure message
- decoded token / hover / signature help の最終 assertion shape

### 4. duplicated string のうち、shared 化対象は「fixture と 1 対 1 に対応する anchor」だけにする

`Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value` のような fixture 上の anchor token は shared 化する価値が高い。  
一方で「hover を出さない」「snapshot 一致後も閉じる」などの文章メッセージは package ごとに持つ。

### 5. shadow / duplicate occurrence は canonical spec に必ず明示する

今回の CodeRabbit 指摘の通り、anchor token ベースへ寄せると `Demo()` と `ShadowedApplication()` のような重複文字列で `occurrenceIndex` 抜けが起きやすい。  
そのため shared spec へ切り出すときは、`occurrenceIndex` を optional 扱いにせず、重複しうる anchor は明示指定を必須にする。

## やらないこと

- server / extension の helper 関数そのものを 1 つへ統合する
- workbook root family 以外の built-in test まで同時に DSL 化する
- failure message を shared spec へ押し込み、review 時の読みやすさを落とす
- JSON 化のためだけに build step や copy step を追加する

## 次の最小実装単位

1. `test-support/workbookRootFamilyCaseTables.cjs` を追加する
2. 対象は `ApplicationWorkbookRootBuiltIn.bas` と `WorksheetBroadRootBuiltIn.bas` の workbook root family matrix に限定する
3. 最初は semantic token と completion の anchor spec だけを shared 化し、hover / signature help は adapter の読みやすさを見て追随させる
4. server / extension の各 test は shared spec を読み、package-local helper へ変換する薄い adapter を持つ

## 受け入れ条件

- workbook root family の fixture anchor が server / extension で二重管理されない
- `CompletionItem.detail` や async wait 条件のような package 固有事情は shared spec に漏れ出さない
- review 時に「どの anchor を shared 正本で持ち、どの期待値を package-local で持つか」が 1 画面で追える
