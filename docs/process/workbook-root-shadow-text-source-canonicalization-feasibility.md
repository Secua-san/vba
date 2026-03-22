# workbook root family の shadow text source canonical 化要否メモ

## 目的

`ApplicationWorkbookRootShadowed.bas` と server 側 inline shadow text を、今の段階で 1 つの canonical text source へ寄せるべきかを整理する。  
焦点は「shared anchor spec とは別に、shadow 用の全文テキスト正本が必要かどうか」であり、resolver や case spec schema を広げることではない。

## 現状

- extension 側の shadow case は [ApplicationWorkbookRootShadowed.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootShadowed.bas) に分離済み
- server 側の shadow case は [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の inline text を維持している
- completion / hover / signature / semantic の canonical anchor は [test-support/workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) に寄せている
- shared spec で管理していないのは、server の inline text 全体、extension fixture の module header / local declaration、package-local の failure message と wait 条件だけ

## 観測

### 1. 直近の review コストの主因は text drift ではなかった

shadow fixture 分離以降に出た review 指摘は、主に次の 2 系統だった。

- duplicate anchor / `occurrenceIndex` / `state: "shadowed"` の扱い
- extension 側 helper の待機条件や message builder の brittle さ

いずれも shared anchor spec と test helper の問題であり、server inline text と extension fixture の全文が別管理であること自体は主因ではなかった。

### 2. すでに canonical 化されている対象は「全文」ではなく「検証 anchor」

現在 drift が問題になるのは、server / extension で参照する anchor や state がずれる場合である。  
その論点はすでに [test-support/workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) が正本になっており、全文テキストまで 1 本化しなくても review 上の説明責任を満たせている。

### 3. full text canonical 化には別のコストがある

server 側を extension fixture file 読み込みへ寄せる、または generator で shadow text を組み立てる場合、次のコストが増える。

- server test の局所性が落ち、fixture から必要な数行だけ読む現在の読みやすさが下がる
- helper / loader / build 前処理が増え、docs-only ではなく test infrastructure の変更になる
- 「anchor drift を防ぐ」以上の効果が薄い割に、failure の切り分け地点が増える

## 比較

### 1. file 正本へ寄せる

案:

- server も `ApplicationWorkbookRootShadowed.bas` を読み、shadow text を fixture file 起点で共有する

利点:

- shadow 用全文は 1 つにできる
- line / token の参照元が明示的になる

欠点:

- server test が fixture file layout に引きずられ、inline text の局所性を失う
- server 固有の最小 text 構成を取りにくくなる
- 現在の drift 主因ではない箇所へ I/O と loader を持ち込むことになる

### 2. generator 正本へ寄せる

案:

- `test-support/` に shadow text builder を置き、server / extension が同じ string template を使う

利点:

- file / inline の違いを吸収しつつ全文を 1 本化できる

欠点:

- `test-support/` が case spec だけでなく text DSL まで抱える
- review 時に「実際の VBA text」が見えにくくなる
- 今の drift 問題に対しては過剰設計

### 3. 現状維持

案:

- shadow 用全文は extension fixture と server inline text の 2 系統を維持する
- canonical 化対象は引き続き shared anchor spec に限定する

利点:

- 現在の review コストに対して最小で済む
- server / extension それぞれの test 読みやすさを維持できる
- drift が起きても、まず shared anchor spec / state / scope の問題か、全文 source の問題かを切り分けやすい

欠点:

- shadow text 自体は二重管理のまま残る
- 将来 shadow section が拡大した場合は、同じ修正を 2 箇所に入れる可能性がある

## 判断

- 現時点では full text の canonical shadow text source は不要
- canonical 正本は引き続き `test-support/workbookRootFamilyCaseTables.cjs` の anchor / state / scope に留める
- server inline shadow text と extension dedicated shadow fixture の二重管理は、そのまま維持する

理由:

- 直近の review / 修正コストは anchor spec と helper 側に集中しており、全文 text source の二重管理は主因ではなかった
- server test の inline 局所性を崩すコストが、現時点の drift 防止効果を上回る
- workbook root family では dedicated shadow fixture 分離後に `occurrenceIndex` override も不要になっており、いま追加で canonical 化を進める根拠が弱い

## 再評価トリガー

- shadow 用 anchor 追加のたびに server inline text と extension fixture の両方で手修正が必要になり、2 PR 以上連続で review 指摘の主因になったとき
- shared anchor spec は一致しているのに、server / extension の全文 text 差分だけが failure や review 指摘の主因になったとき
- workbook root family 以外でも同種の dedicated shadow fixture + inline shadow text の二重管理が 2 family 以上に増え、共通運用ルールが必要になったとき

## 次の見直し候補

1. workbook root family と同種の dedicated shadow fixture 分離が、別 family でも必要になるかを観測する
2. shadow text drift が実際に review コストの主因になった時だけ、file 正本と generator 正本を再比較する
