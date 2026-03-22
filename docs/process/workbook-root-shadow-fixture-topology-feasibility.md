# workbook root family の shadow fixture topology 整理メモ

## 目的

`Application.ThisWorkbook` / `Application.ActiveWorkbook` の shadow hover / signature を、将来 shared case spec へ寄せるならどの fixture topology にそろえるべきかを整理する。  
このメモの焦点は test fixture の構成であり、resolver や shared spec schema を今すぐ変えることではない。

## 状況更新

- 2026-03-22 に候補 2 を採用し、extension 側の shadow section は [ApplicationWorkbookRootShadowed.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootShadowed.bas) へ分離済み
- server 側は [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の inline shadow text を維持している
- shadow hover / signature の direct anchor は extension / server とも `occurrenceIndex = 0` で参照できるようになり、shared case spec 化を完了した

## 実装前の比較

### extension 側

- fixture は [packages/extension/test/fixtures/ApplicationWorkbookRootBuiltIn.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootBuiltIn.bas) の mixed 構成だった
- `Demo()` の中に current-bundle / active-workbook / non-target root がまとまっていた
- `ShadowedApplication()` が同じ module の後半にあり、`Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value` と `Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value` は `Demo()` と重複していた
- そのため shadow hover の direct `.Value` anchor は `occurrenceIndex = 1` が必要だった
- 一方、shadow signature の direct `.Select(` anchor は `ShadowedApplication()` 側にしか無く、`occurrenceIndex = 0` で足りていた

### server 側

- shadow test は [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) 内の inline text で持っている
- `Demo()` には shadow 用 token しか入っておらず、hover / signature とも `occurrenceIndex = 0`
- server では workbook root family matrix の多くを inline text で組み立てており、fixture file を必須にしていない

### 当時の方針

- 正本は [workbook-root-family-case-table-policy.md](workbook-root-family-case-table-policy.md)
- shadow hover / signature は package-local のまま残し、dedicated shadow fixture 分離後に再評価する
- v1 では shared spec schema に `occurrenceIndexByScope` / `occurrenceIndexByKind` を追加しない

## 候補

### 1. shared spec 側に per-scope occurrence override を足す

案:

- `occurrenceIndexByScope.extension = 1`
- `occurrenceIndexByScope.server-application-shadowed = 0`
- あるいは hover / signature ごとに別 override を持つ

問題:

- `test-support/` が単純な正本 table ではなく小さな DSL になり、review しにくい
- shadow 以外の family でも同じ仕組みを使いたくなり、schema が膨らみやすい
- fixture topology の差を schema で吸収してしまい、将来 topology をそろえる動機が弱くなる

結論:

- 不採用

### 2. extension 側の shadow section を専用 fixture へ分離する

案:

- `ApplicationWorkbookRootBuiltIn.bas` から `ShadowedApplication()` を切り出し、shadow 専用 fixture を別 file にする
- extension test は shadow 用 document を別に開く
- server 側は現状どおり inline text を維持するか、後で同じ shadow fixture text へ寄せる

利点:

- hover / signature の direct anchor が extension / server とも `occurrenceIndex = 0` にそろいやすい
- shared spec に override を足さずに shadow hover / signature shared 化を再検討できる
- `Demo()` と `ShadowedApplication()` の責務が分かれ、fixture の読みやすさも上がる

懸念:

- extension test で document を 1 枚増やす必要がある
- shadow completion / semantic をどこまで同じ fixture へ寄せるかは別途決める必要がある

結論:

- shadow shared 化を再開するなら第一候補

### 3. server 側を extension と同じ mixed fixture へ寄せる

案:

- server の inline text に non-shadow section も混ぜ、extension と同じ duplicate anchor 構造を作る

利点:

- file 構成を増やさずに「同じ topology」にできる

懸念:

- server test が意図的に duplicate anchor を持つようになり、現在より読みにくい
- shadow 専用テストで見たい論点に対して、non-shadow token がノイズになる
- inline text の軽さと局所性を捨てる割に、shared spec 側の簡潔さはほとんど増えない

結論:

- 優先しない

## 判断

- 当時の第一候補だった「extension 側 shadow section の専用 fixture 分離」を採用し、現在は実装済み
- server 側は inline text 維持のままでよいという判断も継続した
- canonical shadow text source は、shared shadow spec 導入後も drift が問題化した時点で別タスクとして判断する

## 先に満たす条件

- shadow hover / signature を shared 化する具体的な review コスト、または duplicate anchor 起因の保守コストが見えていること
- extension 側で shadow 専用 fixture を追加しても test 導線が過度に複雑化しないこと
- shared 化対象を hover / signature だけにするのか、completion / semantic まで含めるのかを先に切り分けること

## 次の見直し候補

1. server / extension 間で shadow text source の drift が出るかを観測する
2. drift が問題化した場合だけ canonical shadow text source を検討する
