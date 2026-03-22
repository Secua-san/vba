# workbook root family の shadow 専用 fixture 分離 PoC 整理メモ

## 目的

`Application.ThisWorkbook` / `Application.ActiveWorkbook` の shadow case について、extension 側で `ShadowedApplication()` を別 fixture へ切り出す PoC をやる価値があるかを整理する。  
ここで決めるのは PoC の境界と影響見積もりであり、shared spec schema や server 側 canonical text source を今すぐ導入することではない。

## 現状

### extension 側

- [packages/extension/test/fixtures/ApplicationWorkbookRootBuiltIn.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootBuiltIn.bas) に `Demo()` と `ShadowedApplication()` が同居している
- [packages/extension/test/suite/index.ts](../../packages/extension/test/suite/index.ts) では `ApplicationWorkbookRootBuiltIn.bas` を 1 回だけ開き、その 1 document に対して static / matched / shadowed の completion / hover / signature / semantic token をまとめて検証している
- `Demo()` と `ShadowedApplication()` に同じ direct `.Value` anchor があるため、shadow hover は `occurrenceIndex = 1` が必要
- shadow signature の `.Select(` は現状 `ShadowedApplication()` 側にしか無く、`occurrenceIndex = 0` で足りる

### server 側

- [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の shadow case は inline text を `createWorksheetBroadRootFixture(text)` へ渡す局所テストになっている
- shadow test の text には shadow 用 anchor しか入っておらず、hover / signature は両方とも `occurrenceIndex = 0`
- server 側は workbook root family matrix のかなりの部分を inline text で持っており、fixture file 追加を前提にしていない

## PoC で見る論点

### 1. extension の document 数増加

想定案:

- `ApplicationWorkbookRootBuiltIn.bas` から `ShadowedApplication()` を抜き、たとえば `ApplicationWorkbookRootShadowed.bas` を新設する
- extension test は static / matched 用 document と shadow 用 document を別々に開く

影響:

- `openTextDocument()` / `showTextDocument()` が 1 回増える
- workbook root family helper 自体は再利用できるが、shadow 用 case 群に渡す `TextDocument` が別になる
- semantic token wait も shadow 用 document で別に 1 回走る

見積もり:

- 追加コストは「document が 1 枚増える」「shadow 系 assert 呼び出しの引数 document を差し替える」程度で、helper 全面刷新までは不要

### 2. shared spec への影響

期待できること:

- extension shadow hover の duplicate direct `.Value` anchor が消えるため、server と同じ `occurrenceIndex = 0` にそろえやすい
- shadow hover / signature を shared spec へ寄せる再検討がしやすくなる

残ること:

- completion / semantic token は static / matched / shadowed で state ごとの差分が残るため、fixture 分離だけで全面 shared 化できるわけではない
- shared spec schema を今すぐ変える必要はなく、まずは package-local shadow table のままでも PoC は成立する

判断:

- PoC の目的は「shared 化を即実施すること」ではなく、「shared 化を邪魔している topology 差を extension 側だけでどこまで減らせるか」の確認に置く

### 3. server 側 inline text の扱い

選択肢:

- A. server は inline text のまま維持する
- B. server にも canonical shadow text source を導入する

比較:

- A は現在の局所性を崩さず、PoC 変更点を extension 側に閉じ込められる
- B は将来 shared 化しやすく見えるが、file / generator / test-support のどこを正本にするか別論点を増やす

判断:

- PoC 段階では A を採る
- canonical shadow text source は、「server / extension の shadow anchor 群を 1 つの正本から生成したい」という要求が明確になってから別タスクで判断する

## PoC の最小境界

- extension だけで `ShadowedApplication()` を別 fixture へ分離する
- server 側は inline shadow text を維持する
- shared spec schema は変更しない
- shadow hover / signature の `occurrenceIndex` を extension / server とも `0` にそろえられるかを確認対象にする
- completion / semantic token の shared 化までは PoC 完了条件に含めない

## 期待する効果

- shadow hover / signature の package-local 差分が減り、[workbook-root-family-case-table-policy.md](workbook-root-family-case-table-policy.md) で残している occurrence 差分の再評価がしやすくなる
- extension fixture の責務が `static / matched` と `shadowed` に分かれ、review 時に anchor の意図を追いやすくなる

## 今は見送ること

- server 側 canonical shadow text source の導入
- `test-support/` shared spec schema への per-scope / per-kind override 追加
- completion / semantic token まで含めた shadow 全面 shared 化

## 判断

- shadow 専用 fixture 分離 PoC は実施価値がある
- ただし最初の PoC は extension 側だけに閉じ、server 側正本化は保留する
- PoC 実施後に見るべき指標は「shadow hover / signature の `occurrenceIndex` 差分が消えるか」「document 追加に対して helper 複雑度が許容範囲か」の 2 点とする

## 次の最小候補

1. extension の `ShadowedApplication()` を `ApplicationWorkbookRootShadowed.bas` のような専用 fixture へ切り出す
2. extension test で shadow 用 document を別に開き、shadow hover / signature の `occurrenceIndex = 0` 化を確認する
3. server 側 inline text はそのままにして、PoC 後に shared shadow spec 化の要否を再判断する
