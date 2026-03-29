# dedicated shadow fixture 分離の cross-family 観測メモ

## 目的

workbook root family で採用した `state: "shadowed"` + dedicated shadow fixture 分離が、別の built-in family でも必要になっているかを観測する。  
焦点は「同じ topology 問題が第 2 の family で再発しているか」であり、今すぐ test infrastructure を共通化することではない。

## 観測対象

### 1. workbook root family

現状:

- extension 側は [ApplicationWorkbookRootShadowed.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootShadowed.bas) を dedicated shadow fixture として持つ
- server 側は [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の inline shadow text を維持する
- canonical anchor は [test-support/workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) の `state: "shadowed"` entry で共有している

判断:

- dedicated shadow fixture 分離が実際に必要だった唯一の family
- duplicate anchor と `occurrenceIndex` 差分を吸収する実益があり、completion / hover / signature では server / extension の両方で shared shadow matrix を持っている
- semantic の shadow negative は extension 側だけが shared spec を使っており、server 側まで完全対称ではない

### 2. worksheet broad root family

現状:

- extension 側は [WorksheetBroadRootBuiltIn.bas](../../packages/extension/test/fixtures/WorksheetBroadRootBuiltIn.bas) で broad root の正例 / 非対象 root を 1 つの fixture にまとめている
- server 側には `Worksheets` shadow と `Application` shadow の inline test があるが、これは [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の server-only negative case である
- shared spec `test-support/workbookRootFamilyCaseTables.cjs` の `worksheetBroadRoot` には `state: "shadowed"` が存在しない

判断:

- workbook root family と似た「shadowed root」はあるが、まだ server-only の補助 negative case に留まる
- extension host 側に shadow 専用 matrix も dedicated shadow fixture も無く、topology 差分としては未成熟
- 現時点で dedicated shadow fixture 分離を検討する段階ではない

### 3. built-in signature / hover の shadowed root

現状:

- [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の `BuiltInSignatureShadowed` は、`WorksheetFunction` / `ActiveCell` shadow を 1 本の server-only test で押さえている
- extension 側の shared case table や fixture topology 問題には広がっていない

判断:

- これは matrix family ではなく、isolated negative test である
- dedicated shadow fixture 分離を持ち込む必然性は無い

## 結論

- 現時点で workbook root family と同種の dedicated shadow fixture 分離が必要な第 2 family は無い
- 再発候補として最も近いのは `worksheet broad root family` だが、まだ server-only shadow case であり、shared shadow matrix を持っていない
- したがって、共通運用ルールや shared spec schema 追加をこの時点で一般化しない

## dedicated shadow fixture 分離を再検討する条件

- 第 2 family でも extension / server の両方に shadow matrix ができ、shared anchor spec を要求するようになったとき
- duplicate anchor や `occurrenceIndex` 差分が package 間で再発し、package-local の message / helper だけでは吸収しづらくなったとき
- shadow 用全文 text の drift が、2 PR 以上連続で review 指摘や failure triage の主因になったとき

## 次の見直し候補

1. `worksheet broad root family` の symbol-shadowed case を extension matrix へ広げる必要があるかを整理する
2. 第 2 family が出るまでは、dedicated shadow fixture 分離は workbook root family 固有の対処として扱う
