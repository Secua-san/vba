# workbook root family の shadow 専用 fixture 分離 実装メモ

## 目的

`Application.ThisWorkbook` / `Application.ActiveWorkbook` の shadow case について、extension 側の mixed fixture を分離し、server 側 inline shadow fixture と同じ anchor topology に寄せる。

## 実装結果

- extension 側は [ApplicationWorkbookRootBuiltIn.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootBuiltIn.bas) から `ShadowedApplication()` を抜き、[ApplicationWorkbookRootShadowed.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootShadowed.bas) を追加した
- [packages/extension/test/suite/index.ts](../../packages/extension/test/suite/index.ts) は static / matched 用 document と shadow 用 document を別々に開く構成へ変更した
- extension 側は shadow semantic negative を `state: "shadowed"` へ寄せ、main fixture と shadow fixture の anchor 衝突を shared spec 側で解消した
- server 側は [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の inline shadow text を維持したまま、shadow hover / signature も shared spec から引くように寄せた

## 変わったこと

### extension 側

- shadow hover / signature の direct anchor は専用 fixture 上で一意になり、`occurrenceIndex = 0` で扱えるようになった
- shadow completion / semantic negative も dedicated shadow fixture と `state: "shadowed"` の shared entry で扱えるようになった
- helper 変更は「document が 1 枚増える」「shadow semantic の空 token 待ちを安定化する」範囲に留まり、assertion helper 自体の全面刷新は不要だった

### server 側

- inline text の局所性は維持した
- shadow hover / signature も `test-support/workbookRootFamilyCaseTables.cjs` の shared entry を参照し、extension と同じ anchor 正本を使うようになった

## 判断

- shadow 専用 fixture 分離は有効だった
- extension と server の shadow hover / signature は、scope ごとの `occurrenceIndex` override を追加せずに shared 化できた
- server 側 canonical shadow text source は、この段階では不要と判断し、後続の [workbook-root-shadow-text-source-canonicalization-feasibility.md](workbook-root-shadow-text-source-canonicalization-feasibility.md) を正本にする

## 残したもの

- server 側の shadow text は inline のまま
- canonical shadow text source の導入判断は保留
- package-local に残すのは failure message、`CompletionItem.detail` 断片、decoded token / hover / signature help の assertion shape だけ

## 次の見直し候補

1. 同種の dedicated shadow fixture 分離が別 family でも必要になるかを観測する
2. shadow text drift が review コストの主因になった時だけ、canonical text source の再比較へ戻る
