# worksheet broad root family の server non-target hover / signature negative を shared scope へ広げる要否

## 結論

- 現時点では `worksheetBroadRoot` の non-target hover / signature negative を server scope へ広げない。
- `test-support/workbookRootFamilyCaseTables.cjs` に `server-worksheet-broad-root-direct` / `server-worksheet-broad-root-item` 向けの interaction negative entry は追加しない。
- server 側は引き続き、non-target root の negative は completion だけ shared spec を使い、hover / signature は positive anchor の closed-state coverage と extension E2E に委ねる。

## 目的

`worksheet broad root family` では、shared case spec に non-target hover / signature negative がある一方、server 側はそれを使わず、completion の negative だけを shared 化している。  
このメモでは、その非対称を今の段階で server まで広げるべきかを整理する。

## 現状

### shared spec

- [test-support/workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) の `worksheetBroadRoot.hover.negative` / `signature.negative` は `extension` scope だけを持つ。
- 同じ family の `completion.negative` だけは `server-worksheet-broad-root-direct` / `server-worksheet-broad-root-item` scope まで持っている。

### server 側

- [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の broad root direct / item test は、non-target root について shared completion negative を使っている。
- hover / signature は positive anchor だけ shared spec から読み、`snapshot unavailable` / `mismatch` では閉じることを確認している。
- `Sheets` / `ActiveSheet` / numeric selector / dynamic selector に対する hover / signature negative は、server では個別に追加していない。

### extension 側

- [packages/extension/test/suite/index.ts](../../packages/extension/test/suite/index.ts) は `worksheetBroadRoot.hover.negative` / `signature.negative` を shared spec から読み、`WorksheetBroadRootBuiltIn.bas` 上で end-to-end に固定している。
- negative 対象は `Sheets` / `ActiveSheet` / numeric selector / dynamic selector / root `.Item(1)` / root `.Item(GetIndex())` まで含む。

## 観察結果

### 1. server が今欠いているのは interaction API の「closed-state coverage」ではない

- server 側では、shared positive anchor を使って `snapshot unavailable` / `mismatch` 時に hover / signature が閉じることを既に確認している。
- つまり `getHoverAfterToken()` / `getSignatureHelpAfterToken()` が broad root family に対して閉じる基本動作自体は、すでに server test で踏めている。
- 追加で non-target root の interaction negative を足しても、「hover / signature API が閉じる」ことを別 anchor で重ねる比重が大きい。

### 2. non-target root の主論点は root / selector gating であり、completion negative がその境界を先に固定している

- `Sheets` / `ActiveSheet` は broad root family の対象外 root であり、`Worksheets(1)` / `Worksheets(GetIndex())` / `Worksheets.Item(1)` / `Worksheets.Item(GetIndex())` は selector 境界で閉じる。
- この境界は completion negative を server 側ですでに shared 化しており、`symbol resolution -> workbook root family gating` の失敗点を直接確認できる。
- interaction negative を追加しても、主に再確認するのはその後段の「開かなかった」という結果であり、失敗点の説明力は大きく増えない。

### 3. extension E2E が non-target interaction negative の user-facing 境界をすでに固定している

- broad root family の non-target hover / signature negative は、ユーザー視点では「候補が出ない」「hover が出ない」「signature help が出ない」が一連の挙動である。
- この user-facing 境界は extension 側で `WorksheetBroadRootBuiltIn.bas` を通して固定済みで、host 経由の drift も検知できる。
- server に同じ anchor 群を足しても、host 非同期や editor command を含まないため、extension 側で得ている観測と質的に大きく変わらない。

### 4. server scope を足すと shared spec の entry 数は増えるが、直ちに別 family へ再利用しにくい

- broad root direct / item の 2 scope に non-target hover / signature negative を追加すると、shared spec の entry 数と scope 組み合わせは増える。
- しかしその増分は、現時点では `worksheetBroadRoot` server test のみが消費する。
- `shared spec に載っているが実際には 1 package の 1 family しか使わない entry` を増やすより、今の境界を feature memo で固定した方が review 時に追いやすい。

## 判断

### 今回は server scope へ広げない

- `worksheetBroadRoot.hover.negative` / `signature.negative` に server scope を追加しない。
- server 側は completion negative だけ shared 化し、hover / signature の non-target negative は extension E2E を正本の user-facing coverage とみなす。
- broad root direct / item test では、引き続き positive interaction anchor の closed-state coverage を維持する。

### 再評価の条件

- server 側で broad root non-target hover / signature に関する回帰が実際に起き、completion negative と positive closed-state だけでは原因切り分けが弱いと分かったとき
- `worksheetBroadRoot` 以外の family でも同種の「completion negative は shared、interaction negative は extension-only」が増え、shared spec 境界として説明しづらくなったとき
- `reviewer` または CodeRabbit から、server broad root test に non-target interaction negative 欠落を問題視する指摘が複数回続いたとき
- broad root family の interaction path に route 別 fallback や API 別分岐が増え、completion negative だけでは後段差分を説明できなくなったとき

## 推奨方針

### 維持するもの

- server 側の shared completion negative
- server 側の positive hover / signature anchor と closed-state coverage
- extension 側の non-target hover / signature negative shared entry

### 今やらないこと

- `worksheetBroadRoot` interaction negative に `server-worksheet-broad-root-direct` / `server-worksheet-broad-root-item` を追加する
- server 側で non-target hover / signature negative 専用 helper や local table を新設する
- broad root family の shared spec 境界を、この時点で workbook root family と完全対称にそろえようとする

## 関連文書

- shared spec の正本: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
- broad root shadow 境界: [worksheet-broad-root-shadow-extension-matrix-feasibility.md](./worksheet-broad-root-shadow-extension-matrix-feasibility.md), [worksheet-broad-root-shadow-server-helper-feasibility.md](./worksheet-broad-root-shadow-server-helper-feasibility.md), [worksheet-broad-root-shadow-coverage-feasibility.md](./worksheet-broad-root-shadow-coverage-feasibility.md)
- broad root family の機能正本: [explicit-sheet-name-broad-root-feasibility.md](./explicit-sheet-name-broad-root-feasibility.md)
