# worksheet broad root family の shadow case を extension matrix へ広げる要否

## 結論

- 現時点では `Worksheets` shadow と `Application` shadow を extension matrix へ広げない。
- worksheet broad root family の shadow case は、引き続き server-only の negative case として扱う。
- `test-support/workbookRootFamilyCaseTables.cjs` に `worksheetBroadRoot` の `state: "shadowed"` を追加しない。
- extension 側に dedicated shadow fixture はまだ導入しない。

## 目的

`worksheet broad root family` では、built-in broad root の正例と snapshot gating、さらに extension 側の non-target hover / signature negative は shared case spec へ寄っている一方、`Worksheets` shadow と `Application` shadow は server-only test のまま残っている。  
このメモでは、その非対称を今の段階で extension matrix へ広げるべきかを整理する。

## 現状

### extension 側

- fixture は [WorksheetBroadRootBuiltIn.bas](../../packages/extension/test/fixtures/WorksheetBroadRootBuiltIn.bas) の 1 本だけで、built-in broad root の正例と non-target root をまとめて持つ。
- [packages/extension/test/suite/index.ts](../../packages/extension/test/suite/index.ts) は `worksheetBroadRoot` の shared case spec を使い、completion / hover / signature を matched / closed / non-target で回している。
- shadow 用の section や shadow 専用 document は存在しない。

### server 側

- [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) に `Worksheets` shadow と `Application` shadow の単発 negative test がある。
- どちらも completion / hover / signature の 3 点だけを確認し、manifest match 下でも broad root を開かないことを直接固定している。
- broad root shadow は shared spec へ載せず、inline text の局所 test として閉じている。

## 観察結果

### 1. 現在の shadow 論点は root identifier の built-in 判定だけで閉じる

- broad root shadow case の本質は、`Worksheets` または `Application` が user-defined symbol に shadow されたとき、built-in broad root family に入れないことを保証する点にある。
- これは document service 内の symbol resolution / built-in gating の問題であり、VS Code host 経由の非同期挙動や sidecar I/O の差が主論点ではない。
- そのため現状の coverage では、server 側の直接テストが最短であり、extension E2E を足しても新しい failure mode は増えにくい。

### 2. extension へ広げると fixture topology が変わる

- いまの [WorksheetBroadRootBuiltIn.bas](../../packages/extension/test/fixtures/WorksheetBroadRootBuiltIn.bas) は、matched / non-target を 1 本で追うための fixture であり、shadow 宣言は含んでいない。
- shadow case を extension へ足すには、mixed fixture に shadow section を混ぜるか、shadow 専用 fixture を別で作るかのどちらかになる。
- 前者は duplicate anchor と `occurrenceIndex` 管理の論点を持ち込みやすく、後者は document 数と helper 分岐を増やす。
- workbook root family ではそのコストを払うだけの shared shadow matrix があったが、worksheet broad root family はまだそこまで成熟していない。

### 3. broad root shadow は shared spec 化の受益がまだ小さい

- `worksheetBroadRoot` の shared spec は、現時点では built-in broad root の正例と non-target root を extension / server でそろえるために使っている。
- shadow case は server 側にしか存在せず、shared spec に `state: "shadowed"` を足しても immediate な再利用先が無い。
- extension 側で semantic shadow negative まで持ち込むと、workbook root family と同じく extension-only な補助 entry が増え、shared spec の対称性がむしろ下がる。

### 4. current broad root family の主要リスクは別の場所にある

- broad root family の user-facing で壊れやすい箇所は、manifest match / mismatch、root `.Item("SheetName")`、`OLEObject.Object` と `Shape.OLEFormat.Object` の対称性である。
- これらはすでに extension + server の shared matrix で固定している。
- 一方 shadow case は conservative negative であり、product behavior の広がりを増やす変更ではない。

## 判断

### 今回は extension matrix へ広げない

- `Worksheets` shadow と `Application` shadow は server-only の単発 negative case のまま維持する。
- `worksheetBroadRoot` には `state: "shadowed"` を追加しない。
- extension 側の dedicated shadow fixture 分離も行わない。

### 再評価の条件

- extension 側で broad root shadow に関する実不具合が出て、server unit test だけでは再現・防止しづらいと分かったとき
- `worksheet broad root family` でも completion / hover / signature / semantic の 2 種類以上を shadow 状態で shared 化したい要件が出たとき
- broad root shadow の review 指摘や triage が 2 PR 以上連続し、server-only 局所 test のままでは drift を抑えにくいと判断できたとき
- `Worksheets` shadow と `Application` shadow 以外に broad root family 固有の shadow variant が増え、配列駆動 helper へ寄せた方が読みやすくなったとき

## 推奨方針

### 維持するもの

- server 側の `Worksheets` shadow / `Application` shadow 単発 negative test
- extension 側の [WorksheetBroadRootBuiltIn.bas](../../packages/extension/test/fixtures/WorksheetBroadRootBuiltIn.bas) 1 本構成
- shared spec `worksheetBroadRoot` の対象を built-in positive / non-target / snapshot-closed family に限定する構造

### 今やらないこと

- extension 側に broad root shadow 専用 fixture を足す
- `test-support/workbookRootFamilyCaseTables.cjs` に `worksheetBroadRoot.state = "shadowed"` を追加する
- workbook root family と同じ dedicated shadow fixture 運用を broad root family へ先回りで広げる

## 関連文書

- broad root family の正本: [explicit-sheet-name-broad-root-feasibility.md](./explicit-sheet-name-broad-root-feasibility.md)
- cross-family 観測: [shadow-fixture-split-cross-family-observation.md](./shadow-fixture-split-cross-family-observation.md)
- workbook root family の shared spec 境界: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
