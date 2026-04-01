# worksheet control shapeName path の vocabulary と canonical anchor source

## 結論

- `worksheetControlShapeNamePath` の v1 語彙は、`rootKind` を `document-module` / `workbook-qualified-static` / `workbook-qualified-matched` / `workbook-qualified-closed`、`routeKind` を `ole-object` / `shape-oleformat` に固定する。
- 負例の分類軸は `rootKind` に混ぜず、`reason` として `numeric-selector` / `dynamic-selector` / `code-name-selector` / `plain-shape` / `chartsheet-root` / `non-target-root` に切り分ける。
- `OleObjectBuiltIn.bas` と `ShapesBuiltIn.bas` は route ごとの実行 fixture として残し、family canonical anchor source は将来の専用 case spec に分離する前提で扱う。
- v1 では `test-support/worksheetControlShapeNamePathCaseTables.cjs` のような repo root 配下の dedicated case spec を最終正本候補とし、単独の `.bas` fixture を family 正本にはしない。

## 目的

[worksheet-control-shared-spec-family-candidate-feasibility.md](./worksheet-control-shared-spec-family-candidate-feasibility.md) で、最初の shared spec 候補 family は `worksheet control shapeName path` と整理した。  
このメモでは、その候補を本当に family table に載せる前提として、

- `rootKind` / `routeKind` / `reason` をどの粒度で固定するか
- `OleObjectBuiltIn.bas` と `ShapesBuiltIn.bas` のどこを canonical anchor source と呼ぶか

を切り分ける。

## 現在の入力 source

### 1. `OleObjectBuiltIn.bas`

- `Sheet1.OLEObjects("CheckBox1").Object`
- `Sheet1.OLEObjects.Item("CheckBox1").Object`
- `ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object`
- `ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object`
- `ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object`
- `ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object`

のように、`ole-object` route の shapeName path を広く持っている。

一方で同じ file には、

- generic `OLEObjects` / `OLEObject` surface
- `Chart1` / `ActiveSheet`
- numeric / dynamic selector
- `Worksheets("Sheet1")` の code-name selector

も混在している。

### 2. `ShapesBuiltIn.bas`

- `Sheet1.Shapes("CheckBox1").OLEFormat.Object`
- `Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object`
- `ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object`
- `ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object`
- `ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object`
- `ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object`

のように、`shape-oleformat` route の shapeName path を広く持っている。

一方で同じ file には、

- generic `Shape` / `OLEFormat` surface
- `Chart1`
- `ShapeRange`
- `PlainShape`
- numeric / dynamic selector
- `Worksheets("Sheet1")` の code-name selector

も混在している。

## 観察結果

### 1. family の主語は root よりも `shapeName を key にした control owner promotion` である

- `ole-object` route も `shape-oleformat` route も、最終的に必要なのは `worksheet owner + shapeName + sidecar 一致` である。
- そのため route の違いは family を分ける理由ではなく、同じ family 内の `routeKind` として持つ方が自然である。
- 一方、generic `OLEObject` / `Shape` / `OLEFormat` surface は `shapeName -> control owner` を使わないので family の外で扱うべきである。

### 2. `numeric` / `dynamic` / `code-name` は rootKind ではなく negative reason である

- `ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object` は root 自体が workbook-qualified であっても、失敗理由は `numeric selector` である。
- `ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object` は worksheet root までの形は正しいが、join key が `sheet name` ではなく `code name` なので失敗する。
- `Chart1.Shapes("CheckBox1").OLEFormat.Object` は route ではなく owner source 側が `unsupported` である。
- したがって、これらを `rootKind` へ混ぜると vocabulary が root と selector と source 状態を同時に表すことになり、family table の軸が崩れる。

### 3. `ActiveWorkbook` の open / closed は v1 では rootKind へ折りたたんだ方が読みやすい

- `ThisWorkbook.Worksheets("Sheet One")` は current bundle static root であり、runtime snapshot を必要としない。
- `ActiveWorkbook.Worksheets("Sheet One")` は同じ text anchor でも `matched` と `closed` の 2 状態を持つ。
- workbook root family では `state` を別軸で持つが、`worksheetControlShapeNamePath` v1 はまだ dedicated case spec を持っておらず、route も 2 本に割れている。
- この段階で `rootKind=workbook-qualified-active` と `state=matched/closed` に分けるより、`workbook-qualified-matched` / `workbook-qualified-closed` を rootKind 側で先に固定した方が、fixture anchor の読み替えが少ない。

### 4. 単独 fixture を family canonical source にすると、route-local 文脈が残りすぎる

- `OleObjectBuiltIn.bas` を正本にすると、`shape-oleformat` route は常に別 file 参照になる。
- `ShapesBuiltIn.bas` を正本にすると、`ole-object` route が外部参照になるだけでなく、generic `Shape` / `OLEFormat` surface が family 主語に見えやすい。
- どちらも broad root 正本ではなく、また file 全体が family 専用に整理されてもいない。
- したがって「どちらか 1 つの fixture を family canonical anchor source にする」のではなく、「route-local fixture を参照する dedicated case spec を repo root に置く」方が境界が明瞭である。

## 固定する vocabulary

### `rootKind`

- `document-module`
  - `Sheet1.OLEObjects("CheckBox1").Object`
  - `Sheet1.Shapes("CheckBox1").OLEFormat.Object`
- `workbook-qualified-static`
  - `ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object`
  - `ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object`
- `workbook-qualified-matched`
  - `ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object`
  - `ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object`
- `workbook-qualified-closed`
  - `ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object`
  - `ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object`
  - ただし snapshot / manifest mismatch、unavailable、disabled により user-facing に開かない状態を表す

### `routeKind`

- `ole-object`
  - `.OLEObjects("shapeName").Object`
  - `.OLEObjects.Item("shapeName").Object`
- `shape-oleformat`
  - `.Shapes("shapeName").OLEFormat.Object`
  - `.Shapes.Item("shapeName").OLEFormat.Object`

### `reason`

- `numeric-selector`
  - `Worksheets(1)`、`.Item(1)`、`Shapes(1)`、`OLEObjects(1)` のように string literal `shapeName` / `sheetName` が確定しない
- `dynamic-selector`
  - `GetIndex()`、`i + 1` のように compile time で `shapeName` / `sheetName` が確定しない
- `code-name-selector`
  - `Worksheets("Sheet1")` のように `sheet name` ではなく `code name` で解決しようとしている
- `plain-shape`
  - `Shapes("PlainShape").OLEFormat.Object`
- `chartsheet-root`
  - `Chart1.OLEObjects("CheckBox1").Object`
  - `Chart1.Shapes("CheckBox1").OLEFormat.Object`
- `non-target-root`
  - `ActiveSheet.OLEObjects("CheckBox1").Object`
  - broad root や `Sheet1.ControlCodeName` family に属するものはここへ含めず、別 family 側で扱う

## canonical anchor source の扱い

### 1. v1 の正本は dedicated case spec を前提にする

最終的な family canonical anchor source 候補:

- `test-support/worksheetControlShapeNamePathCaseTables.cjs`

ここに持つもの:

- fixture path
- anchor
- `rootKind`
- `routeKind`
- `reason`
- `scopes`

v1 では別の `state` 軸を足さず、`ActiveWorkbook` path の open / closed は `rootKind=workbook-qualified-matched/workbook-qualified-closed` で表す。

ここに持たないもの:

- `CompletionItem.detail`
- async wait 条件
- package ごとの failure message
- decoded hover / signature / semantic token の最終 assertion shape

### 2. `OleObjectBuiltIn.bas` と `ShapesBuiltIn.bas` は execution source として残す

- `ole-object` route の anchor は [OleObjectBuiltIn.bas](../../packages/extension/test/fixtures/OleObjectBuiltIn.bas) を参照する
- `shape-oleformat` route の anchor は [ShapesBuiltIn.bas](../../packages/extension/test/fixtures/ShapesBuiltIn.bas) を参照する
- broad root は [WorksheetBroadRootBuiltIn.bas](../../packages/extension/test/fixtures/WorksheetBroadRootBuiltIn.bas) と [workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) の既存 family に残す

つまり、family canonical source は

- 単独 fixture ではない
- broad root family の既存正本でもない
- route-local fixture を束ねる dedicated case spec

という 3 層構成で考える。

### 3. dedicated mixed fixture はこの段階では作らない

- `OleObjectBuiltIn.bas` と `ShapesBuiltIn.bas` には、それぞれ generic surface の回帰も入っている
- family 専用 mixed fixture を今すぐ足すと、route-local regression と family canonical anchor を二重管理しやすい
- 先に `test-support` 側の case spec で family vocabulary を固定し、その後でも anchor drift が大きければ fixture 分離を再検討する方が安全である

## 判断

### 採用

- `worksheetControlShapeNamePath` の語彙は `rootKind` 4 種、`routeKind` 2 種、negative `reason` 6 種で固定する
- family canonical anchor source は dedicated case spec を前提にし、単独 fixture を family 正本と呼ばない
- route-local fixture と package-local adapter expectation の境界は workbook root family と同じ方針で保つ

### 非採用

- `rootKind` に `numeric` / `dynamic` / `plain-shape` / `chartsheet` を入れる
- `OleObjectBuiltIn.bas` を単独の family 正本にする
- `ShapesBuiltIn.bas` を単独の family 正本にする
- broad root family の正本へ `worksheetControlShapeNamePath` を吸収する
- family 専用 mixed fixture を先に増やしてから vocabulary を固定する

## 次段の候補

1. `worksheetControlShapeNamePath` の dedicated case spec をどの粒度で切るか整理する  
2. `test-support/worksheetControlShapeNamePathCaseTables.cjs` を置くときに、fixture path / anchor / scopes をどう最小化するか PoC する  
3. dedicated mixed fixture が本当に必要かを、case spec 抽出後の drift 量で再評価する  

## 関連文書

- family 候補の切り出し: [worksheet-control-shared-spec-family-candidate-feasibility.md](./worksheet-control-shared-spec-family-candidate-feasibility.md)
- shared case spec の正本方針: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
- `ole-object` route の入口整理: [worksheet-chart-control-entrypoint-feasibility.md](./worksheet-chart-control-entrypoint-feasibility.md)
- `shape-oleformat` route の入口整理: [worksheet-chart-shapes-root-feasibility.md](./worksheet-chart-shapes-root-feasibility.md), [shape-oleformat-object-promotion-feasibility.md](./shape-oleformat-object-promotion-feasibility.md)
