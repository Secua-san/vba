# worksheet/chart control 系の shared spec 候補 family 切り出し

## 結論

- `worksheet/chart control` 系を 1 つの shared spec family にまとめず、最初の候補は `worksheet control shapeName path` に絞る。
- 具体的には、`OLEObjects("shapeName").Object` と `Shapes("shapeName").OLEFormat.Object` のうち、worksheet identity が静的に確定している path を同じ family 候補として扱う。
- `Sheet1.ControlCodeName` は別 family 候補として切り分ける。join key が `codeName` であり、shape name path と混ぜない。
- `Worksheets("Sheet One")` / `Application.Worksheets("Sheet One")` の broad root は、control family に吸収せず既存の workbook root family に残す。
- chartsheet path は sidecar 上で `unsupported` のままなので、初回候補 family から外す。

## 目的

前段の [workbook-root-family-server-mirror-cross-family-preconditions.md](./workbook-root-family-server-mirror-cross-family-preconditions.md) では、他 family へ server mirror policy を持ち出す前に、まず shared spec 候補となる family の境界を決める必要があると整理した。  
このメモでは、`worksheet/chart control` 系のうち、どこまでを 1 つの family table に載せると自然かを切り分ける。

## 現在の論点の分布

### 1. `OLEObjects(...).Object`

- [OleObjectBuiltIn.bas](../../packages/extension/test/fixtures/OleObjectBuiltIn.bas) と server test は、
  - document module root `Sheet1`
  - workbook-qualified root `ThisWorkbook.Worksheets(...)` / `ActiveWorkbook.Worksheets(...)`
  - chartsheet / `ActiveSheet` / numeric / dynamic selector
  を 1 つの fixture 群で持っている。
- broad root `Worksheets(...)` / `Application.Worksheets(...)` の canonical source はここではなく、[WorksheetBroadRootBuiltIn.bas](../../packages/extension/test/fixtures/WorksheetBroadRootBuiltIn.bas) と [workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) の `worksheetBroadRoot` family にある。
- ただし generic `OLEObject` surface と、`.Object` 後段の control owner promotion が同じファイルに混在している。

### 2. `Shapes(...).OLEFormat.Object`

- [ShapesBuiltIn.bas](../../packages/extension/test/fixtures/ShapesBuiltIn.bas) と server test は、
  - generic `Shape`
  - generic `OLEFormat`
  - `Shape.OLEFormat.Object` 後段の control owner promotion
  を同じ fixture 群で持っている。
- shape name literal と numeric / dynamic / `ShapeRange` / plain shape / chartsheet / workbook-qualified root が同居している。
- `Worksheets("Sheet One")` / `Application.Worksheets("Sheet One")` を軸にした broad root canonical source はここではなく、[WorksheetBroadRootBuiltIn.bas](../../packages/extension/test/fixtures/WorksheetBroadRootBuiltIn.bas) と [workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) に分かれている。`ShapesBuiltIn.bas` にあるのは workbook-qualified / explicit sheet-name path の断片だけである。

### 3. `Sheet1.ControlCodeName`

- [WorksheetControlCodeName.bas](../../packages/extension/test/fixtures/WorksheetControlCodeName.bas) と server test は、`sheetCodeName + codeName` で control owner へ進む direct access だけを扱う。
- document module root `Sheet1` だけを対象にしており、workbook-qualified root や broad root はまだ論点に入っていない。

### 4. broad root family

- `Worksheets("Sheet One")` / `Application.Worksheets("Sheet One")` は、既に [test-support/workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) で `worksheetBroadRoot` family として shared spec 化されている。
- fixture 正本は [WorksheetBroadRootBuiltIn.bas](../../packages/extension/test/fixtures/WorksheetBroadRootBuiltIn.bas) で、`OLEObjects(...).Object` / `Shapes(...).OLEFormat.Object` の broad root anchor はこの family から読む。
- ここで見ている主語は `control owner promotion` だけでなく、workbook binding / snapshot gating を含む broad root identity である。

## 観察結果

### 1. broad root を control family に吸収すると、既存 workbook root family と正本が衝突する

- `worksheetBroadRoot` は既に shared spec と server mirror policy を持つ family である。
- ここへ `OLEObjects(...).Object` / `Shapes(...).OLEFormat.Object` を「control family」として重ねると、[WorksheetBroadRootBuiltIn.bas](../../packages/extension/test/fixtures/WorksheetBroadRootBuiltIn.bas) と [workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) を正本にしている同じ anchor が
  - workbook root family
  - worksheet control family
  の二重正本になりやすい。
- broad root path の主論点は shape name path そのものよりも root gating なので、control family 側へ引き取らない方が境界が明瞭である。

### 2. `shapeName` path と `codeName` path は join key が違う

- `OLEObjects("CheckBox1").Object` と `Shapes("CheckBox1").OLEFormat.Object` は、sidecar 上の `shapeName -> controlType` を使う。
- `Sheet1.chkFinished` は、sidecar 上の `sheetCodeName + codeName -> controlType` を使う。
- 同じ control owner に到達しても、selector / key / root の vocabulary が異なるため、最初から 1 family にまとめると `reason` / `state` 語彙が不揃いになる。

### 3. `OLEObjects(...).Object` と `Shapes(...).OLEFormat.Object` は、shapeName path として同居させやすい

- どちらも最終的には
  - worksheet owner が静的に確定している
  - shape name string literal が取れる
  - sidecar 上の `shapeName -> controlType` が引ける
  という同じ前提で control owner へ昇格する。
- route は違っても、`shapeName` を join key にした control owner promotion という 1 つの主語で並べやすい。
- したがって、最初の shared spec 候補としては `shapeName path` を family の中心に置くのが自然である。

### 4. chartsheet path は候補 family に入れる前提がまだ無い

- sidecar v1 では chartsheet owner は `status: "unsupported"` のままである。
- `Chart1.OLEObjects("CheckBox1").Object` や `Chart1.Shapes("CheckBox1").OLEFormat.Object` を control owner へ進める静的根拠が無いため、worksheet と同じ family table に載せられない。
- chartsheet を混ぜると、unsupported source の話と shared spec family の話が混ざるので、初回候補から外すべきである。

### 5. generic `OLEObject` / `Shape` / `OLEFormat` surface は control family の外に置くべきである

- `Sheet1.OLEObjects(1).Name` は generic `OLEObject` surface の論点であり、`control owner promotion` の論点ではない。
- `Sheet1.Shapes("CheckBox1").OLEFormat.ProgID` も generic `OLEFormat` surface である。
- これらまで同じ family に入れると、sidecar を使う path と使わない path が混ざり、family の主語がぼやける。

## 判断

### 最初の shared spec 候補は `worksheet control shapeName path`

この候補 family に含めるもの:

- `Sheet1.OLEObjects("shapeName").Object`
- `Sheet1.OLEObjects.Item("shapeName").Object`
- `Sheet1.Shapes("shapeName").OLEFormat.Object`
- `Sheet1.Shapes.Item("shapeName").OLEFormat.Object`
- `ThisWorkbook.Worksheets("Sheet One").OLEObjects("shapeName").Object`
- `ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("shapeName").Object`
- `ThisWorkbook.Worksheets("Sheet One").Shapes("shapeName").OLEFormat.Object`
- `ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("shapeName").OLEFormat.Object`
- `ActiveWorkbook.Worksheets("Sheet One")` / `.Item("Sheet One")` の matched / closed path

共通点:

- worksheet owner identity が静的に取れる
- join key が `shapeName`
- sidecar を同じ用途で使う
- terminal owner が同じ control owner 群

### broad root は workbook root family に残す

含めないもの:

- `Worksheets("Sheet One").OLEObjects("shapeName").Object`
- `Application.Worksheets("Sheet One").Shapes("shapeName").OLEFormat.Object`
- root `.Item("Sheet One")` を含む broad root variant

理由:

- 既に `worksheetBroadRoot` shared spec の正本がある
- root gating / workbook binding が主論点で、control family より workbook root family の文脈で見る方が自然

### `Sheet1.ControlCodeName` は別候補 family にする

含めないもの:

- `Sheet1.chkFinished`
- `Sheet1.CheckBox1`

理由:

- join key が `codeName`
- いまの coverage は document module root だけ
- workbook-qualified root / broad root / chartsheet との関係も別途整理が必要

## family 候補のたたき台

### 候補 A: `worksheetControlShapeNamePath`

軸:

- root kind
  - `document-module`
  - `workbook-qualified-static`
  - `workbook-qualified-matched`
  - `workbook-qualified-closed`
- route kind
  - `ole-object`
  - `shape-oleformat`
- selector kind
  - `string-literal`
  - `numeric`
  - `dynamic`
  - `plain-shape`
- state
  - `supported`
  - `closed`
  - `unsupported`

### 候補 B: `worksheetControlCodeNamePath`

軸:

- root kind
  - `document-module`
  - 将来の `workbook-qualified` は別判断
- identifier kind
  - `known-code-name`
  - `unknown-code-name`
- state
  - `supported`
  - `unsupported`

## 次段の候補

1. `worksheetControlShapeNamePath` を shared spec 候補として、root / route / selector / state vocabulary を固定する  
2. `OleObjectBuiltIn.bas` と `ShapesBuiltIn.bas` のうち、control owner promotion に対応する canonical anchor source をどこへ寄せるか整理する  
3. `worksheetControlCodeNamePath` を別 family 候補として残すか、shapeName path との関係をどう説明するかを後続で整理する  

## 今やらないこと

- broad root を新しい control family へ移す
- `Sheet1.ControlCodeName` を shapeName path と同じ table に混ぜる
- chartsheet path を supported worksheet path と同じ family 候補へ入れる
- generic `OLEObject` / `Shape` / `OLEFormat` surface まで同じ family に含める

## 関連文書

- 前段条件: [workbook-root-family-server-mirror-cross-family-preconditions.md](./workbook-root-family-server-mirror-cross-family-preconditions.md)
- 入口整理: [worksheet-chart-control-entrypoint-feasibility.md](./worksheet-chart-control-entrypoint-feasibility.md)
- shape path 整理: [worksheet-chart-shapes-root-feasibility.md](./worksheet-chart-shapes-root-feasibility.md), [shape-oleformat-object-promotion-feasibility.md](./shape-oleformat-object-promotion-feasibility.md)
- code name / sidecar: [worksheet-chart-control-identity-feasibility.md](./worksheet-chart-control-identity-feasibility.md), [worksheet-control-metadata-sidecar-artifact.md](./worksheet-control-metadata-sidecar-artifact.md)
