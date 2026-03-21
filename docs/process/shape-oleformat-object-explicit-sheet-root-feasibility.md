# Shape.OLEFormat.Object Explicit Sheet-Name Root Feasibility

## 結論

- 長く効く判断の正本は ADR [0005 Explicit Sheet-Name Root Policy](../adr/0005-explicit-sheet-name-root-policy.md) とし、この文書は調査結果と実装境界の補足を扱う。
- `Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object` のような explicit sheet-name root は、sidecar schema 上は `sheetName + shapeName` で結合できる。
- ただし current product で直ちに user-facing へ広げる対象は、`ThisWorkbook.Worksheets("Sheet1")` のように workbook identity が静的に固定できる root に絞るべきである。
- unqualified `Worksheets("Sheet1")` と `Application.Worksheets("Sheet1")` は active workbook 依存であり、この文書の段階では不採用とする。`ActiveWorkbook.Worksheets("Sheet1")` は後続の workbook-bound broad root gating で別途扱う。
- join key は `sheetCodeName` ではなく `sheetName` を使う。`sheetCodeName` は worksheet document module alias (`Sheet1`) と control code name 導線のために別 key として維持する。
- `OLEObject.Object` と `Shape.OLEFormat.Object` は、将来 explicit sheet-name root を開くなら同じ `workbook root identity + sheetName + shapeName` lookup helper を共有できる。

## 確認した公式ソース

### Office VBA

- [Worksheet object (Excel)](https://learn.microsoft.com/office/vba/api/excel.worksheet)
  - `Worksheets(index)` は worksheet index number または name で単一 worksheet を返す。
  - `Worksheets("Sheet1")` の `"Sheet1"` は worksheet 名であり、tab に表示される名前として説明されている。
- [Workbook.Worksheets property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.worksheets)
- [Application.Worksheets property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.worksheets)
  - object qualifier が無い `Worksheets` は active workbook を対象にする。
- [Refer to Sheets by Name](https://learn.microsoft.com/office/vba/excel/concepts/workbooks-and-worksheets/refer-to-sheets-by-name)
  - worksheet / chart / dialog sheet は collection の name access で参照できる。
- [Worksheet.CodeName property (Excel)](https://learn.microsoft.com/office/vba/api/excel.worksheet.codename)
  - code name は `Sheet1.Range("A1")` のように expression の代わりに使えるが、sheet name とは独立に変更され得る。
- [Using ActiveX Controls on Sheets](https://learn.microsoft.com/office/vba/excel/concepts/controls-dialogboxes-forms/using-activex-controls-on-sheets)
  - `Shapes` / `OLEObjects` collection で control を名前指定するときは code name ではなく shape name を使う。
  - `Worksheets(1).OLEObjects("CommandButton1").Object.Caption` のように、sheet collection root から `.Object` を辿る例がある。

## 現行実装と sidecar の前提

- current product は worksheet document module alias (`Sheet1`) から始まる path に限って sidecar lookup を許可している。
- sidecar v1 には `sheetName` と `sheetCodeName` の両方が必須 field として入っている。
- `shapeName` と `codeName` は別 key であり、`Shapes("CheckBox1")` / `OLEObjects("CheckBox1")` は `shapeName`、`Sheet1.chkFinished` は `codeName` を使う。
- server の sidecar resolver は document module root 解決時にだけ `rootUri` と `rootModuleName` を持つ。generic `Worksheet` root へ降りた時点では workbook bundle identity を保持していない。

## 観察結果

### 1. `Worksheets("Sheet1")` は `sheetCodeName` ではなく `sheetName` を使う

- `Worksheet` object の正本は `Worksheets(index)` の string selector を worksheet 名として説明している。
- `Worksheet.CodeName` の正本は、code name は `Sheet1.Range("A1")` のように expression の代替であり、sheet name とは別に変更され得るとしている。
- したがって `Worksheets("Sheet1")` を `sheetCodeName` に結び付けるのは誤りであり、explicit sheet-name root を開くなら join key は `sheetName` を使うべきである。

### 2. workbook identity が無いと sidecar lookup は安全にできない

- unqualified `Worksheets("Sheet1")` は active workbook を対象にする。
- `ActiveWorkbook.Worksheets("Sheet1")` も runtime 状態の active workbook を指すため、現在編集中の bundle と同一 workbook だと静的には言えない。
- 一方 `ThisWorkbook.Worksheets("Sheet1")` は local workbook document module alias を起点にできるため、current bundle の sidecar へ結ぶ前提を置きやすい。
- そのため explicit sheet-name root を開く最初の候補は `ThisWorkbook` qualified path に限定するのが最も保守的である。

### 3. current resolver には generic `Worksheet` root の provenance が無い

- `resolveBuiltinMemberOwnerForPath()` は root symbol が document module へ解決できた場合にだけ `DocumentModuleBuiltinContext` を作る。
- `Sheet1` root は `rootUri` と `rootModuleName` を持てるが、`Worksheets("Sheet1")` や `ActiveWorkbook.Worksheets("Sheet1")` は generic `Worksheet` owner へ降りた時点で sheet identity を失う。
- したがって current 実装へそのまま条件分岐を足しても、`which workbook / which worksheet` を sidecar から選べない。

### 4. `OLEObject.Object` と `Shape.OLEFormat.Object` は sheet-name lookup を共有できる

- `Using ActiveX Controls on Sheets` の正本では、`OLEObjects("CommandButton1").Object` も `Shapes` collection access も同じ shape name を使う。
- したがって explicit sheet-name root を開くなら、`sheetName + shapeName -> controlType` の lookup helper を 1 本用意し、`OLEObject.Object` と `Shape.OLEFormat.Object` の両方で共有する方が一貫する。
- 一方 `Sheet1.chkFinished` の direct access は引き続き `sheetCodeName + codeName` を使う別導線のまま維持するべきである。

### 5. `ActiveSheet` / chartsheet / `ShapeRange` はこの議論に混ぜない

- `ActiveSheet` は workbook と sheet の両方が runtime 状態依存であり、explicit sheet-name root よりさらに不安定である。
- chartsheet は current sidecar で `unsupported` のままで、worksheet と同じ inventory source を前提にできない。
- `Shapes.Range(Array(...))` は `ShapeRange` であり、単一 `shapeName` lookup とは別 surface である。
- したがって explicit sheet-name root の検討でも、これらは引き続き除外境界として維持する。

## 推奨方針

### フェーズ 1: docs 上の整理完了

- explicit sheet-name root の join key は `sheetName`、document module alias / control code name 導線の join key は `sheetCodeName` と整理する。
- `Worksheets("Sheet1")` と `Application.Worksheets("Sheet1")` は active workbook 依存のため、この段階では user-facing にしない。`ActiveWorkbook.Worksheets("Sheet1")` は後続の workbook-bound broad root gating で扱う。
- 最初の実装候補は `ThisWorkbook.Worksheets("Sheet1")` に限定する。

### フェーズ 2: workbook-qualified worksheet root の最小実装

- `ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object`
- `ThisWorkbook.Worksheets("Sheet1").Shapes.Item("CheckBox1").OLEFormat.Object`
- `ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object`
- `ThisWorkbook.Worksheets.Item("Sheet1").Shapes.Item("CheckBox1").OLEFormat.Object`
- 必要なら同時に `ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object` 系も同じ helper へ寄せる。
- negative は `ActiveWorkbook`、unqualified `Worksheets`、`Application.Worksheets`、`ActiveSheet`、chartsheet、numeric / dynamic selector、`ShapeRange` を維持する。

## 2026-03-15 時点の完了状態

- `ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object` / `.Item("CheckBox1")` と `ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object` / `.Item("CheckBox1")` は user-facing に解決する。
- shared helper により `ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object` と `.Item("CheckBox1").Object` も同じ workbook-qualified root から user-facing に解決する。
- resolver は `ThisWorkbook` 起点の workbook root identity を `Worksheets("Sheet1")` 連鎖でも保持し、current bundle の sidecar から `sheetName + shapeName` を引ける。
- `ThisWorkbook.Worksheets(1)`、unqualified `Worksheets("Sheet1")`、`Application.Worksheets("Sheet1")`、`ActiveSheet`、chartsheet、`ShapeRange` は引き続き除外境界として維持する。`ActiveWorkbook.Worksheets("Sheet1")` は後続の broad root gating で user-facing 化済みだが、この文書の対象外とする。
- `sheetName + shapeName` lookup helper は `OLEObject.Object` と `Shape.OLEFormat.Object` の両方で共有できる形へ寄せた。

### フェーズ 3: broad root 展開の再評価

- broad root の扱いは正本 [explicit-sheet-name-broad-root-feasibility.md](./explicit-sheet-name-broad-root-feasibility.md) に分離した。
- 2026-03-21 時点で `ActiveWorkbook.Worksheets("SheetName")` は workbook-bound gating により user-facing 化済みである。unqualified `Worksheets("SheetName")` / `Application.Worksheets("SheetName")` は [explicit-sheet-name-broad-root-feasibility.md](./explicit-sheet-name-broad-root-feasibility.md) で別途扱う。
- 再評価は、current bundle と target workbook の同一性を明示できる workbook binding が導入されたときだけ行う。

## この文書作成時の完了条件

- explicit sheet-name root の join key を `sheetName` と確定する。
- `Sheet1` alias 限定実装と競合しない理由を整理する。
- `ThisWorkbook` qualified root だけを次段候補とし、`ActiveWorkbook` / unqualified `Worksheets` / `ActiveSheet` は除外境界として残す。
- `OLEObject.Object` と `Shape.OLEFormat.Object` の helper 共通化余地を文書化する。

## 次段の候補

- broad root を再評価する前提となる workbook binding / manifest 方針を整理する。
- その binding が入ったときに `sheetName + shapeName` lookup helper の入力をどう拡張するかを設計する。
