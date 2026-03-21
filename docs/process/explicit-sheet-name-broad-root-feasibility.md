# Explicit Sheet-Name Broad Root Feasibility

## 結論

- `ActiveWorkbook.Worksheets("SheetName")` は、`available` snapshot、manifest 存在、manifest match、対応 owner の 4 条件がそろったときだけ current bundle sidecar lookup を開いてよい。この条件は現行実装で user-facing に有効化済みである。
- unqualified `Worksheets("SheetName")` と `Application.Worksheets("SheetName")` は Office VBA 上で active workbook を対象にするため、静的 current bundle root ではないが、broad root gating 条件は `ActiveWorkbook.Worksheets("SheetName")` と同一 family として扱ってよい。
- broad root family の対象構文は `Worksheets("literal sheetName")` / `Worksheets.Item("literal sheetName")` と `Application.Worksheets("literal sheetName")` / `Application.Worksheets.Item("literal sheetName")` を同一扱いにし、`available` snapshot と manifest match がそろったときだけ sidecar lookup を開く。
- built-in broad root gating は `Worksheets` root が built-in collection として解決できた場合にだけ適用し、同名の変数、関数、メンバーへ shadow されているときは user-defined symbol を優先する。
- `OLEObject.Object` と `Shape.OLEFormat.Object` の broad root 境界は同じ PR、同じ条件で動かす。

## 確認した公式ソース

### Office VBA

- [Application.ThisWorkbook property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.thisworkbook)
  - `ThisWorkbook` は「current macro code が動いている workbook」を返す。
  - add-in では `ActiveWorkbook` は add-in workbook ではなく、呼び出し元 workbook を返す。
- [Application.ActiveWorkbook property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.activeworkbook)
  - `ActiveWorkbook` は active window の workbook を返す。
  - active window が無い場合や Protected View では `Nothing` になり得る。
- [Application.Worksheets property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.worksheets)
  - object qualifier 無しの `Worksheets` は active workbook の worksheet collection を返す。
- [Worksheets.Item property (Excel)](https://learn.microsoft.com/office/vba/api/excel.worksheets.item)
  - `Item` は collection から単一 object を返し、default member として `ActiveWorkbook.Worksheets.Item(1)` と `ActiveWorkbook.Worksheets(1)` は等価である。
- [Returning an Object from a Collection (Excel)](https://learn.microsoft.com/office/vba/excel/concepts/workbooks-and-worksheets/returning-an-object-from-a-collection-excel)
  - collection の `Item` は既定メンバーであり、省略形と同じ object を返す。
- [Application.Sheets property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.sheets)
  - object qualifier 無しの `Sheets` は `ActiveWorkbook.Sheets` と等価であり、worksheet 以外も混在する。
- [Workbook.Worksheets property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.worksheets)
  - object qualifier 無しの `Worksheets` は active workbook を対象にする。
- [Refer to Sheets by Name](https://learn.microsoft.com/office/vba/excel/concepts/workbooks-and-worksheets/refer-to-sheets-by-name)
  - `Worksheets("Sheet1")` の string selector は active workbook 上の sheet name を指す。
- [Worksheet object (Excel)](https://learn.microsoft.com/office/vba/api/excel.worksheet)
  - `Worksheets(index)` の string selector は worksheet 名であり、numeric selector も許す。
- [Workbook object (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook)
  - `ThisWorkbook` は「コードが存在する workbook」、`ActiveWorkbook` は「現在 active な workbook」であり、特に add-in では一致しない。

## 現行実装と sidecar の前提

- current product の sidecar lookup は bundle-local artifact `/.vba/worksheet-control-metadata.json` を対象にする。
- `ThisWorkbook.Worksheets("Sheet One")` 経路では workbook root identity を保持し、current bundle の sidecar から `sheetName + shapeName` を引ける。
- `ActiveWorkbook.Worksheets("Sheet One")` は workbook binding manifest と active workbook snapshot が一致したときだけ、`OLEObject.Object` / `Shape.OLEFormat.Object` の既存 worksheet control owner 導線へ進める。
- unqualified `Worksheets("Sheet One")`、`Application.Worksheets("Sheet One")`、`ActiveSheet` は現時点では保守動作として未解決のまま固定している。
- `Sheet1.Shapes("CheckBox1").OLEFormat.Object` と `Sheet1.OLEObjects("CheckBox1").Object` は、document module alias 起点の bundle identity を持てるため user-facing に解決している。
- unqualified `Worksheets(1)` / `Worksheets(i + 1)` は built-in `Worksheet` surface までは既に user-facing だが、`sheetName + shapeName` lookup を要する control owner 昇格には使っていない。

## 観察結果

### 1. `ActiveWorkbook` と unqualified `Worksheets` は current bundle を静的には指さない

- Office VBA の正本では、`ActiveWorkbook` は active window 上の workbook であり、コードを含む workbook とは定義されていない。
- `Worksheets("Sheet1")` と `Application.Worksheets("Sheet1")` は active workbook を対象にする。
- したがって `ActiveWorkbook.Worksheets("Sheet1")` と unqualified `Worksheets("Sheet1")` を current bundle の sidecar に静的に結ぶと、公式の意味論とずれる。

### 2. `ActiveWorkbook` と unqualified `Worksheets` は同じ runtime gating family にできる

- `Application.Worksheets` の正本は、object qualifier 無しの `Worksheets` が active workbook の worksheet collection を返すと明記している。
- `Worksheets.Item` の正本は `Item` が既定メンバーであると明記しているため、`Worksheets.Item("Sheet1")` と `Application.Worksheets.Item("Sheet1")` も direct call form と同じ worksheet selector とみなしてよい。
- したがって `Worksheets("Sheet1")` / `Worksheets.Item("Sheet1")` と `Application.Worksheets("Sheet1")` / `Application.Worksheets.Item("Sheet1")` は、user-facing に開くなら `ActiveWorkbook.Worksheets("Sheet1")` と同じ runtime 条件で開くべきである。
- root の書き方だけで gating 条件が変わると、「同じ active workbook root なのに `ActiveWorkbook` では解決し unqualified では解決しない」という docs / 実装の不整合が起きやすい。

### 3. `ThisWorkbook` だけが「コードを含む workbook」を静的に固定できる

- `Application.ThisWorkbook` の正本は、add-in を含めて「current macro code が動いている workbook」を返すと明記している。
- 一方 `ActiveWorkbook` と unqualified `Worksheets` は active workbook family であり、runtime state を挟まない限り current bundle を指してよい根拠が無い。
- current product は loose file / sidecar を bundle 単位で管理しているため、静的 current bundle root と runtime active-workbook root を分けて扱う必要がある。

### 4. broad root は `OLEObject.Object` と `Shape.OLEFormat.Object` でそろえるべき

- どちらの経路も最終的には同じ `sheetName + shapeName -> controlType` lookup に依存する。
- 片側だけ broad root を開くと、「同じ worksheet root なのに `OLEObjects` は解決するが `Shapes` は解決しない」またはその逆、という user-facing 不整合が出る。
- したがって broad root の可否は両経路で同時に判断し、`ActiveWorkbook` と unqualified `Worksheets` でも同じ条件・同じ PR でそろえるべきである。

### 5. `Sheets` / `ActiveSheet` / grouped / numeric / dynamic selector は broad root family に混ぜない

- `Sheets` は worksheet だけでなく chart / dialog / module sheet も混在し、`Worksheet` 固定の sidecar join key と一致しない。
- `ActiveSheet` は workbook と sheet の両方が runtime state 依存であり、sheetName literal も持たない。
- `Worksheets(1)` / `Worksheets(i + 1)` / `Worksheets(Array(...))` は、generic `Worksheet` surface までは扱えても、control owner 昇格に必要な stable `sheetName` を compile time に復元できない。

### 6. shadow される `Worksheets` は broad root gating の対象外にするべきである

- unqualified `Worksheets` は explicit `ActiveWorkbook.Worksheets` よりも、同名の変数、関数、メンバーと衝突しやすい。
- そのため broad root gating は、root identifier が built-in `Worksheets` collection として解決できるときにだけ適用し、workspace symbol が優先される場合は sidecar lookup へ進めない方が安全である。

## 推奨方針

### 現時点で user-facing な経路

- user-facing:
  - `ThisWorkbook.Worksheets("Sheet1").OLEObjects("ShapeName").Object`
  - `ThisWorkbook.Worksheets("Sheet1").Shapes("ShapeName").OLEFormat.Object`
  - `ActiveWorkbook.Worksheets("Sheet1").OLEObjects("ShapeName").Object`
  - `ActiveWorkbook.Worksheets("Sheet1").Shapes("ShapeName").OLEFormat.Object`
  - `Worksheets("Sheet1").OLEObjects("ShapeName").Object`
  - `Worksheets.Item("Sheet1").OLEObjects("ShapeName").Object`
  - `Application.Worksheets("Sheet1").Shapes("ShapeName").OLEFormat.Object`
  - `Application.Worksheets.Item("Sheet1").Shapes("ShapeName").OLEFormat.Object`
- 非 user-facing:
  - `Sheets("Sheet1").OLEObjects("ShapeName").Object`
  - `ActiveSheet.OLEObjects("ShapeName").Object`
  - `Worksheets(1).OLEObjects("ShapeName").Object`
  - `Worksheets(GetSheetName()).OLEObjects("ShapeName").Object`
  - `Worksheets(Array("Sheet1")).OLEObjects("ShapeName").Object`

### unqualified broad root を開く条件

- `available` snapshot
- current bundle の `workbook-binding.json` 存在
- manifest と active workbook snapshot の match
- root が built-in `Worksheets` collection として解決できる
- path が `Worksheets("literal sheetName")` / `Worksheets.Item("literal sheetName")` または `Application.Worksheets("literal sheetName")` / `Application.Worksheets.Item("literal sheetName")`
- `OLEObject.Object` / `Shape.OLEFormat.Object` を同じ PR・同じ条件で開く

### 維持する除外境界

- `Sheets`
- `ActiveSheet`
- numeric selector
- dynamic selector
- grouped selector
- `Worksheets` の shadow case

## 今回の完了条件

- `ActiveWorkbook.Worksheets("Sheet1")` は runtime gating で user-facing に開く一方、unqualified `Worksheets("Sheet1")` は別判断であることを docs 上で矛盾なく整理する。
- unqualified `Worksheets("Sheet1")` と `Application.Worksheets("Sheet1")` を same broad-root family として扱う条件を、一次情報ベースで整理する。
- `Sheets` / `ActiveSheet` / grouped / numeric / dynamic selector / shadow case を broad root family から除外する理由を残す。
- broad root の可否を `OLEObject.Object` / `Shape.OLEFormat.Object` 共通の判断として整理する。

## 次段の候補

- `available` snapshot と manifest match がそろったときだけ、`Worksheets("SheetName")` と `Application.Worksheets("SheetName")` broad root を current bundle sidecar lookup へ限定接続する。
- `OLEObject.Object` / `Shape.OLEFormat.Object` の両方で broad root の gating 条件、shadow case、非対象 selector の負例をそろえる。
