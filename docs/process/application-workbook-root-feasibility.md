# Application Workbook Root Feasibility

## 結論

- `Application.ThisWorkbook.Worksheets("SheetName")` と `Application.ThisWorkbook.Worksheets.Item("SheetName")` は、`ThisWorkbook.Worksheets("SheetName")` / `.Item("SheetName")` と同じ static current-bundle root family として扱ってよい。
- `Application.ActiveWorkbook.Worksheets("SheetName")` と `Application.ActiveWorkbook.Worksheets.Item("SheetName")` は、`ActiveWorkbook.Worksheets("SheetName")` / `.Item("SheetName")` と同じ active-workbook broad-root family として扱ってよい。
- `Application` qualifier を付けても workbook identity の意味は変わらないため、`OLEObject.Object` と `Shape.OLEFormat.Object` は direct root と同じ条件で同時に開閉するべきである。
- ただし `Application` 自体が user-defined symbol に shadow される場合は built-in family に入れず、sidecar lookup は無効のまま維持する。
- 今回は policy の整理までとし、user-facing の最小接続は後続タスクへ分離する。

## 確認した公式ソース

### Office VBA

- [Application.ThisWorkbook property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.thisworkbook)
  - `ThisWorkbook` は current macro code が動いている workbook を返す。
  - add-in では `ActiveWorkbook` は add-in workbook ではなく caller workbook を返す。
- [Application.ActiveWorkbook property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.activeworkbook)
  - `ActiveWorkbook` は active window の workbook を返す。
  - active window が無い場合や Protected View では `Nothing` になり得る。
- [Application.Worksheets property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.worksheets)
  - object qualifier 無しの `Worksheets` は active workbook の worksheet collection を返す。
- [Workbook.Worksheets property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.worksheets)
  - workbook qualifier が付いた `Worksheets` はその workbook の worksheet collection を返す。
- [Worksheets.Item property (Excel)](https://learn.microsoft.com/office/vba/api/excel.worksheets.item)
  - `Item` は既定メンバーであり、`.Item("Sheet1")` と `("Sheet1")` は同じ worksheet selector を表す。

## 現行実装との対応

- user-facing 済み:
  - `ThisWorkbook.Worksheets("SheetName")` / `.Item("SheetName")`
  - `ActiveWorkbook.Worksheets("SheetName")` / `.Item("SheetName")`
  - unqualified `Worksheets("SheetName")` / `.Item("SheetName")`
  - `Application.Worksheets("SheetName")` / `.Item("SheetName")`
- policy は未整理だったが、未実装:
  - `Application.ThisWorkbook.Worksheets("SheetName")` / `.Item("SheetName")`
  - `Application.ActiveWorkbook.Worksheets("SheetName")` / `.Item("SheetName")`
- 既存 docs では `ThisWorkbook` と `ActiveWorkbook` の意味は整理済みであり、今回の論点は `Application.` qualifier を挟んでも workbook root identity が変わらないかどうかに限られる。

## 観察結果

### 1. `Application.ThisWorkbook` は `ThisWorkbook` と同じ current-bundle root である

- 正本は `Application.ThisWorkbook` を「current macro code が動いている workbook」と定義している。
- したがって `Application.ThisWorkbook.Worksheets("SheetName")` は、`ThisWorkbook.Worksheets("SheetName")` と同じ workbook identity を使って current bundle の sidecar lookup へ進めてよい。
- `Item("SheetName")` も既定メンバー規則により direct call form と同じ family に含めてよい。

### 2. `Application.ActiveWorkbook` は `ActiveWorkbook` と同じ broad-root family である

- `Application.ActiveWorkbook` は `ActiveWorkbook` の明示 qualifier 付き形であり、active workbook 以外の意味を追加しない。
- そのため `Application.ActiveWorkbook.Worksheets("SheetName")` / `.Item("SheetName")` を user-facing に開く条件は、既存 `ActiveWorkbook.Worksheets("SheetName")` / `.Item("SheetName")` と完全に同じでよい。
- 具体的には `available` snapshot、manifest 存在、manifest match、対応 owner の 4 条件がそろったときだけ sidecar lookup を開く。

### 3. `Application` qualifier は built-in 解決できるときだけ特別扱いする

- `Application.Worksheets("SheetName")` を broad-root family に含めたのと同じく、`Application.ThisWorkbook` / `Application.ActiveWorkbook` も built-in `Application` qualifier として解決できることが前提になる。
- `Application` が user-defined symbol に shadow されている場合、built-in workbook root family とみなすと user code の意味を壊す。
- したがって `Application` qualifier を含む sidecar lookup は、root identifier が built-in `Application` として解決できたときだけ有効化するべきである。

### 4. control owner 昇格条件は direct root とそろえるべきである

- `Application.ThisWorkbook...OLEObjects(...).Object` と `ThisWorkbook...OLEObjects(...).Object` は、最終的に同じ `sheetName + shapeName -> controlType` lookup を使う。
- `Application.ActiveWorkbook...Shapes(...).OLEFormat.Object` と `ActiveWorkbook...Shapes(...).OLEFormat.Object` も同様に、違うのは workbook root identity の決め方だけである。
- したがって `OLEObject.Object` と `Shape.OLEFormat.Object` は qualifier 有無で別ルールにせず、direct root と同じ条件で同じ PR にそろえて開くべきである。

### 5. 既存の除外境界はそのまま維持する

- `codeName` selector
- numeric selector
- dynamic selector
- grouped selector
- `Sheets`
- `ActiveSheet`
- chartsheet / unsupported owner
- `Application` shadow case

## 推奨方針

### static current-bundle family

- `ThisWorkbook.Worksheets("SheetName")`
- `ThisWorkbook.Worksheets.Item("SheetName")`
- `Application.ThisWorkbook.Worksheets("SheetName")`
- `Application.ThisWorkbook.Worksheets.Item("SheetName")`

### active-workbook broad-root family

- `ActiveWorkbook.Worksheets("SheetName")`
- `ActiveWorkbook.Worksheets.Item("SheetName")`
- `Application.ActiveWorkbook.Worksheets("SheetName")`
- `Application.ActiveWorkbook.Worksheets.Item("SheetName")`
- unqualified `Worksheets("SheetName")`
- unqualified `Worksheets.Item("SheetName")`
- `Application.Worksheets("SheetName")`
- `Application.Worksheets.Item("SheetName")`

### 開閉条件

- static current-bundle family:
  - `Application.ThisWorkbook` を含め manifest / snapshot 非依存
  - current bundle sidecar が見つかる
  - selector は literal `sheetName`
- active-workbook broad-root family:
  - `available` snapshot
  - manifest 存在
  - manifest match
  - root が built-in family として解決できる
  - selector は literal `sheetName`

## 今回の完了条件

- `Application.ThisWorkbook` を `ThisWorkbook` と同じ static current-bundle root family とみなせる根拠を整理する。
- `Application.ActiveWorkbook` を `ActiveWorkbook` と同じ broad-root family とみなせる根拠を整理する。
- `Application` shadow 時は built-in family に入れない境界を残す。
- `OLEObject.Object` と `Shape.OLEFormat.Object` を qualifier 有無で分けない方針を残す。

## 次段の候補

- `Application.ThisWorkbook.Worksheets("SheetName")` / `.Item("SheetName")` を `ThisWorkbook` direct root と同じ helper で最小接続する。
- `Application.ActiveWorkbook.Worksheets("SheetName")` / `.Item("SheetName")` を既存 broad-root gating helper へ寄せて最小接続する。
