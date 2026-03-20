# Explicit Sheet-Name Broad Root Feasibility

## 結論

- `ActiveWorkbook.Worksheets("Sheet1")` と unqualified `Worksheets("Sheet1")` は、current bundle の sidecar へ静的に結ばない。
- この非公開境界は `OLEObject.Object` と `Shape.OLEFormat.Object` の両方でそろえて維持する。
- user-facing に開く explicit sheet-name root は、引き続き `ThisWorkbook.Worksheets("Sheet1")` のように workbook identity を静的に固定できる経路に限る。
- broad root を再評価する条件は、正本 [workbook-binding-manifest-feasibility.md](./workbook-binding-manifest-feasibility.md) と [active-workbook-identity-provider-contract.md](./active-workbook-identity-provider-contract.md) で定義する workbook binding / host identity 契約が揃ったときだけとする。

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
- [Workbook.Worksheets property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.worksheets)
  - object qualifier 無しの `Worksheets` は active workbook を対象にする。
- [Refer to Sheets by Name](https://learn.microsoft.com/office/vba/excel/concepts/workbooks-and-worksheets/refer-to-sheets-by-name)
  - `Worksheets("Sheet1")` の string selector は active workbook 上の sheet name を指す。
- [Workbook object (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook)
  - `ThisWorkbook` は「コードが存在する workbook」、`ActiveWorkbook` は「現在 active な workbook」であり、特に add-in では一致しない。

## 現行実装と sidecar の前提

- current product の sidecar lookup は bundle-local artifact `/.vba/worksheet-control-metadata.json` を対象にする。
- `ThisWorkbook.Worksheets("Sheet One")` 経路では workbook root identity を保持し、current bundle の sidecar から `sheetName + shapeName` を引ける。
- `ActiveWorkbook.Worksheets("Sheet One")`、unqualified `Worksheets("Sheet One")`、`ActiveSheet` は保守動作として未解決のまま固定している。
- `Sheet1.Shapes("CheckBox1").OLEFormat.Object` と `Sheet1.OLEObjects("CheckBox1").Object` は、document module alias 起点の bundle identity を持てるため user-facing に解決している。

## 観察結果

### 1. `ActiveWorkbook` と unqualified `Worksheets` は current bundle を指さない

- Office VBA の正本では、`ActiveWorkbook` は active window 上の workbook であり、コードを含む workbook とは定義されていない。
- `Worksheets("Sheet1")` と `Application.Worksheets("Sheet1")` は active workbook を対象にする。
- したがって `ActiveWorkbook.Worksheets("Sheet1")` と unqualified `Worksheets("Sheet1")` を current bundle の sidecar に静的に結ぶと、公式の意味論とずれる。

### 2. `ThisWorkbook` だけが「コードを含む workbook」を静的に固定できる

- `Application.ThisWorkbook` の正本は、add-in を含めて「current macro code が動いている workbook」を返すと明記している。
- 一方 `ActiveWorkbook` は add-in workbook を返さず、呼び出し元 workbook を返し得る。
- current product は loose file / sidecar を bundle 単位で管理しているため、静的解析で current bundle を指してよい根拠があるのは `ThisWorkbook` 経路だけである。

### 3. active workbook は runtime state であり、静的解析では固定できない

- active workbook は window focus と user 操作で変化し得る。
- workbook が複数 window で開かれるケースや add-in 呼び出しでは、編集中 bundle と runtime active workbook が一致するとは限らない。
- sidecar は compile-time artifact であり、runtime state に追従する仕組みを持たない。

### 4. broad root は `OLEObject.Object` と `Shape.OLEFormat.Object` でそろえるべき

- どちらの経路も最終的には同じ `sheetName + shapeName -> controlType` lookup に依存する。
- 片側だけ broad root を開くと、「同じ worksheet root なのに `OLEObjects` は解決するが `Shapes` は解決しない」またはその逆、という user-facing 不整合が出る。
- したがって broad root の可否は両経路で同時に判断し、現段階では両方とも閉じる。

### 5. broad root を開けるのは workbook binding が明示化された後だけである

- 例えば次のいずれかが追加されれば再評価余地がある。
  - bundle-local `workbook-binding.json`
  - host から active workbook `FullName` を受け取る契約
- これらが無い限り、broad root を user-facing にすると誤補完リスクが残る。

## 推奨方針

### 現時点で維持する境界

- user-facing:
  - `ThisWorkbook.Worksheets("Sheet1").OLEObjects("ShapeName").Object`
  - `ThisWorkbook.Worksheets("Sheet1").Shapes("ShapeName").OLEFormat.Object`
- 非 user-facing:
  - `ActiveWorkbook.Worksheets("Sheet1").OLEObjects("ShapeName").Object`
  - `ActiveWorkbook.Worksheets("Sheet1").Shapes("ShapeName").OLEFormat.Object`
  - `Worksheets("Sheet1").OLEObjects("ShapeName").Object`
  - `Worksheets("Sheet1").Shapes("ShapeName").OLEFormat.Object`

### 再評価の条件

- broad root を再評価するときは、`current bundle == target workbook` を静的または明示設定で保証できる仕組みを先に導入する。
- そのときも `sheetName` を join key に使う点と、`sheetCodeName` を document module alias / control code name 導線へ分ける点は維持する。
- `OLEObject.Object` と `Shape.OLEFormat.Object` の境界は同じ PR で動かす。

## 今回の完了条件

- `ActiveWorkbook.Worksheets("Sheet1")` と unqualified `Worksheets("Sheet1")` を current bundle の sidecar に静的接続しない理由を一次情報ベースで整理する。
- `ThisWorkbook` 限定を当面の user-facing 境界として固定する。
- broad root の可否を `OLEObject.Object` / `Shape.OLEFormat.Object` 共通の判断として整理する。
- broad root 再評価に必要な前提を docs に残す。

## 次段の候補

- `available` snapshot と manifest match がそろったときだけ、`ActiveWorkbook.Worksheets("SheetName")` broad root を current bundle sidecar lookup へ限定接続する。
- `OLEObject.Object` / `Shape.OLEFormat.Object` の両方で broad root の gating 条件と負例をそろえる。
