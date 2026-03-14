# Worksheet / Chart Control Collection Feasibility

## 結論

- `DialogSheet` で導入した literal-only selector 正規化は、2026-03-14 時点では `Worksheet` / `Chart` へそのまま横展開しない。
- `Worksheet.Buttons` / `CheckBoxes` / `OptionButtons` と `Chart.Buttons` / `CheckBoxes` / `OptionButtons` は、Office VBA の正本ではなく .NET interop 側にある補助ソース候補として扱う。
- worksheet / chart sheet 上の ActiveX control について、Office VBA の正本道線は `OLEObjects`、`Shapes`、control code name であり、`Buttons` 系 collection method を先に user-facing に出す優先度は低い。
- 将来導入する場合でも、`DialogSheet` と同じく `Count` / `Item` の最小公開と literal-only selector 正規化を前提にし、`Add` / `Group` / `Duplicate` のような変更系 member は抑止を維持する。
- 次段は `Worksheet` / `Chart` の `OLEObjects` / control name 導線の整理を優先し、`Buttons` 系 collection は docs 段階に留める。

## 確認した公式ソース

### Office VBA

- [Using ActiveX Controls on Sheets](https://learn.microsoft.com/office/vba/excel/concepts/controls-dialogboxes-forms/using-activex-controls-on-sheets)
  - worksheet と chart sheet の ActiveX control は `OLEObjects` collection と `Shapes` collection で扱うと明記されている。
  - 「Most often, your Visual Basic code will refer to ActiveX controls by name.」として、`Sheet1.CommandButton1.Caption` のような control code name 導線が例示されている。
  - `OLEObjects("CommandButton1").Object.Caption` のように、collection 経由では shape name と code name の違いも含めて扱う例が示されている。
- [Worksheet.OLEObjects method (Excel)](https://learn.microsoft.com/office/vba/api/excel.worksheet.oleobjects)
  - `Worksheet.OLEObjects(Index)` は単体 `OLEObject` または `OLEObjects` collection を返し、name / number selector を Office VBA の正本として提供している。
- [Chart.OLEObjects method (Excel)](https://learn.microsoft.com/office/vba/api/excel.chart.oleobjects)
  - chart sheet でも `OLEObjects(Index)` が同様に用意されている。

### Microsoft Learn .NET interop

- [WorksheetClass.Buttons(Object)](https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.worksheetclass.buttons?view=excel-pia)
- [WorksheetClass.CheckBoxes(Object)](https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.worksheetclass.checkboxes?view=excel-pia)
- [WorksheetClass.OptionButtons(Object)](https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.worksheetclass.optionbuttons?view=excel-pia)
- [ChartClass.Buttons(Object)](https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.chartclass.buttons?view=excel-pia)
- [ChartClass.CheckBoxes(Object)](https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.chartclass.checkboxes?view=excel-pia)
- [ChartClass.OptionButtons(Object)](https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.chartclass.optionbuttons?view=excel-pia)
  - いずれも `Public Overridable Function ...(Optional Index As Object) As Object` で、selector あり無しにかかわらず戻り値は `Object` である。

## 観察結果

### 1. Office VBA の user-facing 導線は `Buttons` 系ではない

- worksheet / chart sheet の ActiveX control は、Office VBA の正本では `OLEObjects` と `Shapes`、および control code name で説明されている。
- `Sheet1.CommandButton1.Caption`、`Worksheets(1).OLEObjects("CommandButton1").Object.Caption`、`For Each s In Worksheets(1).Shapes` のような例が揃っており、`Worksheet.Buttons(1)` ではない。
- そのため `Buttons` / `CheckBoxes` / `OptionButtons` を先に補完 surface へ出すと、公式導線より interop 由来の補助導線を優先する形になる。

### 2. `Buttons` / `CheckBoxes` / `OptionButtons` は host object ごとに同じ曖昧さを持つ

- `DialogSheet` と同様に、`WorksheetClass.Buttons(Object)` や `ChartClass.OptionButtons(Object)` も `Optional Index As Object -> As Object` である。
- product 側で user-facing owner を決めるには、少なくとも以下の分岐が必要になる。
  - 引数省略時は collection owner
  - literal selector のみ item owner
  - expression selector / grouped selector は collection owner のまま維持
- これは `DialogSheet` で入れた literal-only 正規化と同じ系統だが、host object を増やすと surface とテストが大きく広がる。

### 3. worksheet / chart sheet では code name と shape name のずれも問題になる

- Office VBA では、event procedure は control code name を使う一方、`Shapes` / `OLEObjects` から name で引くときは shape name を使う必要がある。
- つまり `Buttons("CheckBox1")` のような selector を補完で支援しても、VBA 利用者が実際に知りたいのは code name か shape name かで揺れる。
- 先に `Buttons` 系 collection を足すより、`Sheet1.CommandButton1` と `OLEObjects("ShapeName").Object` のどちらを product として支援するかを整理する方が筋がよい。

### 4. collection owner の変更系 member は誤補完リスクが高い

- interop collection page は `Add` / `Group` / `Duplicate` など変更系 member を持ちやすい。
- `DialogSheet` では `Count` / `Item` だけに絞ることで誤補完を抑えており、`Worksheet` / `Chart` でも同じ抑制が前提になる。
- ただし Office VBA では ActiveX control の生成・操作に `OLEObjects.Add` や `Shapes` の方が明示されているため、`Buttons.Add` まで出す意義はさらに低い。

## 推奨方針

- `Worksheet` / `Chart` の `Buttons` / `CheckBoxes` / `OptionButtons` は、当面 supplemental interop source の導入候補として docs に留める。
- `DialogSheet` で作った literal-only selector 正規化 helper は将来再利用できる前提を維持するが、現時点では host object を増やさない。
- `Worksheet` / `Chart` で先に整理すべき対象は以下とする。
  - `Sheet1.CommandButton1` のような control code name 導線
  - `Worksheet.OLEObjects(Index)` / `Chart.OLEObjects(Index)` の chain 解決
  - `Shapes` / `OLEObjects` で shape name と code name がずれるケースの扱い
- もし `Buttons` 系を将来導入する場合でも、初回は以下を同時に固定する。
  - `Count` / `Item` だけの最小 collection surface
  - literal-only selector 正規化
  - expression selector / grouped selector の抑止
  - `Add` / `Group` / `Duplicate` の未公開維持
  - `Worksheet` / `Chart` / `DialogSheet` で共通に使う監査テスト

## 実装へ進む前に見るべき点

- `Sheet1.CommandButton1` を built-in member として扱う場合、VBComponent 情報だけで control code name を得られるか
- `OLEObjects("ShapeName").Object` を辿る場合、`.Object` の先を既知 control type へ安全に落とせるか
- chart sheet 上の control を `Chart` root と `OLEObjects` root のどちらで優先表示するか
- `Buttons("ControlName")` と `OLEObjects("ShapeName").Object` の両方を出したとき、どちらが誤案内になりやすいか
