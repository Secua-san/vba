# Worksheet / Chart Shapes Root Feasibility

## 結論

- `Worksheet.Shapes(Index)` / `Chart.Shapes(Index)` と `Shapes.Item(Index)` は、collection 全体ではなく単一 `Shape` owner へ正規化して user-facing に出す。
- ただし `Shapes` root は control 専用 root にはしない。`Shape.OLEFormat` までは generic な `Shape` / `OLEFormat` surface として扱い、`Shape.OLEFormat.Object` の先を `CheckBox` や `CommandButton` へ昇格させない。
- `Shapes("CheckBox1")` の selector は shape name 前提で扱い、control code name とは結び付けない。`Sheet1.chkFinished` は別導線として維持する。
- `Shape.Type = msoOLEControlObject` のような実行時判定と sidecar の `shapeName -> controlType` を組み合わせる設計余地はあるが、現段階では解析時に安全に確定できないため後続へ送る。

## 確認した公式ソース

### Office VBA

- [Shapes object (Excel)](https://learn.microsoft.com/office/vba/api/excel.shapes)
  - `Shapes(index)` は shape name または index number から単一 `Shape` object を返す。
  - `Shapes.Range(Array(...))` は subset を `ShapeRange` として返す。
  - ActiveX control には shape name と code name の 2 つがあり、`Shapes` / `OLEObjects` collection から name で引くときは shape name を使う。
- [Shapes.Item method (Excel)](https://learn.microsoft.com/office/vba/api/excel.shapes.item)
  - `Item(Index)` も単一 `Shape` object を返す。
- [Using ActiveX Controls on Sheets](https://learn.microsoft.com/office/vba/excel/concepts/controls-dialogboxes-forms/using-activex-controls-on-sheets)
  - most often は `Sheet1.CommandButton1` のような code name 導線だが、`Shapes` / `OLEObjects` collection からは shape name を使う。
  - `For Each s In Worksheets(1).Shapes : If s.Type = msoOLEControlObject Then ...` の例があり、`Shapes` は drawing object 全体から control を絞り込む入口として扱われている。
- [Shape.Type property (Excel)](https://learn.microsoft.com/office/vba/api/excel.shape.type)
- [MsoShapeType enumeration (Office)](https://learn.microsoft.com/office/vba/api/office.msoshapetype)
  - `msoOLEControlObject`、`msoEmbeddedOLEObject`、`msoLinkedOLEObject` が区別される。
- [Shape.OLEFormat property (Excel)](https://learn.microsoft.com/office/vba/api/excel.shape.oleformat)
- [OLEFormat object (Excel)](https://learn.microsoft.com/office/vba/api/excel.oleformat)
  - `Shape` が linked / embedded OLE object でない場合は `OLEFormat` が失敗する。
  - `OLEFormat.Object` は generic な `Object` であり、埋め込みアプリ object を返し得る。

## 観察結果

### 1. `Shapes(Index)` と `Shapes.Item(Index)` は `Shape` へ正規化してよい

- Office VBA の正本はどちらも単一 `Shape` object を返す。
- そのため、`Sheet1.Shapes(1).Name` や `Chart1.Shapes("CheckBox1").Left` のような generic `Shape` member 補完までは、sidecar に依存せず user-facing に出せる。
- `Shapes.Range(Array(...))` は別途 `ShapeRange` の導線なので、単一 `Shape` とは切り分ける。

### 2. `Shapes` は control 専用 collection ではない

- `Shapes` は AutoShape、picture、chart、linked / embedded OLE object、OLE control object を含む drawing layer 全体を表す。
- 同じ shape name selector でも、control である保証は `Shape.Type` の実行時値や workbook metadata が無いと確定できない。
- したがって `Shapes("CheckBox1")` だからといって、直ちに `CheckBox` や `CommandButton` owner へ落とすのは誤補完リスクが高い。

### 3. `Shape.OLEFormat` は generic OLE surface に留めるのが安全

- Office VBA は `Shape.OLEFormat` を提供しているが、shape が OLE object でなければ失敗する。
- `OLEFormat.Object` の返り値は generic `Object` であり、control だけでなく embedded / linked document の top-level interface もあり得る。
- そのため現段階では、`Shapes("CheckBox1").OLEFormat.ProgID` のような generic `OLEFormat` member は出してよいが、`.Object` の先を sidecar だけで control owner へ進めない方が安全である。

### 4. shape name と code name は別導線として扱う

- `Shapes("CheckBox1")` は Office VBA 上も shape name selector であり、`Sheet1.chkFinished` のような code name 導線とは意味が異なる。
- 現行 product は worksheet document module root に限って `Sheet1.ControlCodeName` を sidecar で支援しているため、`Shapes` root まで同じ metadata を流し込むと導線の意味が混ざる。
- `Shapes` root は shape name ベース、code name は direct access ベース、という分離を維持する。

## 実装方針

- `Sheet1.Shapes.` / `Chart1.Shapes.` は `Shapes` collection のまま扱う。
- `Sheet1.Shapes(1)` / `Sheet1.Shapes("CheckBox1")` / `Sheet1.Shapes(i + 1)` は `Shape` owner へ正規化する。
- `Sheet1.Shapes.Item(1)` / `.Item("CheckBox1")` / `.Item(i + 1)` も同様に `Shape` owner へ正規化する。
- `Shape.OLEFormat` は `OLEFormat` owner として補完対象に含めるが、`Shape.OLEFormat.Object` の先は generic `Object` のまま止める。
- worksheet control metadata sidecar は `OLEObject.Object` と `Sheet1.ControlCodeName` の導線だけに使い、`Shapes` path には接続しない。

## 今回の完了条件

- `Shapes(Index)` / `Shapes.Item(Index)` が `Shape` owner へ進むことを server / extension test で固定する。
- `Sheet1.Shapes("CheckBox1").OLEFormat.` までは generic `OLEFormat` member が出ることを確認する。
- `Sheet1.Shapes("CheckBox1").OLEFormat.Object.` では `CheckBox.Value` のような control-specific member が出ないことを確認する。
- `shape name != code name` と `msoOLEControlObject` の論点は docs に残し、後続タスクへ切り出す。

## 次段の候補

- `Shape.OLEFormat.Object` の昇格条件整理は正本 [shape-oleformat-object-promotion-feasibility.md](./shape-oleformat-object-promotion-feasibility.md) へ移した。
- `msoEmbeddedOLEObject` / `msoLinkedOLEObject` は generic OLE object のまま維持し、control-only path と分離する。
- `Shapes.Range(Array(...))` / `ShapeRange` の surface をどこまで出すかを別タスクで整理する。
