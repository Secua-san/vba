# Shape.OLEFormat.Object Promotion Feasibility

## 結論

- 現行実装では、`Shape.OLEFormat.Object` の先を control owner へ昇格する条件を `worksheet document module root + shape name string literal + sidecar 一致` に限定している。
- `Shape.Type = msoOLEControlObject` は runtime では有効な判定だが、解析時にその値は直接取れない。代わりに、worksheet control metadata sidecar の provenance を静的な根拠として使う。
- `Sheet1.Shapes("CheckBox1").OLEFormat.Object` と `Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object` は、現行 product で user-facing に解決する。
- ただし `Shapes(1)` / `Shapes.Item(1)`、dynamic selector、`Chart1` root、`ShapeRange` / grouped selector、code name ベース access は昇格条件から外す。

## 確認した公式ソース

### Office VBA

- [Using ActiveX Controls on Sheets](https://learn.microsoft.com/office/vba/excel/concepts/controls-dialogboxes-forms/using-activex-controls-on-sheets)
  - ActiveX control は `OLEObjects` collection の `OLEObject` として表され、同時に `Shapes` collection の member でもある。
  - control property の実体には `OLEObject.Object` で進める。
  - `Shapes` / `OLEObjects` collection で control を名前指定するときは code name ではなく shape name を使う。
  - `For Each s In Worksheets(1).Shapes : If s.Type = msoOLEControlObject Then ...` の例があり、`msoOLEControlObject` が runtime 側の control 判定である。
- [Shapes object (Excel)](https://learn.microsoft.com/office/vba/api/excel.shapes)
  - `Shapes(index)` は単一 `Shape` を返す。
  - `Shapes.Range(index)` は単一 `Shape` ではなく `ShapeRange` を返す。
  - `Shapes` は drawing layer 全体であり、AutoShape、OLE object、picture などを含む。
- [Shape.OLEFormat property (Excel)](https://learn.microsoft.com/office/vba/api/excel.shape.oleformat)
  - `Shape.OLEFormat` は OLE object properties を返す。
  - `Shapes(1)` が embedded OLE object でない場合、この access は失敗する。
- [OLEFormat.Object property (Excel)](https://learn.microsoft.com/office/vba/api/excel.oleformat.object)
  - `OLEFormat.Object` の戻り値は read-only `Object` であり、example でも embedded Word document object を返している。
- [Shape.Type property (Excel)](https://learn.microsoft.com/office/vba/api/excel.shape.type)
- [MsoShapeType enumeration (Office)](https://learn.microsoft.com/office/vba/api/office.msoshapetype)
  - `msoOLEControlObject`、`msoEmbeddedOLEObject`、`msoLinkedOLEObject` は区別される。
- [ShapeRange object (Excel)](https://learn.microsoft.com/office/vba/api/excel.shaperange)
  - `Shapes.Range(Array(...))` は subset を `ShapeRange` として返し、単一 `Shape` と同一視できない。

## 現行実装と sidecar の前提

- 現行 product は `Shapes(Index)` / `Shapes.Item(Index)` を generic `Shape` owner へ進め、`Shape.OLEFormat.Object` は `Sheet1.Shapes("shapeName")` / `Sheet1.Shapes.Item("shapeName")` の string literal selector にだけ限定して control owner へ進める。
- sidecar v1 には supported worksheet owner ごとに `sheetCodeName`、`shapeName`、`codeName`、`controlType`、`progId` / `classId` が入る。
- sidecar generator は未知 `controlType` を fail-fast にしており、`controlType` が書かれた record は既知 control owner へ正規化済みである。
- `shapeId` は Open XML drawing 側の identifier であり、`Shapes(1)` の collection index とは別物である。したがって numeric selector を sidecar に結び付ける根拠にはできない。

## 観察結果

### 1. `Shape.Type = msoOLEControlObject` は runtime 条件であって、静的 resolver 条件ではない

- Office VBA の正本では、`Shapes` から control を選り分ける条件は `s.Type = msoOLEControlObject` で表現される。
- ただし静的解析時には `Shape.Type` の実行結果は分からないため、この条件をそのまま resolver に持ち込めない。
- 将来の実装では、workbook package から生成した sidecar が「この shape name は ActiveX control である」という静的な証拠として機能する。

### 2. `OLEFormat.Object` は embedded document と control を区別しない

- `OLEFormat.Object` は常に `Object` を返し、公式 example でも embedded Word document object を返している。
- `Shape.OLEFormat` 自体も OLE object でない場合は失敗するため、`Shape` から無条件に `.Object` の先を control owner に進めるのは誤補完になる。
- よって「`Shape` path で sidecar record が存在すること」を昇格の前提にしない限り、`CheckBox.Value` のような member は出せない。

### 3. shape name string literal は sidecar と結合できるが、numeric / dynamic selector はできない

- sidecar の join key は `shapeName` であり、`Shapes("CheckBox1")` / `Shapes.Item("CheckBox1")` はこの key と直接結び付けられる。
- 一方、`Shapes(1)` / `Shapes.Item(1)` は collection index であり、sidecar が保持する `shapeId` と同義ではない。
- `Shapes(GetIndex())` や `Shapes.Item(GetIndex())` のような dynamic selector も compile time で `shapeName` が確定しないため、昇格条件から外すべきである。

### 4. `ShapeRange` / grouped selector は単一 control owner へ落とさない

- `Shapes.Range(Array(...))` は `ShapeRange` を返し、単一 `Shape` と同じ path ではない。
- `ShapeRange` は複数 shape を含み得るため、たとえ 1 要素 array でも `CheckBox` や `CommandButton` のような単一 owner へ直結しない方が安全である。
- grouped selector を user-facing に出すなら、`ShapeRange` surface 自体を別 task として扱うべきである。

### 5. chartsheet と code name 導線は別の問題として閉じる

- chartsheet owner は current sidecar / probe で inventory が未確立のため、`Chart1.Shapes("CheckBox1").OLEFormat.Object` は昇格対象にしない。
- `Sheet1.chkFinished` は code name 導線であり、shape name literal を使う `Shapes("CheckBox1")` とは join key が異なる。
- したがって `Shape.OLEFormat.Object` の昇格条件には code name を混ぜず、shape name path と control code name path を分けて扱う。

## 実装した最小昇格条件

- root が explicit な worksheet document module alias (`Sheet1`) に解決できること
- root owner が sidecar 上で `ownerKind: "worksheet"` かつ `status: "supported"` であること
- path が `Shapes("shapeName").OLEFormat.Object` または `Shapes.Item("shapeName").OLEFormat.Object` であること
- selector が string literal であり、`shapeName` を compile time に復元できること
- sidecar 上に同じ `shapeName` を持つ control record があり、`controlType` が既知 owner に正規化済みであること

## 将来も除外する条件

- `Shapes(1)` / `Shapes.Item(1)` の numeric selector
- `Shapes(GetIndex())` / `Shapes.Item(GetIndex())` の dynamic selector
- `Shapes.Range(Array(...))` / `ShapeRange` / grouped selector
- `Chart1` や `ActiveChart` のような chartsheet root
- `ActiveSheet` / `Worksheets(1)` のように explicit document module identity が取れない root
- `Sheet1.chkFinished` のような code name 導線との混用

## 2026-03-15 時点の完了状態

- `worksheet document module + shape name string literal + sidecar 一致` の path を実装し、`CheckBox.Value` / `Select` の completion / hover / signature help / semantic token を user-facing にした。
- numeric / dynamic / chart root / `ShapeRange` の `Shape.OLEFormat.Object` 非昇格を server / extension test で回帰固定した。
- `ShapeRange` / grouped selector は引き続き docs 上で別 task 扱いに切り分けている。

## 次段の候補

- explicit sheet-name root の整理と `ThisWorkbook.Worksheets("Sheet1")` 限定実装は正本 [shape-oleformat-object-explicit-sheet-root-feasibility.md](./shape-oleformat-object-explicit-sheet-root-feasibility.md) を参照する。
- 次の候補は `ActiveWorkbook.Worksheets("Sheet1")` / unqualified `Worksheets("Sheet1")` の broad root 展開可否を整理し、`current bundle` と runtime active workbook の同一視前提を user-facing に許容するかを決める。
