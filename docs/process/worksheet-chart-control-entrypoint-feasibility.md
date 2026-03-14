# Worksheet / Chart Control Entry Point Feasibility

## 結論

- `Worksheet` / `Chart` 上の control 導線について、最初の実装候補は `Worksheet.OLEObjects(Index)` / `Chart.OLEObjects(Index)` を `OLEObject` owner へ落とす最小プロトタイプとする。
- `Sheet1.CommandButton1` は Office VBA の自然な導線だが、現行リポジトリの静的入力だけでは worksheet / chart sheet 上の control code name と型の inventory を安定して得られないため、初回対象から外す。
- `Shapes` は Office VBA の正本道線には含まれるが、control 以外の drawing object を広く含むため、最初の user-facing root にはしない。
- `OLEObject.Object` と `Shape.OLEFormat.Object` の先はどちらも `Object` で、control 種別や埋め込みアプリ型が固定できないため、初回は型付き chain 解決へ進めない。
- 次段は `Worksheet` / `Chart` の `OLEObjects` root だけを追加し、`OLEObject` surface の completion / hover / signature help へ到達できる状態を目標にする。

## 確認した公式ソース

### Office VBA

- [Using ActiveX Controls on Sheets](https://learn.microsoft.com/office/vba/excel/concepts/controls-dialogboxes-forms/using-activex-controls-on-sheets)
  - worksheet / chart sheet 上の ActiveX control は `OLEObjects` collection、`Shapes` collection、control code name で扱う。
  - `Sheet1.CommandButton1.Caption`、`Worksheets(1).OLEObjects("CommandButton1").Object.Caption`、`Worksheets(1).Shapes` の例が示されている。
- [Worksheet.OLEObjects method (Excel)](https://learn.microsoft.com/office/vba/api/excel.worksheet.oleobjects)
- [Chart.OLEObjects method (Excel)](https://learn.microsoft.com/office/vba/api/excel.chart.oleobjects)
  - いずれも name / number selector から単一 `OLEObject` または `OLEObjects` collection を返す。
- [OLEObjects.Item method (Excel)](https://learn.microsoft.com/office/vba/api/excel.oleobjects.item)
  - `Item(Index)` は name / index number で単一 object を返す。
- [OLEObject object (Excel)](https://learn.microsoft.com/office/vba/api/excel.oleobject)
  - `OLEObject` は ActiveX control または linked / embedded OLE object を表し、`Name`、`Top`、`Left`、`Visible`、`Activate`、`Select` など user-facing member を持つ。
- [OLEObject.Object property (Excel)](https://learn.microsoft.com/office/vba/api/excel.oleobject.object)
  - 返り値は `Object` で、例は embedded Word document を返している。
- [Worksheet.Shapes property (Excel)](https://learn.microsoft.com/office/vba/api/excel.worksheet.shapes)
- [Chart.Shapes property (Excel)](https://learn.microsoft.com/office/vba/api/excel.chart.shapes)
- [Shapes object (Excel)](https://learn.microsoft.com/office/vba/api/excel.shapes)
  - `Shapes` は sheet 上のすべての drawing object を含む。
- [Shape.OLEFormat property (Excel)](https://learn.microsoft.com/office/vba/api/excel.shape.oleformat)
- [OLEFormat.Object property (Excel)](https://learn.microsoft.com/office/vba/api/excel.oleformat.object)
  - `Shape` が OLE object でない場合は `OLEFormat` が失敗し、`OLEFormat.Object` の返り値も `Object` である。

## 観察結果

### 1. `OLEObjects` は control 専用 collection で、object page が揃っている

- `Worksheet.OLEObjects` / `Chart.OLEObjects` は Office VBA 側に個別 page があり、selector の意味も `name or number` と明示されている。
- `OLEObjects.Item` も同じく単一 object を返す page があり、`Worksheet.OLEObjects(1)` と `Worksheet.OLEObjects.Item(1)` を同系統で正規化しやすい。
- `OLEObject` owner には user-facing member page が揃っているため、最小実装でも completion / hover に使える surface を作りやすい。

### 2. `Sheet1.CommandButton1` は自然だが、静的 inventory の前提が不足している

- Office VBA の概念記事では control code name 経由がもっとも自然な書き方として示されている。
- ただし現行リポジトリの解析入力は `.bas` / `.cls` / `.frm` / `.frx` が主で、worksheet / chart sheet 上の ActiveX control 一覧や control type を静的に列挙する仕組みはまだ持っていない。
- event procedure 名から code name の断片は見えても、module 外から使う一般的な inventory としては不足するため、最初の実装入口に据えるには前提が足りない。

### 3. `Shapes` は広すぎて、control 専用の user-facing root には向かない

- `Shapes` は AutoShape、picture、embedded OLE object など drawing layer 全体を含む。
- `Shapes("CheckBox1")` のような書き方自体は可能だが、戻るのは `Shape` であり、control 専用 member へ進むには `OLEFormat` と `Object` の追加判定が必要になる。
- Office VBA でも `Shape.OLEFormat` は OLE object でない場合に失敗するとされており、control 専用の安全な root としては `OLEObjects` より一段不利である。

### 4. `.Object` の先は control 型とも埋め込みアプリとも限らない

- `OLEObject.Object` と `OLEFormat.Object` はどちらも `Object` を返す。
- 公式例には embedded Word document もあり、常に `Caption` / `Value` / `Select` のような control member へ落とせるわけではない。
- `progID` や workbook 内 metadata から型を絞る戦略を持たない段階で `.Object` を既知 control 型へ進めると、誤補完のリスクが高い。

## 推奨方針

### フェーズ 1: `OLEObjects` root の最小プロトタイプ

- `Worksheet.OLEObjects` / `Chart.OLEObjects` は index 省略時に `OLEObjects` collection として扱う。
- `Worksheet.OLEObjects(Index)` / `Chart.OLEObjects(Index)` と `OLEObjects.Item(Index)` は、name / number selector の単一 object として `OLEObject` owner へ正規化する。
- 初回は `OLEObject` page にある member だけを user-facing に出し、`.Object` の先には進めない。

### フェーズ 2: `.Object` の型付け条件整理

- `OLEObject.progID`、control type metadata、または workbook 由来の design-time 情報から、`.Object` の具体型をどこまで絞れるかを別途整理する。
- control 型が絞れない場合は、`.Object` は未解決のまま維持する。

### フェーズ 3: control code name / `Shapes` の導線整理

- `Sheet1.CommandButton1` を扱うには、code name inventory と owner/type 判定の正本をどこから得るかを決める。
- `Shapes` は control 以外も含むため、もし user-facing root に含めるなら `msoOLEControlObject` 相当の判定か、`OLEFormat` 成功条件を組み合わせる必要がある。

## 次段の完了条件

- `Worksheet.OLEObjects` / `Chart.OLEObjects` / `OLEObjects.Item` の owner 正規化ルールを文書化する。
- `OLEObject` と `OLEObjects` のどの member surface を初回公開するかを fixed list で決める。
- `.Object` を未解決のまま残す理由と、将来型付けへ進むための必要条件を明記する。
- `Sheet1.CommandButton1` と `Shapes` を初回対象から外す理由を、入力ソースと誤補完リスクの両面から固定する。
