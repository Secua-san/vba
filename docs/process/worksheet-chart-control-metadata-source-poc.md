# Worksheet / Chart Control Metadata Source PoC

## 結論

- `Worksheet` / `Chart` 上の ActiveX control metadata source の第 1 PoC は、workbook package を正本にして行う。
- ただし、現時点で workbook package から経路が十分に確認できているのは worksheet 側であり、chart sheet 側は `codeName` と drawing part への到達までは確認できるが、control inventory の復元経路は未証明である。
- そのため次段の最小 PoC は「worksheet 限定の workbook package probe」とし、出力は loose files 側で再利用できる sidecar JSON を想定する。
- `OLEObject.Object` の後段型付けと `Sheet1.CommandButton1` のような control code name 導線は、この sidecar 形式で `sheet module`、`shape name`、`code name`、`ProgID` または同等の型識別子を安定して復元できるまで未解決のまま維持する。

## 確認した公式ソース

### Office VBA

- [Using ActiveX Controls on Sheets](https://learn.microsoft.com/office/vba/excel/concepts/controls-dialogboxes-forms/using-activex-controls-on-sheets)
  - `Sheet1.CommandButton1.Caption` のような direct access は control code name を使う。
  - `ActiveSheet.OLEObjects("CheckBox1").Object.Value` のような collection access は shape name を使う。
- [OLEObjects object (Excel)](https://learn.microsoft.com/office/vba/api/excel.oleobjects)
  - ActiveX control には shape name と code name の 2 つの名前がある。
- [OLEObject.progID property (Excel)](https://learn.microsoft.com/office/vba/api/excel.oleobject.progid)
  - `OLEObject` から `ProgID` を取得できる。
- [Chart.OLEObjects method (Excel)](https://learn.microsoft.com/office/vba/api/excel.chart.oleobjects)
  - chart sheet でも `OLEObjects` 導線自体は object model 上に存在する。
- [OLE programmatic identifiers (Office)](https://learn.microsoft.com/office/vba/library-reference/concepts/ole-programmatic-identifiers-office)
  - `Forms.CommandButton.1`、`Forms.CheckBox.1` など control type と `ProgID` の対応がある。

### Open XML / SpreadsheetML

- [SheetProperties](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.sheetproperties?view=openxml-3.0.1)
  - `sheetPr@codeName` は「user input では変わらない stable name」で、code から sheet を参照する名前として使うと明記されている。
- [ChartSheetProperties](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.chartsheetproperties?view=openxml-3.0.1)
  - chart sheet でも `sheetPr@codeName` は保持される。
- [Controls](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.controls?view=openxml-3.0.1)
  - worksheet 上の `controls` collection は control code name の列挙と drawing 情報の参照に使うと明記されている。
- [Control](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.control?view=openxml-3.0.1)
  - `control@name` は control の code name、`control@shapeId` は drawing part の shape id、`r:id` は Embedded Control Data part への relationship である。
- [Worksheet](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.worksheet?view=openxml-3.0.1)
  - worksheet は `controls` と `drawing` を child element として持つ。
- [Chartsheet](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.chartsheet?view=openxml-3.0.1)
  - chart sheet は `drawing` と `sheetPr` は持つが、`controls` / `oleObjects` は child element 一覧に現れない。
- [Drawing](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.drawing?view=openxml-3.0.1)
  - worksheet / chart sheet の `drawing` は drawingML part への relationship を持つ。
- [WorksheetDrawing](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.spreadsheet.worksheetdrawing?view=openxml-3.0.1)
  - drawing part には shape 群が入る。
- [NonVisualDrawingProperties.Id](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.spreadsheet.nonvisualdrawingproperties.id?view=openxml-3.0.1)
  - `xdr:cNvPr@id` は document 内で一意な drawing object id。
- [NonVisualDrawingProperties.Name](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.spreadsheet.nonvisualdrawingproperties.name?view=openxml-3.0.1)
  - `xdr:cNvPr@name` は drawing object の name。
- [OleObject](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.oleobject?view=openxml-3.0.1)
  - worksheet 側の `oleObject@progId` と `oleObject@shapeId` が定義されている。
- [ActiveXControlData.ActiveXControlClassId](https://learn.microsoft.com/dotnet/api/documentformat.openxml.office.activex.activexcontroldata.activexcontrolclassid?view=openxml-3.0.1)
  - Embedded Control Data part には `classid` がある。

## 観察結果

### 1. exported modules だけでは control inventory を復元できない

- `.cls` から `Sheet1` や `Chart1` の sheet module 名は読める。
- しかし exported `Sheet1.cls` / `Chart1.cls` 自体には、control 一覧、shape name、code name、`ProgID` の対応は含まれない。
- event procedure 名が存在しても、その sheet に載っている全 control inventory にはならない。
- そのため `.bas` / `.cls` / `.frm` / `.frx` だけで `OLEObject.Object` や `Sheet1.CommandButton1` を型付けするのは無理がある。

### 2. workbook package は worksheet 側の primary source として十分に筋が良い

- `sheetPr@codeName` から sheet module 名を取れる。
- `controls/control@name` から control code name を取れる。
- `controls/control@shapeId` から drawing shape への接続点を取れる。
- drawing part の `xdr:cNvPr@id` / `@name` から shape id と shape name を取れる。
- `oleObject@progId`、または Embedded Control Data part の `classid` から control type 判定に使える識別子を取れる可能性が高い。

### 3. chart sheet では workbook package の経路がまだ対称でない

- chart sheet でも `sheetPr@codeName` と `drawing` part までは確認できる。
- 一方、Open XML の `Chartsheet` page には worksheet と違って `controls` / `oleObjects` が出てこない。
- `Chart.OLEObjects` は object model 上は存在するため、package 内の別経路があるか、実ファイルでしか確認できない差分がある可能性は残る。
- ただし、現時点の公式 docs だけでは `chart sheet -> control code name / ProgID` の復元経路を断定できない。

### 4. extraction artifact と manifest は delivery format としては有望だが、source of truth には向かない

- extract 時の sidecar artifact なら、workbook package から拾った metadata を loose files と一緒に配布できる。
- 将来の manifest も同じ情報を持てるが、schema 設計、更新タイミング、整合監査を新たに抱える。
- どちらも consumer format としては有効だが、元データをどこから取るかという PoC では workbook package の方が根拠が強い。

## 入力源比較

| 入力源 | sheet module | shape name | code name | `ProgID` / 型識別 | 評価 |
| --- | --- | --- | --- | --- | --- |
| exported `.bas` / `.cls` / `.frm` / `.frx` | 一部可 | 不可 | inventory 不可 | 不可 | 不十分 |
| workbook package（worksheet） | 可 | 可 | 可 | 可または `classid` まで可 | 第 1 PoC 候補 |
| workbook package（chart sheet） | 可 | drawing までは可 | 未証明 | 未証明 | 要実ファイル確認 |
| extract 時 sidecar artifact | 設計次第で可 | 設計次第で可 | 設計次第で可 | 設計次第で可 | delivery format 候補 |
| 将来 manifest | 設計次第で可 | 設計次第で可 | 設計次第で可 | 設計次第で可 | 後続候補 |

## 推奨方針

### フェーズ 1: workbook package probe を worksheet 限定で作る

- 最初の probe は `.xlsm` / `.xlam` のような Open XML package を対象にする。
- `.xlsb` を扱うかどうかは、この PoC とは別に入力形式の追加判断として切り分ける。
- 少なくとも以下を抽出できるかを確認する。
  - workbook 内 sheet 名
  - sheet module code name
  - control code name
  - shape id
  - shape name
  - `ProgID` または `classid`
- 出力は loose files から再利用しやすい JSON にする。
- PoC では `controls/control@shapeId`、`oleObject@shapeId`、drawing の `xdr:cNvPr@id` が同じ control を指すことを必須検証に含める。

### フェーズ 2: sidecar artifact を static input に寄せる

- probe が成立したら、extract 時に sidecar artifact を吐くか、workbook package を直接読む補助コマンドを作るかを決める。
- 既存の `.bas` / `.cls` / `.frm` / `.frx` 主軸を崩さないため、extension / server の日常入力は loose files + sidecar を優先候補とする。

### フェーズ 3: chart sheet support を別条件で開く

- chart sheet は object model と package docs の非対称が残るため、worksheet と同じ前提で一緒に実装しない。
- 実 workbook package の inspection で control inventory を確認できた場合だけ、worksheet と同じ sidecar schema に統合する。
- 確認できない間は `Chart1.<control code name>` と `Chart.OLEObjects(...).Object` の後段型付けは保守動作のまま維持する。

## 最小 PoC の出力案

```json
{
  "version": 1,
  "workbook": "sample.xlsm",
  "worksheets": [
    {
      "sheetName": "Sheet1",
      "sheetCodeName": "Sheet1",
      "controls": [
        {
          "shapeName": "CheckBox1",
          "codeName": "chkFinished",
          "shapeId": 3,
          "progId": "Forms.CheckBox.1",
          "classId": "{...}"
        }
      ]
    }
  ]
}
```

- `progId` と `classId` の両方が取れるなら両方残す。
- user-facing 解決で必須なのは `sheetCodeName`、`shapeName`、`codeName`、control type 判定子である。
- chart sheet が未対応の間は、`chartsheets` は空か省略でよい。

## 次段の完了条件

- workbook package から worksheet control inventory を抽出する最小 probe を追加する。
- 初回対象の入力形式を `.xlsm` / `.xlam` のどちらか、または両方に固定する。
- `shape name != code name` のケースを 1 つ以上含む fixture で、両方が別値として取れることを確認する。
- `controls`、`oleObjects`、drawing の `shapeId` が同じ control へ収束することを fixture で確認する。
- `ProgID` と `classid` のどちらを type 判定の正本にするかを決める。
- chart sheet については「同じ経路で取れる」か「別 artifact が必要」か「当面保留」かのいずれかを明文化し、保留継続なら解除条件も残す。
