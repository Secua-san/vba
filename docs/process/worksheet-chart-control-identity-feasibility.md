# Worksheet / Chart Control Identity Feasibility

## 結論

- `Worksheet` / `Chart` 上で `OLEObject.Object` の先を型付きにすることと、`Sheet1.CommandButton1` のような control code name 導線を user-facing に出すことは、どちらも同じ前提に依存する。
- その前提は「sheet module」「shape name」「code name」「control type / ProgID」を結び付けた control inventory である。
- Microsoft Learn は、collection access では shape name を使い、event procedure では code name を使うこと、さらに `OLEObject.progID` で control type を取得できることを示している。
- 一方で、現行リポジトリの静的入力である `.bas` / `.cls` / `.frm` / `.frx` だけでは、worksheet / chart sheet 上の ActiveX control inventory を安定して復元できない。
- そのため現段階でも、静的入力だけで `OLEObject.Object` や `Sheet1.CommandButton1` を解決することはできないが、sidecar v1 がある worksheet document module root に限っては `shapeName -> controlType` と `codeName -> controlType` の例外接続が成立する。

## 実装更新

- 2026-03-14 時点で、sidecar v1 を使った `Sheet1.OLEObjects("ShapeName").Object` と `.Item("ShapeName").Object` の string literal selector だけは user-facing に接続済みである。
- 2026-03-15 時点で、same sidecar v1 を使った `Sheet1.chkFinished.Value` のような worksheet document module root の direct access も user-facing に接続済みである。
- ただしこれらはいずれも worksheet document module root に限った例外であり、`Chart1`、`ActiveSheet`、dynamic selector、`Worksheets(1).chkFinished`、sidecar 未検出時の direct access は引き続き未解決のまま維持する。

## 確認した公式ソース

### Office VBA

- [Using ActiveX Controls on Sheets](https://learn.microsoft.com/office/vba/excel/concepts/controls-dialogboxes-forms/using-activex-controls-on-sheets)
  - `Sheet1.CommandButton1.Caption` のように、sheet class module 内では control code name を使う。
  - `Worksheets(1).OLEObjects("CommandButton1").Object.Caption` のように、collection access では `Object` property を経由して control 固有 property へ進む。
  - `Shapes` / `OLEObjects` から control を名前で返すときは、code name ではなく shape name を使う。
- [OLEObjects object (Excel)](https://learn.microsoft.com/office/vba/api/excel.oleobjects)
  - ActiveX control には shape name と code name の 2 つの名前がある。
  - `OLEObjects("CheckBox1")` のような collection access では shape name を使う。
- [OLEObject.progID property (Excel)](https://learn.microsoft.com/office/vba/api/excel.oleobject.progid)
  - `OLEObject` から programmatic identifier を取得できる。
  - `progID` は control type 判定の候補になる。
- [OLE programmatic identifiers (Office)](https://learn.microsoft.com/office/vba/library-reference/concepts/ole-programmatic-identifiers-office)
  - `Forms.CommandButton.1`、`Forms.CheckBox.1`、`Forms.OptionButton.1` など、ActiveX control と ProgID の対応表がある。

## 観察結果

### 1. `.Object` の型付けと code name 導線は、必要な metadata が同じ

- `Worksheets(1).OLEObjects("CheckBox1").Object.Value` を型付きにするには、`"CheckBox1"` がどの control type を指すかを知る必要がある。
- `Sheet1.chkFinished.Value` を built-in member として解決するにも、`chkFinished` がどの control type を指すかを知る必要がある。
- つまり `.Object` 後段型付けと `Sheet1.CommandButton1` 支援は別機能に見えて、実際には同一の control identity source を必要とする。

### 2. Microsoft Learn は shape name / code name / ProgID の役割分担を示している

- sheet class module の event procedure や direct member access は code name を使う。
- `OLEObjects(...)` / `Shapes(...)` は shape name を使う。
- `OLEObject.progID` は control type 判定に使える。
- この 3 つが揃えば、`shapeName -> codeName -> controlType` の対応表を作れる。

### 3. 現行リポジトリの静的入力には、その対応表を作る材料がない

- たとえば [Sheet1.cls](/C:/Users/tagi0/Documents/dev/vba/packages/extension/test/fixtures/Sheet1.cls) は `VB_Name` / `VB_Base` / `VB_PredeclaredId` と `Option Explicit` だけで、control 一覧や control type を持たない。
- [Chart1.cls](/C:/Users/tagi0/Documents/dev/vba/packages/extension/test/fixtures/Chart1.cls) も同様で、chart sheet 上の ActiveX control inventory は含まれない。
- `.frm` / `.frx` は userform には有効だが、worksheet / chart sheet 上の ActiveX control metadata source とは別である。
- 現状の parser / symbol pipeline に新しい型推論規則を足しても、元データが無い以上 `Sheet1.CommandButton1` や `.Object.Caption` を安全に既知型へ落とせない。

### 4. 名前からの推測で型を決めるのは危険

- `CheckBox1` や `CommandButton1` のような名前は既定命名ではあるが、ユーザーが自由に変更できる。
- Office VBA の正本でも、shape name と code name は変更によりずれ得るとされている。
- そのため、文字列 selector や identifier の見た目だけで `CheckBox` / `CommandButton` を推測すると誤補完になりやすい。

## 推奨方針

### フェーズ 1: sidecar あり worksheet root の限定公開

- `OLEObjects(Index)` / `OLEObjects.Item(Index)` は、既定では `OLEObject` までで止める。
- `Sheet1.OLEObjects("ShapeName").Object` / `.Item("ShapeName").Object` の string literal selector と、`Sheet1.ControlCodeName` の direct access は、sidecar v1 がある worksheet document module root に限って公開する。
- 数値 selector、dynamic selector、`ActiveSheet`、chart sheet root、sidecar 未検出の direct access は未解決のまま維持する。

### フェーズ 2: metadata source の PoC

- worksheet / chart sheet 上の ActiveX control について、少なくとも以下を取れる source を比較する。
  - sheet module 名
  - shape name
  - code name
  - ProgID または control type
- 候補は workbook package、抽出ツールの補助 artifact、将来の manifest 生成などとする。
- この PoC が成立しない限り、worksheet document module root 以外の `.Object` と control code name は未解決のまま据え置く。
- PoC が成立した場合は、このメモの「静的入力だけでは復元不可」という前提を、`sidecar を含む静的入力ならどこまで公開できるか` へ更新する。

### フェーズ 3: user-facing 解決の追加

- metadata source が固まったら、まず `.Object` 後段型付けと `Sheet1.CommandButton1` のどちらを先に出すかを決める。
- どちらを先に出す場合でも、shape name と code name がずれるケースを回帰テストで固定する。

## 次段の完了条件

- worksheet / chart sheet 上の ActiveX control inventory を取る候補 source を 1 つ以上絞る。
- その source から `shape name`、`code name`、`ProgID` のどこまで取れるかを表で整理する。
- 情報が足りない場合は、`.Object` と control code name を未解決のまま維持する理由を更新する。
