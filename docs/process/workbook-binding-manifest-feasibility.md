# Workbook Binding Manifest Feasibility

## 結論

- broad root 再評価の前提となる workbook identity は、専用 artifact `<bundle-root>/.vba/workbook-binding.json` へ切り出す。
- binding の primary key は `Workbook.FullName` の正規化値とし、`Workbook.Name` / `Workbook.Path` は診断用の補助 field とする。
- `Workbook.Path` が空の workbook と `Workbook.IsAddin = True` の workbook は、manifest があっても broad root binding の対象にしない。
- v1 では saved かつ non-addin workbook にだけ manifest を生成し、manifest 不在は binding disabled として扱う。
- workbook package mode は manifest を生成する source of truth としては使えるが、loose file workflow の runtime binding transport には使わない。
- broad root を user-facing に開くのは、manifest に加えて host から active workbook identity を受け取る契約が定義された後とする。
- runtime 側の契約と state schema は、正本 [active-workbook-identity-provider-contract.md](./active-workbook-identity-provider-contract.md) に分離する。

## 確認した公式ソース

### Office VBA

- [Application.ThisWorkbook property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.thisworkbook)
  - `ThisWorkbook` は current macro code が動いている workbook を返す。
  - add-in では `ActiveWorkbook` は add-in workbook ではなく caller workbook を返す。
- [Application.ActiveWorkbook property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.activeworkbook)
  - `ActiveWorkbook` は active window の workbook を返す。
- [Workbook.FullName property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.fullname)
  - workbook の path を含む名前を返す。
- [Workbook.Path property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.path)
  - workbook の完全 path を返す。
- [Workbook.Saved property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.saved)
  - workbook が一度も保存されていない場合、`Path` は空文字列になる。
- [Workbook.IsAddin property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.isaddin)
  - add-in workbook は通常 workbook と異なる実行条件を持つ。

## 現行前提

- broad root は既定では閉じているが、現行実装では `ActiveWorkbook.Worksheets("Sheet1")` / `.Item("Sheet1")` と、unqualified `Worksheets("Sheet1")` / `.Item("Sheet1")`、`Application.Worksheets("Sheet1")` / `.Item("Sheet1")` が `available` snapshot と manifest match のときだけ user-facing に開く。
- sidecar v1 は `shapeName` / `codeName` / `controlType` の inventory を持つが、runtime workbook identity は持たない。
- `loose files + sidecar` が日常入力であり、extension / server が workbook package を都度直接読む運用はまだ採らない。
- `.vba/` 配下には sidecar 以外の補助 artifact を追加できる余地がある。

## 選択肢比較

| 選択肢 | 利点 | 欠点 | 判断 |
| --- | --- | --- | --- |
| `worksheet-control-metadata.json` に binding を混ぜる | file 数が増えない | path / add-in 状態のような volatile 情報が control inventory と混ざる | 不採用 |
| workspace / user config に workbook path を置く | 実装が単純 | multi-bundle workspace で誤結合しやすく、共有 artifact にしにくい | 不採用 |
| workbook package mode を runtime binding に使う | source of truth と 1 本化できる | loose file workflow と相性が悪く、普段の解析入力を変えてしまう | 不採用 |
| bundle-local manifest + host active workbook identity | loose file workflow を維持しつつ binding を明示できる | host 契約が別途必要 | 採用 |

## 観察結果

### 1. primary key は `Workbook.FullName` が最も筋がよい

- `Workbook.Name` だけでは同名 workbook が複数開いている場合に衝突する。
- `Workbook.Path` だけでは file 名が落ちる。
- `Workbook.FullName` は path を含む workbook 名であり、saved workbook なら最も具体的な identity として扱いやすい。

### 1.5 `Workbook.FullName` の v1 正規化は narrow に始める

- compare は Windows 前提で case-insensitive とする。
- path separator は `/` と `\` の揺れだけを吸収し、manifest / host 双方で `\` にそろえる。
- UNC path は `\\server\share\...` の形を保持し、drive letter path と混同しない。
- `FullNameURLEncoded` は v1 では使わず、URL decode も行わない。
- つまり v1 は「Windows file system path としての `FullName` 一致」だけに絞る。

### 2. unsaved workbook は broad root binding の対象にできない

- `Workbook.Saved` の正本では、「一度も保存されていない workbook は `Path` が空文字列」と明記されている。
- `Path` が空だと `FullName` の安定性も期待できないため、manifest と runtime workbook を再現可能に照合できない。
- そのため unsaved workbook は broad root binding 非対応とする方が安全である。

### 3. add-in workbook は broad root の対象外にするべきである

- `ThisWorkbook` と `ActiveWorkbook` は add-in で乖離し得る。
- broad root が指したいのは active workbook だが、bundle が add-in 由来の場合、current bundle 側の manifest と runtime caller workbook は別物になりやすい。
- したがって `isAddIn = true` の bundle は manifest があっても broad root binding を有効にしない。

### 4. binding 情報は sidecar から分離した方が運用しやすい

- sidecar は control inventory の再生成物であり、workbook rename / move のたびに path 情報まで差分へ混ぜるとレビュー効率が落ちる。
- binding manifest を分離すれば、control inventory と runtime identity の更新タイミングを分けられる。
- `.vba/` 配下へ置けば bundle-local artifact という前提も保てる。

### 5. manifest だけでは broad root は開けない

- broad root を有効にするには、current bundle 側の manifest と runtime 側の active workbook identity を照合する必要がある。
- current product にはまだ host から active workbook identity を受け取る経路が無い。
- したがって今回決めるのは manifest policy までであり、feature 開放は別 task とする。

### 6. generator と consumer の責務は v1 で分けておく

- generator は saved かつ non-addin workbook にだけ `workbook-binding.json` を生成する。
- unsaved workbook と add-in workbook では manifest を生成せず、理由は log または docs で説明する。
- consumer は manifest 不在を error にせず、「binding disabled」として broad root を閉じたままにする。
- これにより unsupported workbook state を invalid manifest で表現せずに済む。

## 推奨 manifest v1

```json
{
  "version": 1,
  "artifact": "workbook-binding-manifest",
  "bindingKind": "active-workbook-fullname",
  "workbook": {
    "fullName": "C:\\Work\\Book1.xlsm",
    "name": "Book1.xlsm",
    "path": "C:\\Work",
    "isAddIn": false,
    "sourceKind": "openxml-package"
  }
}
```

## Field 方針

| field | 必須 | 用途 |
| --- | --- | --- |
| `version` | 必須 | schema version |
| `artifact` | 必須 | 固定値 `workbook-binding-manifest` |
| `bindingKind` | 必須 | 初版は `active-workbook-fullname` |
| `workbook.fullName` | 必須 | primary match key |
| `workbook.name` | 必須 | 診断 / log 用 |
| `workbook.path` | 必須 | unsaved 判定と診断用 |
| `workbook.isAddIn` | 必須 | broad root 対象外判定 |
| `workbook.sourceKind` | 必須 | 初版は `openxml-package` |

## Matching rule v1

- `manifest.workbook.fullName` と runtime active workbook `FullName` を Windows path として比較する。
- 比較前に以下だけを行う。
  - case-insensitive 化
  - `/` を `\` へ統一
  - 末尾 separator の余分な揺れを除去
- 以下は v1 では行わない。
  - URL decode
  - short path / long path 展開
  - symlink / junction 解決
  - UNC と drive mapping の同一視

## なぜ sidecar へ混ぜないか

- `worksheet-control-metadata.json` は control inventory の schema として既に安定化を進めている。
- binding は workbook の保存場所や add-in 状態に依存し、control inventory より変動しやすい。
- したがって sidecar に `fullName` を追加するより、別 manifest へ切り出す方が責務分離と diff 管理の両面で有利である。

## workbook package mode との関係

- workbook package probe / sidecar generator は、manifest を生成する source of truth 候補として使える。
- ただし runtime 側で broad root を判定するときに workbook package を毎回直接読む設計は、loose file workflow と相性が悪い。
- そのため runtime consumer は manifest を読み、generator 側が workbook package から manifest を再生成する分担がよい。

## user-facing 境界

### 既に開いているものの代表例

- `ActiveWorkbook.Worksheets("Sheet1").OLEObjects("ShapeName").Object`
- `ActiveWorkbook.Worksheets("Sheet1").Shapes("ShapeName").OLEFormat.Object`
- `ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects("ShapeName").Object`
- `ActiveWorkbook.Worksheets.Item("Sheet1").Shapes("ShapeName").OLEFormat.Object`
- unqualified `Worksheets("Sheet1").OLEObjects("ShapeName").Object`
- unqualified `Worksheets.Item("Sheet1").OLEObjects("ShapeName").Object`
- `Application.Worksheets("Sheet1").Shapes("ShapeName").OLEFormat.Object`
- `Application.Worksheets.Item("Sheet1").Shapes("ShapeName").OLEFormat.Object`

- 上の列挙は broad-root family の代表例であり、実装済みの全パターンを網羅列挙するものではない。

### policy に沿って user-facing 済みのもの

- `Application.ThisWorkbook.Worksheets("Sheet1")...`
  - `ThisWorkbook` direct root と同じ static current-bundle family として `OLEObject.Object` / `Shape.OLEFormat.Object` の sidecar lookup へ接続済み。
- `Application.ActiveWorkbook.Worksheets("Sheet1")...`
  - `ActiveWorkbook` direct root と同じ manifest + snapshot gating family として、match 時だけ `OLEObject.Object` / `Shape.OLEFormat.Object` の sidecar lookup へ接続済み。

### active-workbook family を開く条件

- bundle-local `workbook-binding.json` が存在する
- manifest の `workbook.fullName` が runtime active workbook の `FullName` と一致する
- `workbook.path` が空でない
- `workbook.isAddIn` が `false`
- 同じ PR で `OLEObject.Object` と `Shape.OLEFormat.Object` をそろえて開く

## 今回の完了条件

- manifest / config / workbook package mode のどれを binding transport にするか決める。
- primary key として使う workbook identity を決める。
- unsaved workbook / add-in workbook を broad root 対象にしない理由を残す。
- v1 の `FullName` 正規化ルールと generator / consumer の責務境界を残す。
- broad root feature 開放に別途 host 契約が必要であることを明文化する。

## 関連メモ

- `Application.ThisWorkbook` / `Application.ActiveWorkbook` root の整理結果は [application-workbook-root-feasibility.md](./application-workbook-root-feasibility.md) を参照する。
