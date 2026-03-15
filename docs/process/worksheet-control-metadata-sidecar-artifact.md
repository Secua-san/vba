# Worksheet Control Metadata Sidecar Artifact

## 結論

- workbook package probe の次段では、workspace で常用する static input を `loose files + sidecar` に寄せる。
- sidecar artifact の正本パスは `<bundle-root>/.vba/worksheet-control-metadata.json` に固定する。
- sidecar schema は `worksheet` と将来の `chartsheet` を同じ `owners[]` 配列で表し、未対応 owner も `status` と `reason` を持つ record として残す。
- `OLEObject.Object` 後段型付けと `Sheet1.CommandButton1` 支援に最低限必要なのは、`sheetCodeName`、`shapeName`、`codeName`、control type 判定子である。
- chart sheet は現時点では sidecar に `status: "unsupported"` として記録し、consumer は補完や型解決へ使わない。

## 目的

- workbook package probe の出力を、`.bas` / `.cls` / `.frm` / `.frx` と一緒に扱える runtime input へ落とす。
- `OLEObjects("ShapeName").Object` と `Sheet1.ControlCodeName` の両方で再利用できる共通 inventory format を先に固定する。
- chart sheet 未対応や source 不足を schema 上で表現し、未対応 owner を暗黙の欠落にしない。

## 非目標

- workbook package を extension / server が都度直接読む運用は、この文書では決めない。
- chart sheet control inventory の取得経路そのものは、この文書では新規に仮定しない。
- `Shapes` root や embedded document の `.Object` 型付けまでを今回の sidecar schema に含めない。

## 配置ルール

### 正本パス

- sidecar のファイル名は `worksheet-control-metadata.json` に固定する。
- sidecar は hidden metadata directory である `.vba/` の配下へ置く。
- したがって保存パスは常に `<bundle-root>/.vba/worksheet-control-metadata.json` とする。

### `bundle-root` の定義

- `bundle-root` は、同じ workbook export bundle に属する loose files 群の共通親 directory とする。
- 例:
  - workbook export が workspace root 直下なら `.vba/worksheet-control-metadata.json`
  - workbook export が `samples/book1/` 配下なら `samples/book1/.vba/worksheet-control-metadata.json`

### この形にする理由

- sidecar を module file と同じ階層に散らさず、追加 metadata を `.vba/` に閉じ込められる。
- 将来 `workspace-symbol-index.json` や `control-manifest.json` のような補助 artifact を増やす余地を残せる。
- workbook 名ベースの lookup よりも、現在開いている file から bundle を引く方が誤結合しにくい。

## Lookup ルール

### 探索順

1. 現在解析中の `.bas` / `.cls` / `.frm` / `.frx` file がある directory から探索を始める。
2. 親 directory を順にたどり、最初に見つかった `.vba/worksheet-control-metadata.json` を採用する。
3. workspace folder root を越えて上位へは進まない。
4. multi-root workspace では、現在 file が属する workspace folder ごとに独立して探索する。
5. workspace root を確定できない single-file context では lookup 自体を行わない。

### 採用規則

- 同じ ancestor chain に複数の sidecar がある場合は「最も近いもの」を採用する。
- 複数 sidecar を merge しない。
- sidecar が見つからない file は、従来どおり sidecar 無しで保守動作に戻す。

### 失敗時の扱い

- top-level schema 不正、version 不一致、JSON parse failure の場合は sidecar 全体を無視する。
- owner 単位や control 単位の record 不正は、その record だけを無視し、他は読む。
- 無視した理由は log に残すが、diagnostic や user-facing error は初回では出さない。

## Schema v1

```json
{
  "version": 1,
  "artifact": "worksheet-control-metadata-sidecar",
  "workbook": {
    "name": "Book1.xlsm",
    "sourceKind": "openxml-package"
  },
  "owners": [
    {
      "ownerKind": "worksheet",
      "sheetName": "Sheet1",
      "sheetCodeName": "Sheet1",
      "status": "supported",
      "controls": [
        {
          "shapeName": "CheckBox1",
          "codeName": "chkFinished",
          "shapeId": 3,
          "controlType": "CheckBox",
          "progId": "Forms.CheckBox.1",
          "classId": "{8BD21D40-EC42-11CE-9E0D-00AA006002F3}"
        }
      ]
    },
    {
      "ownerKind": "chartsheet",
      "sheetName": "Chart1",
      "sheetCodeName": "Chart1",
      "status": "unsupported",
      "reason": "chart-sheet-metadata-unproven"
    }
  ]
}
```

## Field 定義

### top-level

| field | 必須 | 意味 |
| --- | --- | --- |
| `version` | 必須 | schema version。初版は `1`。 |
| `artifact` | 必須 | 固定値 `worksheet-control-metadata-sidecar`。別 artifact との取り違え防止用。 |
| `workbook.name` | 必須 | 元 workbook のファイル名。lookup の key ではなく説明用。 |
| `workbook.sourceKind` | 必須 | 初版は `openxml-package` 固定。 |
| `owners` | 必須 | worksheet / chartsheet をまとめた owner 配列。 |

### owner

| field | 必須 | 意味 |
| --- | --- | --- |
| `ownerKind` | 必須 | `worksheet` / `chartsheet` / 将来の `unknown`。 |
| `sheetName` | 必須 | Excel UI 上の sheet 名。 |
| `sheetCodeName` | 必須 | VBA から参照する sheet module 名。 |
| `status` | 必須 | `supported` / `unsupported`。 |
| `reason` | `status=unsupported` で必須 | `unsupported` 理由の固定文字列。 |
| `controls` | `status=supported` で必須 | control inventory。 |

### control

| field | 必須 | 意味 |
| --- | --- | --- |
| `shapeName` | 必須 | `OLEObjects("...")` や `Shapes("...")` で使う name。 |
| `codeName` | 必須 | `Sheet1.CommandButton1` のような direct access で使う name。 |
| `shapeId` | 必須 | workbook package 内の突合と sidecar 再生成検証用。 |
| `controlType` | 必須 | product 側 owner 名。初版候補は `CheckBox` / `CommandButton` / `OptionButton` など。 |
| `progId` | 任意だが推奨 | raw source として保持。 |
| `classId` | 任意だが推奨 | `progId` 欠落時の fallback 識別子。 |

## 最小 field 要件

### `OLEObject.Object` 後段型付けに必要な field

- `sheetCodeName`
- `shapeName`
- `controlType`
- `progId` または `classId`

理由:
- `Worksheet.OLEObjects("CheckBox1")` の selector は shape name で引くため、`shapeName -> controlType` の対応が必要。
- `.Object` の先を `CheckBox` や `CommandButton` へ落とすには、raw source の `progId` / `classId` も残しておく方が再生成や検証に使いやすい。

### `Sheet1.CommandButton1` 支援に必要な field

- `sheetCodeName`
- `codeName`
- `controlType`

理由:
- direct access は workbook 名や sheet display name ではなく、`sheetCodeName + codeName` の組み合わせで解決するのが最短。

### 両方を 1 本で満たす最小集合

- `sheetCodeName`
- `shapeName`
- `codeName`
- `controlType`
- `progId` または `classId`

## 現行 probe 出力との差分

### すでに probe が持っている field

- `workbook`
- `worksheets[].sheetName`
- `worksheets[].sheetCodeName`
- `worksheets[].controls[].shapeName`
- `worksheets[].controls[].codeName`
- `worksheets[].controls[].shapeId`
- `worksheets[].controls[].progId`
- `worksheets[].controls[].classId`

### sidecar v1 で追加するもの

- top-level `artifact`
- `workbook` の object 化
- `worksheets[]` ではなく `owners[]`
- owner ごとの `ownerKind`
- owner ごとの `status`
- unsupported owner の `reason`
- consumer が直接使える `controlType`

### 変換ルール

- probe の `workbook: "Book1.xlsm"` は sidecar では `workbook.name` へ移す。
- probe の `worksheets[]` は sidecar では `ownerKind: "worksheet"` を持つ `owners[]` へ変換する。
- `controlType` は `progId` / `classId` から生成器側で正規化して書く。
- chart sheet は probe が取れない間、sidecar 生成時に空配列へ黙殺せず、`status: "unsupported"` entry を出せる設計にする。

## `controlType` の扱い

- sidecar は raw source だけでなく、product 側で直接使う `controlType` を保持する。
- 初版では `controlType` は current built-in owner 名と一致させる。
- `Worksheet` / `Chart` 上の ActiveX control では `CommandButton` / `CheckBox` / `OptionButton` を使い、`DialogSheet` 側の `Button` / `CheckBox` / `OptionButton` とは別 owner として扱う。
- 例:
  - `Forms.CheckBox.1` -> `CheckBox`
  - `Forms.CommandButton.1` -> `CommandButton`
  - `Forms.OptionButton.1` -> `OptionButton`
- raw の `progId` / `classId` は再生成・監査・将来の正規化更新のために残す。

## 未対応 owner の表し方

### chart sheet

- chart sheet は初版で `ownerKind: "chartsheet"` を予約し、未対応の間は次のように表す。

```json
{
  "ownerKind": "chartsheet",
  "sheetName": "Chart1",
  "sheetCodeName": "Chart1",
  "status": "unsupported",
  "reason": "chart-sheet-metadata-unproven"
}
```

### consumer 規則

- `status: "unsupported"` の owner は inventory source として使わない。
- `controls` が空だから supported なのか、source 未証明なのかを区別するため、unsupported は明示 record にする。
- 将来 chart sheet support を開く場合も、同じ `ownerKind` を `supported` へ更新すれば schema を壊さず拡張できる。

## 保存時の追加ルール

- `generatedAt` のような diff ノイズになりやすい volatile field は sidecar v1 には入れない。
- key 順は固定し、再生成で不要差分を増やさない。
- 1 bundle に 1 sidecar を基本とし、sheet ごとの分割 sidecar は初版では採用しない。

## 2026-03-14 時点の実装状況

- 生成経路:
  - `npm run generate:worksheet-control-sidecar -- <workbook-path> --bundle-root <bundle-root>`
  - または `npm run probe:worksheet-control-metadata -- <workbook-path> --format sidecar --bundle-root <bundle-root>`
- generator は probe 出力を schema v1 へ変換し、`<bundle-root>/.vba/worksheet-control-metadata.json` を書き出せる。
- `packages/core` には nearest ancestor lookup、workspace root での打ち切り、schema v1 validation、`status: "unsupported"` owner の切り分け helper を実装済み。
- workspace root が確定していない single-file context では sidecar lookup 自体を行わず、誤結合を避ける。
- `packages/server` には read-only cache と log を実装済みで、sidecar は `DocumentState` へ保持される。
- sidecar generator は未知 `controlType` を黙って落とさず fail-fast とし、対応していない workbook 方言差は生成時点で止める。
- `Sheet1.OLEObjects("ShapeName").Object` と `Sheet1.OLEObjects.Item("ShapeName").Object` の string literal selector にだけ接続済みで、`shapeName -> controlType` を使って `CheckBox` などの control owner へ進める。
- `Sheet1.chkFinished.Value` のような worksheet document module root の direct access も接続済みで、`sheetCodeName + codeName -> controlType` を使って `CheckBox` などの control owner へ進める。
- 数値 selector、dynamic selector、`ActiveSheet` root、supported/unsupported を問わない chartsheet owner は従来どおり未解決のまま維持する。

## 次段の実装候補

1. `Worksheet.Shapes("ShapeName")` / `Chart.Shapes("ShapeName")` と `Shape.OLEFormat.Object` のどこまでを sidecar と組み合わせて公開するかを整理する。
2. drawing object 全体を含む `Shapes` root を control 専用 owner へ誤昇格させない境界条件を、`msoOLEControlObject` 判定や `OLEFormat` 成功条件と合わせて固定する。
3. workbook / standard module からの非 document-module access を広げる必要がある場合は、`sheetCodeName` 以外の root identity をどこまで許可するかを別タスクとして切り出す。
