# DialogSheet Interop Source Feasibility

## 結論

- `DialogSheet` interop page を、そのまま VBA 補完用の owner データへ自動取り込みするのは現時点では不採用とする。
- ただし、補助ソースとして限定利用すること自体は可能であり、将来対応するなら「明示 allow list + skip rule + 監査テスト」の形が現実的である。
- 2026-03-13 時点では、この限定利用を `Application.DialogSheets` / `Workbook.DialogSheets` の補助 root まで広げることを許容する。ただし `DialogSheet1.` document module root は引き続き未公開のまま維持する。

## 確認した公式ソース

- Office VBA 概念記事 `Refer to Sheets by Name`
  - `DialogSheets("Dialog1").Activate` が示され、dialog sheet 自体は VBA から参照可能
  - `Sheets` collection が worksheet / chart / module / dialog sheet を含むことを明記
- Office VBA API
  - `DialogSheetView object` は存在する
  - ただし `DialogSheet object` は Office VBA object page として見つからず、現行 `mslearn-vba-reference.json` にも owner が存在しない
- Microsoft Learn .NET interop
  - `DialogSheet Interface` は存在し、property / method 一覧を持つ
  - ただしページ先頭に `Reserved for internal use.` があり、`_Dummy*` や `_SaveAs` のような内部向け・旧表記の member を含む

## 現行生成器との相性

- 現行の `scripts/generate-mslearn-vba-reference.mjs` は Office VBA API TOC JSON を起点に owner を列挙している
- interop page はこの TOC 経路に乗っていないため、導入するなら個別 URL manifest か別の discovery 経路が必要
- 既存の抽出ロジックは `Syntax` / `Parameters` / `Return value` を前提にしているが、interop page の member 一覧は owner page 上ではリンク表中心で、owner 単位の一括抽出には別処理が必要

## member 品質の観点

### そのまま導入しにくい理由

- owner page 自体が `Reserved for internal use.` で、VBA 向け正本としての優先度が低い
- `_Dummy*` が大量に含まれ、機械取り込みするとノイズが高い
- `_SaveAs` や `_Evaluate` のような旧表記・内部表記と公開 member が同居している
- control collection 系 (`Buttons`, `CheckBoxes`, `OptionButtons`, `DialogFrame` など) は `DialogSheet` 固有だが、戻り先 owner や補完の深掘り先を別途用意する必要がある

### 補助ソースとしては使える理由

- `Activate`, `Evaluate`, `ExportAsFixedFormat`, `Move`, `PrintOut`, `SaveAs`, `Select`, `Unprotect` など、既存 `Worksheet` / `Chart` と重なる common callable を確認できる
- `DialogSheet.CodeName` など document module root の文脈で最低限欲しい member の存在確認には使える
- individual member page を辿れば syntax 情報があるため、限定メンバーなら signature 抽出候補にできる

## 導入するなら必要な制約

### 1. source manifest を分離する

- Office VBA TOC とは別に、interop 補助ソース専用の owner / member manifest を持つ
- 初回対象は `DialogSheet` に限定する

### 2. skip rule を先に固定する

- `_Dummy*` は全除外
- `_SaveAs` / `_Evaluate` のような `_` 接頭辞 legacy member は全除外
- owner / member 説明に `Reserved for internal use.` が含まれるものは既定で除外

### 3. allow list ベースで始める

- common callable だけを明示 allow list に載せる
- 初回候補:
  - `Activate`
  - `Evaluate`
  - `ExportAsFixedFormat`
  - `Move`
  - `PrintOut`
  - `SaveAs`
  - `Select`
  - `Unprotect`
- control collection や `DialogFrame` は別 owner 設計が必要なので後回しにする

### 4. 監査テストを追加する

- 生成 JSON に `_Dummy` が混入しない
- `DialogSheet` owner の member 名に `_` 接頭辞の legacy member が残らない
- allow list 以外の interop member を owner page から機械追加しない

## 推奨方針

- 現段階では `DialogSheet` 全面導入は行わず、`docs/adr/0004-dialogsheet-document-module-policy.md` の保守方針を維持する
- 次に進めるなら、interoperability 補助ソースの最小プロトタイプとして common callable だけを抽出対象にした小さな実験 PR に分ける

## 2026-03-13 の Workbook / Application root 展開判断

- Office VBA には `Application.Sheets` / `Workbook.Sheets` があり、workbook root から sheet collection を辿る基本パターン自体は正本で確認できる
- 一方、`DialogSheets` 専用 property page は Office VBA 側に無く、Microsoft Learn では interop の `ApplicationClass.DialogSheets` / `WorkbookClass.DialogSheets` が `ReadOnly Property DialogSheets As Sheets` を示す
- 実装ではこの property value をそのまま `Sheets` にすると `DialogSheet` item owner へ落とせないため、user-facing root としては `DialogSheets` collection owner へ正規化する
- この正規化は `Application.DialogSheets(1)` / `ActiveWorkbook.DialogSheets(1)` の単一 selector にだけ効かせ、`Application.DialogSheets(Array(...))` のような grouped selector は collection のまま維持する
- `DialogSheet1.` の document module root 昇格は引き続き別論点とし、この PR では扱わない
