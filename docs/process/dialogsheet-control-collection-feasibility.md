# DialogSheet Control Collection Feasibility

## 結論

- `DialogSheet` 専用 control collection は、全面導入ではなく段階導入が妥当である。
- 最初の導入候補は `DialogFrame` に限定する。`DialogSheet.DialogFrame` は `DialogFrame` 型を直接返すため、owner 正規化と chain 解決の複雑さが最も小さい。
- `Buttons` / `CheckBoxes` / `OptionButtons` は将来的な導入候補にはできるが、`Optional Index As Object -> As Object` のため、単一要素アクセスと collection access を分ける補助ルールが先に必要である。
- 2026-03-13 時点では、`DialogSheet` common callable と `Application/Workbook.DialogSheets` root までで止め、control collection は docs 先行で整理したうえで次段タスクへ送る。

## 確認した公式ソース

### Office VBA

- Excel 概念記事 `Using ActiveX Controls on Sheets`
  - worksheet / chart sheet 上の ActiveX control は通常 `OLEObjects` / `Shapes` 経由で扱う導線が正本。
  - `Sheet1.CommandButton1.Caption` や `Worksheets(1).OLEObjects("CommandButton1").Object.Caption` のように、sheet 上 control の一般導線は collection method より control 名参照へ寄っている。
- Excel 概念記事 `Refer to Sheets by Name`
  - `DialogSheets("Dialog1").Activate` が示され、dialog sheet 自体は VBA から参照可能。

### Microsoft Learn .NET interop

- `DialogSheet.Buttons(Object)` / `CheckBoxes(Object)` / `OptionButtons(Object)`
  - いずれも `Public Function ...(Optional Index As Object) As Object`
  - `Index` ありでも無しでも戻り値が `Object` で、owner を機械的に一意決定できない。
- `DialogSheet.DialogFrame`
  - `Public ReadOnly Property DialogFrame As DialogFrame`
  - 直接 `DialogFrame` owner へ落とせる。
- `Buttons` / `CheckBoxes` / `OptionButtons` interface
  - いずれも `Reserved for internal use.`
  - collection 自体の property / method table を持つが、`_Dummy*` を含む。
  - `Add(...)` はそれぞれ `Button` / `CheckBox` / `OptionButton` を返す一方、`Item(Object)` は `Object` を返す。
- `Button` / `CheckBox` / `OptionButton` / `DialogFrame` interface
  - いずれも `Reserved for internal use.` だが、member table 自体は存在し、`Caption` / `Name` / `OnAction` / `Value` / `Select` など user-facing に使えそうな member を持つ。

## owner 設計の難所

### 1. `DialogFrame` は direct property、他 3 系統は selector 分岐が必要

- `DialogFrame` は `DialogSheet.DialogFrame -> DialogFrame` で固定できる。
- `Buttons` / `CheckBoxes` / `OptionButtons` は `DialogSheet.Buttons(Index)` の返り値が `Object` のため、以下を product 側で補わないといけない。
  - 引数省略時は collection owner (`Buttons` など)
  - 単一 selector 時は item owner (`Button` など)
  - grouped selector や複雑 selector 時は collection のまま維持

### 2. `Item` が concrete type を返さない

- `Buttons.Item(Object)`、`CheckBoxes.Item(Object)`、`OptionButtons.Item(Object)` はいずれも `Object` を返す。
- そのため `DialogSheets(1).Buttons.Item(1).Caption` を解決したい場合、`Item -> Button` のような `memberTypeOverrides` を product 側で持つ必要がある。
- これは既存 `DialogSheets.Item -> DialogSheet` と同系統の特例だが、control collection ごとに増える。

### 3. collection page に `_Dummy*` が混ざる

- `Buttons` / `CheckBoxes` / `OptionButtons` は collection page に `_Dummy*` method が残る。
- `DialogFrame` にも `_Dummy*` は存在するため、単純な owner 一括取り込みはノイズが大きい。
- 既存の `DialogSheet` common callable と同じく、allow list と skip rule を前提にする必要がある。

### 4. Worksheet / Chart 側の既存パターンとそろえる必要がある

- `WorksheetClass.Buttons(Object)` や `ChartClass.OptionButtons(Object)` も同じ `Optional Index As Object -> As Object` パターンを持つ。
- `DialogSheet` だけ個別実装すると、後で `Worksheet` / `Chart` へ広げるときに chain 解決ルールが分岐しやすい。
- selector の解釈ルールは `DialogSheet` 専用に書き捨てず、host object 共通へ寄せる前提で設計したほうがよい。

## 導入するなら必要な最小構成

### フェーズ 1: `DialogFrame` のみ

- 補助 property として `DialogSheet.DialogFrame -> DialogFrame` を追加する。
- `DialogFrame` owner には allow list を使い、`_Dummy*` を除外する。
- まず `Caption` / `Name` / `OnAction` / `Text` / `Select` 程度の最小 member だけで始める。

### フェーズ 2: single-selector な control collection

- `DialogSheet.Buttons("Button 1")` / `Buttons(1)` を `Button` owner に正規化する。
- `CheckBoxes` / `OptionButtons` も同様に、単一 selector だけ item owner へ落とす。
- `DialogSheet.Buttons()`、`Buttons(Array(...))`、複雑 selector は collection owner のまま維持する。
- `Item(Object)` には `memberTypeOverrides` を入れ、`.Item(1)` も同じ item owner に落とす。

### フェーズ 3: collection owner 自体の公開

- `Buttons` / `CheckBoxes` / `OptionButtons` collection owner を user-facing に出すかを再判断する。
- collection 自体にも `Caption` / `Value` / `Visible` のような一括操作系 member があるため、誤補完とのトレードオフを見て公開範囲を絞る。
- 初回から collection member 全面公開はしない。

## 推奨方針

- 次の実装候補は `DialogFrame` 単独プロトタイプとする。
- `Buttons` / `CheckBoxes` / `OptionButtons` は、共通 selector ルールを切り出せる見込みが立つまで docs 段階に留める。
- control collection を導入する場合でも、初回は以下を同時に入れる。
  - allow list
  - `_Dummy*` 除外
  - `Item` の `memberTypeOverrides`
  - single-selector と grouped selector の境界テスト
  - `Worksheet` / `Chart` へ横展開できるような共通 helper

## 次段タスクで見るべき点

- `DialogFrame` owner の allow list をどこまで絞るか
- `Button` / `CheckBox` / `OptionButton` へ落とす selector 判定を既存 indexed collection helper へ寄せられるか
- `Worksheet` / `Chart` の同系 method と一緒に一般化すべきか、それとも `DialogSheet` 限定で先に試すべきか
- collection owner の `Add` / `Group` / `Duplicate` のような変更系 member を user-facing に出してよいか
