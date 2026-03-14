# Microsoft Learn 組み込み署名再生成メモ

## 目的
- `WorksheetFunction` や `Range` などの Microsoft Learn 由来メンバーが追加されたときに、署名データの更新箇所を迷わず追えるようにする
- `XLookup` / `XMATCH` や動的配列関連の `Range` メンバーのような監視対象が Learn に掲載された際の作業を、1 本の手順へ集約する

## トリガー
- `scripts/test/mslearnReferenceAudit.test.mjs` の監視テストが失敗し、watch list に置いていたメンバーが Learn スナップショットへ追加されたことを示したとき
- `resources/reference/mslearn-vba-reference.json` の再生成結果に、これまで無かった `WorksheetFunction` メンバーが現れたとき
- Microsoft Learn のページを手動確認し、署名ヘルプ対象へ取り込みたいメンバーが増えたと判断したとき

## 現在の優先 watch list

| owner | member | 背景 |
| --- | --- | --- |
| `WorksheetFunction` | `XLookup`, `XMATCH` | 現行 Learn スナップショットには未掲載だが、Excel の近年の lookup 系で実利用価値が高い |
| `Range` | `HasSpill`, `SavedAsArray`, `SpillParent` | 動的配列と spill 挙動の確認で使う代表的メンバーだが、現行 Learn スナップショットには未掲載 |

## 2026-03-11 の owner inventory 結果
- `Application` / `Workbook` / `Worksheet` の object page を現行 Microsoft Learn で確認し、methods / properties / events の一覧をローカル参照 JSON と照合した
- 2026-03-11 時点では、この 3 owner の object page に載っている member はローカル参照 JSON へ既に入っていたため、watch list 追加は行っていない
- 次の改善対象は、未掲載監視ではなく、`ActiveWorkbook` / `ThisWorkbook` のような root alias から既存 built-in member データへ到達できるようにすること
- `ActiveSheet` は Excel で chart sheet を返す場合もあるため、現時点では `Worksheet` 固定の `typeName` は付けず、保守的なまま維持する

## 2026-03-13 の callable 優先度見直し
- 現行 Microsoft Learn では `Workbook.SaveAs` / `Workbook.Close` / `Workbook.ExportAsFixedFormat`、`Worksheet.Evaluate` / `Worksheet.SaveAs` の各ページで syntax と parameter table を取得できることを確認した
- ただし現行の built-in root 到達性では `ActiveWorkbook` / `ThisWorkbook` から `Workbook` callable をすぐ活用できる一方、`Worksheet` callable は `ActiveSheet` を保守的に未型付けとしているため、同じ労力でも Workbook 側の効果が大きい
- そのため、callable 昇格は先に `Workbook` 側を取り込み、`Worksheet.Evaluate` / `Worksheet.SaveAs` は worksheet root 解決方針と合わせて後続へ送る

## 2026-03-13 の Worksheet callable 昇格
- `Worksheet.Evaluate` / `Worksheet.SaveAs` / `Worksheet.ExportAsFixedFormat` を allow list へ追加し、署名データを再生成した
- `ActiveSheet` は引き続き未型付けのまま維持し、代わりに `Worksheets(1)` / `ActiveWorkbook.Worksheets(1)` のような indexed collection access を built-in owner 解決で `Worksheet` として扱うようにした
- collection index 付き member access を扱う場合は、parser 側の path 解決と collection item type の対応表をセットで更新し、`ActiveSheet` のような曖昧 root には拡げない
- 文字列リテラルや単純式の selector は単一 `Worksheet` として扱える一方、`Array(...)` や関数呼び出し selector は複数シートや不明型を返し得るため、`Worksheets` collection のまま維持する

## 2026-03-13 の DialogSheet 調査結果
- Office VBA 側では `DialogSheetView` object は取得できるが、`DialogSheet` object page は参照 JSON へ入っていない
- `Refer to Sheets by Name` では `DialogSheets("Dialog1").Activate` が示され、dialog sheet 自体は VBA から参照できることを確認した
- Microsoft Learn の .NET interop `DialogSheet` page には member 一覧があるが、`Reserved for internal use.` と `dummy` member を含むため、現行の VBA 参照 JSON へそのまま取り込む正本にはしない
- そのため、`VB_Base = 0{00020830-0000-0000-C000-000000000046}` の dialog sheet document module は、owner 未公開のまま保守動作を維持する
- interop 補助ソースとしての導入可否と制約は `docs/process/dialogsheet-interop-source-feasibility.md` に切り出して管理する

## 2026-03-13 の DialogSheet common callable プロトタイプ
- `scripts/lib/supplementalReferenceConfig.mjs` に `DialogSheet` 用の interop allow list と `DialogSheets` collection clone 設定を追加した
- 補助ソースの user-facing root は `DialogSheets(1)` とし、`DialogSheet1.` の document module root は引き続き built-in owner へ昇格させない
- `DialogSheet` owner は interop page の `Methods` table から allow list だけを抜き出し、`_Dummy*` と `_SaveAs` などの `_` 接頭辞 member は生成時に除外する
- `DialogSheets` collection owner は `Sheets` の member を clone して使い、`Item` だけは `typeName: "DialogSheet"` を明示して `DialogSheets.Item(1)` も単一 selector と同じ root 解決にそろえる
- 監査テストでは `dummy` / legacy member 混入、正規化後の重複名、allow list member の署名欠落、`DialogSheets.Item` の `typeName` 欠落を検知する

## 2026-03-13 の DialogSheet Workbook / Application root 展開
- `ApplicationClass.DialogSheets` / `WorkbookClass.DialogSheets` は Microsoft Learn interop で `ReadOnly Property DialogSheets As Sheets` を示す
- 生成器では `Application.DialogSheets` / `Workbook.DialogSheets` を補助 property として追加し、`typeName` は user-facing root を保つため `DialogSheets` へ正規化する
- `As Sheets` をそのまま採用すると `DialogSheet` item owner へ到達できないため、collection clone と property alias を組み合わせて `Application.DialogSheets(1)` / `ActiveWorkbook.DialogSheets(1)` を単一 selector root として扱う
- grouped selector は従来どおり collection のまま維持し、`Application.DialogSheets(Array(...)).SaveAs` のような誤補完は出さない
- 監査テストでは `Application.DialogSheets` / `Workbook.DialogSheets` の `typeName: "DialogSheets"` を固定し、root 展開が壊れたら検知する

## 2026-03-13 の DialogSheet control collection 調査結果
- `DialogSheet.DialogFrame` は interop で `As DialogFrame` を返すため、補助 property として最も導入しやすい
- `DialogSheet.Buttons` / `CheckBoxes` / `OptionButtons` は `Optional Index As Object -> As Object` なので、単一 selector と collection selector を product 側で分岐しないと user-facing owner を決められない
- `Buttons` / `CheckBoxes` / `OptionButtons` collection page は `_Dummy*` を含み、`Item(Object)` も `Object` を返すため、導入するなら allow list、`memberTypeOverrides`、grouped selector 抑止をセットで入れる
- 調査の正本は `docs/process/dialogsheet-control-collection-feasibility.md` に切り出して管理する

## 2026-03-14 の Worksheet / Chart control collection 方針
- Office VBA の `Using ActiveX Controls on Sheets` は、worksheet / chart sheet 上の control を `OLEObjects`、`Shapes`、control code name で扱う導線を正本としている
- 一方で `WorksheetClass.Buttons` / `CheckBoxes` / `OptionButtons`、`ChartClass.Buttons` / `CheckBoxes` / `OptionButtons` は interop 側で `Optional Index As Object -> As Object` を返し、`DialogSheet` と同じ owner 曖昧性を持つ
- そのため、`Worksheet` / `Chart` への横展開は現段階では supplemental interop source の候補に留め、`Buttons` 系 collection の公開より `OLEObjects` / control name 導線の整理を優先する
- 正本の判断メモは `docs/process/worksheet-chart-control-collection-feasibility.md` に切り出して管理する

## owner 候補の選び方
- まず、`packages/core/src/reference/builtinReference.ts` の root object から到達しやすい owner を優先する
- 次に、最新 Excel で利用頻度が高い機能領域を優先する。現時点では lookup と動的配列を最優先とする
- 既に Learn スナップショットへ載っているメンバーは watch list に入れず、allow list 追加または built-in 解決の検討へ進める
- watch list へ入れるメンバーは、Microsoft Learn に個別ページがあり、将来スナップショットへ取り込まれた時点で署名化したいものに絞る
- Office VBA object page が無く、interop page しか無い owner は、そのまま allow list へ入れず、まず `Reserved for internal use` と `dummy` member の扱いを設計判断として切り分ける

## 更新箇所の一覧

| ファイル | 役割 | 更新が必要なケース |
| --- | --- | --- |
| `scripts/lib/referenceSignatureConfig.mjs` | 署名抽出対象の allow list と未掲載監視の watch list | 新しいメソッドを署名抽出対象に加えるとき、または未掲載監視を追加・解除するとき |
| `scripts/lib/supplementalReferenceConfig.mjs` | interop 補助ソースの allow list、clone、補助 property 設定 | `DialogSheet` のように Office VBA object page が無い owner を限定導入するとき |
| `scripts/generate-mslearn-vba-reference.mjs` | Learn 取得、署名抽出、override | Learn 側の表記ゆれ、要約補正、optional/required 補正が必要なとき |
| `resources/reference/mslearn-vba-reference.json` | 生成済み参照データ | 再生成後の成果物をコミットするとき |
| `scripts/test/mslearnReferenceAudit.test.mjs` | 監視と生成データ監査 | 監視対象の状態や audit 条件を見直すとき |
| `packages/core/src/reference/builtinReference.ts` | built-in index と member 解決 | 新しい root / 返却型 / chain 解決が必要なとき |
| `packages/server/test/documentService.test.js` | server の signature/completion/hover 回帰 | server 側で新メンバーの挙動を固定するとき |
| `packages/extension/test/fixtures/BuiltInMemberSignature.bas` | extension の署名 fixture | 追加メンバーの呼び出し例を fixture に足すとき |
| `packages/extension/test/suite/index.ts` | extension の UI 回帰 | VS Code 側の signature/hover/completion を固定するとき |
| `TASKS.md` | 進捗管理 | 作業開始、完了、次候補の更新 |
| `docs/process/coderabbit-review-summaries.md` | レビュー要約ログの入口 | PR 完了後に当月ログを確認・更新するとき |

## 標準手順

### 1. Learn 側の追加を確認する
- `npm run test:scripts` または `npm test` を実行し、監視テストの失敗内容を確認する
- `resources/reference/mslearn-vba-reference.json` で対象 owner と member を検索し、既に取得されているかを確認する
- Learn ページの `Syntax` / `Parameters` / `Return value` を見て、既存抽出ロジックで足りるかを判断する
- `scripts/lib/referenceSignatureConfig.mjs` の watch list に残っている対象が失敗した場合は、以降の手順で watch list から allow list への移動を行う

### 2. 署名抽出対象へ追加する
- `scripts/lib/referenceSignatureConfig.mjs` の watch list から対象メンバー名を外し、allow list の owner に追加する
- `WorksheetFunction` のように既存 owner 配下へ足すだけでよいか、別 owner の追加が必要かを確認する
- 追加後に watch list と allow list の重複が無いことを確認する
- Office VBA object page が無く interop 補助ソースを使う場合は、`scripts/lib/supplementalReferenceConfig.mjs` に allow list を追加し、owner page 側の table から余計な member が混ざらないことを先に固定する
- collection clone を併用する場合は、`Item` などの返却型が推論に必要な member を `memberTypeOverrides` で明示し、clone 元 page の構造変化で欠落したときは生成を失敗させる

### 3. 抽出ロジックの補正要否を確認する
- Learn の parameter table が連番省略、表記ゆれ、説明欠落を含む場合は `scripts/generate-mslearn-vba-reference.mjs` の共通ロジックで吸収できるかを確認する
- 個別補正が必要な場合だけ `signatureMetadataOverrides` を追加する
- `Workbook.Close` のように Learn 側で戻り値節を持たない `Sub` 相当 member は、生成データでは `returnType: "Void"` を保持しつつ、表示ラベルには `As Void` を出さない
- `required` / `optional` の判断が Learn 表記と実利用で食い違う場合は、`docs/process/coderabbit-review.md` の運用判断基準に従う
- interop member page は `Syntax` ではなく `Definition` の `vb` code block を持つため、owner ごとに補助ソース parser を足す場合は `Definition` / `Parameters` の構造差も確認する
- 補助ソース owner の allow list member で署名抽出に失敗した場合は、黙って JSON を出さず生成を失敗させる

### 4. 参照データを再生成する
- `npm run generate:reference-data` を実行する
- 生成差分で以下を確認する
  - `summary`
  - `signature.label`
  - `parameters[].dataType`
  - `parameters[].description`
  - `parameters[].isRequired`
  - `returnType`
- PR 前に、対象 owner 以外の `resources/reference/mslearn-vba-reference.json` 差分が混入していないことを確認する

### 5. built-in 解決への影響を確認する
- `WorksheetFunction.XLookup` のような既存 owner 配下の method 追加だけなら、通常は `packages/core/src/reference/builtinReference.ts` の追加変更は不要
- `Application.ActiveCell.Address` のように root object や member chain の型解決が絡む場合は、`typeName` や root completions の調整が必要になる
- property / event / method の区別で fallback signature を抑止する既存ルールに影響しないかも確認する
- collection root を追加する場合は、単一 selector だけ item owner に落とすのか、grouped selector は collection のまま維持するのかを `INDEXED_COLLECTION_OWNER_TYPES` とテストで同時に固定する

### 6. server / extension の回帰を追加する
- `packages/server/test/documentService.test.js`
  - signature help の取得
  - hover の Learn URL / summary 表示
  - optional/required 判定
- `packages/extension/test/fixtures/BuiltInMemberSignature.bas`
  - 対象メンバーの呼び出し例を追加
- `packages/extension/test/suite/index.ts`
  - VS Code から見える signature/hover/completion の期待値を追加
- variadic な `WorksheetFunction` 系は、先頭だけでなく末尾引数も確認する

### 7. 監視テストを更新する
- `scripts/lib/referenceSignatureConfig.mjs` の watch list から、掲載済みメンバーを外す
- `scripts/test/mslearnReferenceAudit.test.mjs` は watch list 定義を参照しているため、通常は個別の member 名編集は不要
- 追加メンバーを audit 対象に含め、型、説明、required/optional の欠落がないことを確認する

### 8. 品質ゲートを通す
- `npm run lint`
- `npm test`
- `npm run package`

### 9. 運用ドキュメントを更新する
- `TASKS.md` に作業内容を反映する
- PR 後に `docs/process/coderabbit-review-summaries.md` の案内に従って当月ログへ要約と横展開候補を追記する

## `XLookup` / `XMATCH` で最初に触る場所
- `scripts/lib/referenceSignatureConfig.mjs`
- `scripts/test/mslearnReferenceAudit.test.mjs`
- `packages/server/test/documentService.test.js`
- `packages/extension/test/fixtures/BuiltInMemberSignature.bas`
- `packages/extension/test/suite/index.ts`

## `Range` 動的配列メンバーで最初に触る場所
- `scripts/lib/referenceSignatureConfig.mjs`
- `scripts/test/mslearnReferenceAudit.test.mjs`
- `packages/core/src/reference/builtinReference.ts`
- `packages/server/test/documentService.test.js`
- `packages/extension/test/suite/index.ts`

## 判断メモ
- Learn と実利用で `required` / `optional` が食い違う場合は、Learn 準拠を機械的に優先しない
- 既存テスト、既存ユーザー互換、誤案内防止を優先し、判断理由を PR 本文と CodeRabbit 要約へ残す
