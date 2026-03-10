# Microsoft Learn 組み込み署名再生成メモ

## 目的
- `WorksheetFunction` などの Microsoft Learn 由来メンバーが追加されたときに、署名データの更新箇所を迷わず追えるようにする
- 特に `XLookup` / `XMATCH` のような監視対象メンバーが Learn に掲載された際の作業を、1 本の手順へ集約する

## トリガー
- `scripts/test/mslearnReferenceAudit.test.mjs` の監視テストが失敗し、`XLookup` / `XMATCH` が Learn スナップショットへ追加されたことを示したとき
- `resources/reference/mslearn-vba-reference.json` の再生成結果に、これまで無かった `WorksheetFunction` メンバーが現れたとき
- Microsoft Learn のページを手動確認し、署名ヘルプ対象へ取り込みたいメンバーが増えたと判断したとき

## 更新箇所の一覧

| ファイル | 役割 | 更新が必要なケース |
| --- | --- | --- |
| `scripts/lib/referenceSignatureConfig.mjs` | 署名抽出対象の allow list | 新しいメソッドを署名抽出対象に加えるとき |
| `scripts/generate-mslearn-vba-reference.mjs` | Learn 取得、署名抽出、override | Learn 側の表記ゆれ、要約補正、optional/required 補正が必要なとき |
| `resources/reference/mslearn-vba-reference.json` | 生成済み参照データ | 再生成後の成果物をコミットするとき |
| `scripts/test/mslearnReferenceAudit.test.mjs` | 監視と生成データ監査 | 監視対象の状態や audit 条件を見直すとき |
| `packages/core/src/reference/builtinReference.ts` | built-in index と member 解決 | 新しい root / 返却型 / chain 解決が必要なとき |
| `packages/server/test/documentService.test.js` | server の signature/completion/hover 回帰 | server 側で新メンバーの挙動を固定するとき |
| `packages/extension/test/fixtures/BuiltInMemberSignature.bas` | extension の署名 fixture | 追加メンバーの呼び出し例を fixture に足すとき |
| `packages/extension/test/suite/index.ts` | extension の UI 回帰 | VS Code 側の signature/hover/completion を固定するとき |
| `TASKS.md` | 進捗管理 | 作業開始、完了、次候補の更新 |
| `docs/process/coderabbit-review-summaries.md` | レビュー要約ログ | PR 完了後に学びと横展開候補を残すとき |

## 標準手順

### 1. Learn 側の追加を確認する
- `npm run test:scripts` または `npm test` を実行し、監視テストの失敗内容を確認する
- `resources/reference/mslearn-vba-reference.json` で対象 owner と member を検索し、既に取得されているかを確認する
- Learn ページの `Syntax` / `Parameters` / `Return value` を見て、既存抽出ロジックで足りるかを判断する

### 2. 署名抽出対象へ追加する
- `scripts/lib/referenceSignatureConfig.mjs` の owner に対象メンバー名を追加する
- `WorksheetFunction` のように既存 owner 配下へ足すだけでよいか、別 owner の追加が必要かを確認する

### 3. 抽出ロジックの補正要否を確認する
- Learn の parameter table が連番省略、表記ゆれ、説明欠落を含む場合は `scripts/generate-mslearn-vba-reference.mjs` の共通ロジックで吸収できるかを確認する
- 個別補正が必要な場合だけ `signatureMetadataOverrides` を追加する
- `required` / `optional` の判断が Learn 表記と実利用で食い違う場合は、`docs/process/coderabbit-review.md` の運用判断基準に従う

### 4. 参照データを再生成する
- `npm run generate:reference-data` を実行する
- 生成差分で以下を確認する
  - `summary`
  - `signature.label`
  - `parameters[].dataType`
  - `parameters[].description`
  - `parameters[].isRequired`
  - `returnType`

### 5. built-in 解決への影響を確認する
- `WorksheetFunction.XLookup` のような既存 owner 配下の method 追加だけなら、通常は `packages/core/src/reference/builtinReference.ts` の追加変更は不要
- `Application.ActiveCell.Address` のように root object や member chain の型解決が絡む場合は、`typeName` や root completions の調整が必要になる
- property / event / method の区別で fallback signature を抑止する既存ルールに影響しないかも確認する

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
- `scripts/test/mslearnReferenceAudit.test.mjs` の「未掲載」監視は、掲載後に別の監視へ置き換える
- 追加メンバーを audit 対象に含め、型、説明、required/optional の欠落がないことを確認する

### 8. 品質ゲートを通す
- `npm run lint`
- `npm test`
- `npm run package`

### 9. 運用ドキュメントを更新する
- `TASKS.md` に作業内容を反映する
- PR 後に `docs/process/coderabbit-review-summaries.md` へ要約と横展開候補を追記する

## `XLookup` / `XMATCH` で最初に触る場所
- `scripts/lib/referenceSignatureConfig.mjs`
- `scripts/test/mslearnReferenceAudit.test.mjs`
- `packages/server/test/documentService.test.js`
- `packages/extension/test/fixtures/BuiltInMemberSignature.bas`
- `packages/extension/test/suite/index.ts`

## 判断メモ
- Learn と実利用で `required` / `optional` が食い違う場合は、Learn 準拠を機械的に優先しない
- 既存テスト、既存ユーザー互換、誤案内防止を優先し、判断理由を PR 本文と CodeRabbit 要約へ残す
