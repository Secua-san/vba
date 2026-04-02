# Microsoft Learn 組み込み署名再生成メモ

`WorksheetFunction` や `Range` などの Microsoft Learn 由来メンバーを更新するときの正本手順。履歴メモは置かず、作業に必要な判断だけを残す。

## この文書を開くとき

- `scripts/test/mslearnReferenceAudit.test.mjs` の監視テストが失敗したとき
- `resources/reference/mslearn-vba-reference.json` の再生成差分に新しい owner / member が出たとき
- 手動で Microsoft Learn を確認し、署名ヘルプ対象へ追加したいメンバーが増えたとき

## 必要なときだけ開く関連文書

- DialogSheet 補助ソース: [dialogsheet-interop-source-feasibility.md](dialogsheet-interop-source-feasibility.md), [dialogsheet-control-collection-feasibility.md](dialogsheet-control-collection-feasibility.md)
- Worksheet / Chart control owner: [worksheet-chart-control-collection-feasibility.md](worksheet-chart-control-collection-feasibility.md), [worksheet-chart-control-entrypoint-feasibility.md](worksheet-chart-control-entrypoint-feasibility.md), [worksheet-chart-control-identity-feasibility.md](worksheet-chart-control-identity-feasibility.md)
- `required` / `optional` などの運用判断: [coderabbit-review.md](coderabbit-review.md)

## 現在の優先 watch list

| owner | member | 背景 |
| --- | --- | --- |
| `WorksheetFunction` | `XLookup`, `XMATCH` | 現行 Learn スナップショットには未掲載だが、lookup 系で実利用価値が高い |
| `Range` | `HasSpill`, `SavedAsArray`, `SpillParent` | 動的配列と spill 挙動で使う代表メンバーだが、現行 Learn スナップショットには未掲載 |

## まず触るファイル

| ファイル | 役割 | 触るとき |
| --- | --- | --- |
| `scripts/lib/referenceSignatureConfig.mjs` | allow list と watch list | 対象メンバーの追加、監視解除 |
| `scripts/lib/supplementalReferenceConfig.mjs` | interop 補助ソース設定 | Office VBA object page が無い owner を限定導入するとき |
| `scripts/generate-mslearn-vba-reference.mjs` | Learn 取得、署名抽出、override | Learn 側の表記ゆれや補正が必要なとき |
| `resources/reference/mslearn-vba-reference.json` | 生成済み参照データ | 再生成結果を反映するとき |
| `scripts/test/mslearnReferenceAudit.test.mjs` | 監視と監査 | watch list や監査条件を見直すとき |
| `packages/core/src/reference/builtinReference.ts` | built-in root と member 解決 | root / chain 解決へ影響するとき |
| `packages/server/test/documentService.test.js` | server 回帰 | signature / hover / completion を固定するとき |
| `packages/extension/test/fixtures/BuiltInMemberSignature.bas` | extension fixture | 呼び出し例を追加するとき |
| `packages/extension/test/suite/index.ts` | extension UI 回帰 | VS Code 側の期待値を固定するとき |
| `TASKS.md` | 日常参照用の進捗要点 | 作業開始、完了、次候補、重要事項の更新 |
| `TASKLOG.md` | 詳細な履歴ログ | docs-only の判断記録や長い補足を残すとき |

## 標準手順

### 1. Learn 側の追加を確認する

- `npm run test:scripts` または `npm test` を実行し、監視テストの失敗内容を確認する
- `resources/reference/mslearn-vba-reference.json` で対象 owner と member を検索し、既に取得済みかを確認する
- Learn ページの `Syntax`、`Parameters`、`Return value` を見て、既存抽出ロジックで足りるかを確認する

### 2. watch list と allow list を整理する

- `scripts/lib/referenceSignatureConfig.mjs` の watch list から対象メンバーを外し、allow list へ追加する
- 追加後に watch list と allow list が重複していないことを確認する
- Office VBA object page が無い owner は、そのまま追加せず関連メモで補助ソース化の条件を確認する

### 3. 抽出ロジックの補正要否を確認する

- Learn の parameter table に表記ゆれ、説明欠落、連番省略がある場合は共通ロジックで吸収できるかを先に確認する
- 個別補正が必要な場合だけ `signatureMetadataOverrides` を追加する
- `Workbook.Close` のような `Sub` 相当 member は、生成データでは `returnType: "Void"` を保持しつつ表示ラベルには `As Void` を出さない
- interop 補助ソースを使う場合は allow list member の署名抽出失敗を黙殺せず、生成を失敗させる

### 4. 参照データを再生成する

- `npm run generate:reference-data` を実行する
- 生成差分で `summary`、`signature.label`、`parameters[].dataType`、`parameters[].description`、`parameters[].isRequired`、`returnType` を確認する
- 対象 owner 以外の `resources/reference/mslearn-vba-reference.json` 差分が混入していないことを確認する

### 5. built-in 解決への影響を確認する

- 既存 owner 配下のメンバー追加だけなら、通常は `packages/core/src/reference/builtinReference.ts` の変更は不要
- root object や member chain の型解決が絡む場合は、`typeName`、root completion、collection item owner を確認する
- grouped selector を collection のまま維持するか、単一 selector だけ item owner に落とすかをテストと一緒に固定する

### 6. server / extension の回帰を追加する

- `packages/server/test/documentService.test.js` で signature help、hover の Learn URL / summary、optional / required 判定を固定する
- `packages/extension/test/fixtures/BuiltInMemberSignature.bas` に呼び出し例を追加する
- `packages/extension/test/suite/index.ts` で VS Code から見える signature、hover、completion を固定する
- variadic な `WorksheetFunction` 系は先頭だけでなく末尾引数も確認する

### 7. 監視テストを更新する

- `scripts/lib/referenceSignatureConfig.mjs` の watch list から掲載済みメンバーを外す
- `scripts/test/mslearnReferenceAudit.test.mjs` は watch list 定義を参照しているため、通常は個別 member 名の直接編集は不要
- 追加メンバーを audit 対象に含め、型、説明、required / optional の欠落がないことを確認する

### 8. 品質ゲートを通す

- `npm run lint`
- `npm test`
- `npm run package`

### 9. 運用文書を更新する

- `TASKS.md` に作業要点を反映する
- 長い判断メモや docs-only の履歴が必要な場合だけ `TASKLOG.md` に追記する
- Learn と実利用の判断差がある場合は PR 本文へ理由を残す

## 判断原則

- Learn と実利用で `required` / `optional` が食い違う場合は、Learn 準拠を機械的に優先しない
- 既存テスト、既存ユーザー互換、誤案内防止を優先する
- 補助ソース owner は、関連メモで公開条件が固まるまで不用意に広げない
