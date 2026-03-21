# CodeRabbit Review Workflow

## 確認対象
- PR review
- review comment
- review thread
- 差分へのインラインコメント
- 要約コメント内の重要指摘
- 最新の `coderabbitai` コメント
- `CodeRabbit` status context

## 基本方針
- CodeRabbit の指摘を無条件で採用しない
- まず重複、ノイズ、誤検知の可能性を整理する
- 妥当な指摘のみ修正対象にする
- 修正理由または非採用理由を説明できる状態を保つ
- PR 前の自己レビューと CodeRabbit が同じ論点を独立に指摘した場合、その指摘は原則として採用し、修正する
- 特に `required` / `optional` のような運用時の挙動に関わる指摘は、この文書の判断基準に従って、互換性と実利用を優先して修正する
- ドキュメント由来データと運用判断がずれる場合は、PR 本文に理由を残す
- レビュー記録を残す場合は `docs/process/coderabbit-review-logs/YYYY-MM.md` に直接追記する
- レビュー記録は証跡専用とし、運用判断や参照順の正本にしない

## PR 作成前のセルフレビュー（サブエージェント）
- `gh pr create` の前に、必ずサブエージェントで差分レビューを 1 回実行する
- 既定では `reviewer` を使い、必要に応じて `default` で補助観点を追加する
- `reviewer` が利用できない状態では PR 作成を進めず、`C:\Users\tagi0\.codex\config.toml` と `C:\Users\tagi0\.codex\agents\reviewer.toml` の設定を確認する
- 観点は、回帰リスク、境界条件、テスト不足、不要差分の 4 点を最低限含める
- サブエージェントの指摘を主担当がトリアージし、重大度の高い指摘（機能不整合、明確なバグ、テスト欠落）は PR 作成前に解消する
- CodeRabbit 受領後は、自己レビュー指摘との重複有無を必ず照合し、同一論点なら修正対象として扱う
- サブエージェント結果の要約を残してから PR を作成する

## 待機ルール
- 初回レビュー待ちでは、PR 作成直後に 10 秒単位で詰めて確認しない。最初の確認は 1 分後に行う
- `@coderabbitai review` を手動で実行した場合も、同じく最初の確認は 1 分後に行う
- 最新の `coderabbitai` コメントが `review in progress` を示している間は、CodeRabbit のレビュー完了を待つ
- `review in progress` の間は 2 から 3 分間隔で再確認し、マージしない
- `CodeRabbit` status context が `PENDING` の間も、最新コメントと合わせて継続確認する
- `CodeRabbit` status context が `SUCCESS` になり、最新の `coderabbitai` コメントにも未解決指摘や待機時間が無いことを確認してから次へ進む
- CodeRabbit が待機時間を提示した場合は、その時間をそのまま待つ
- 待機時間の上限は設けず、必要なら待機後に `@coderabbitai review` を再実行する
- `@coderabbitai review` の直後に `CodeRabbit is an incremental review system and does not re-review already reviewed commits...` という最新コメントが付いた場合、その再レビューは未実行として扱う

## トリアージ分類
- 採用して修正する
- 人間確認が必要
- 誤検知または非採用

## 重複指摘の扱い
- PR 前自己レビューと CodeRabbit が同じ内容を指摘した場合は、独立した一致として重く扱い、原則修正する
- 同じ内容とは、表現が違っても同じ根本原因や同じ挙動差を指していれば含む
- `required` / `optional`、署名、補完候補、診断の有無のような実利用に直結する挙動差は、重複指摘ならこの文書の判断基準でより妥当な挙動を正とする
- 運用上妥当な挙動を優先した結果、Microsoft Learn などの出典と差が出る場合は、非採用ではなく「運用優先の採用」として記録する

## 運用判断の基準
- `required` / `optional` などの挙動判断は、まず既存の拡張挙動、既存テスト期待値、既存ユーザーへの互換性を優先して比較する
- 次に、誤った補完、誤診断、誤シグネチャ表示を減らせるかどうかを確認し、実利用での誤案内が少ない側を採用する
- 同系統の既存メンバーや生成ルールと一貫するかを確認し、個別例外よりも横展開しやすい側を優先する
- 出典との差分があっても、上記の観点で運用優先と判断した場合は修正し、その理由を PR 本文に残す
- 互換性、既存テスト、ユーザー影響のどれを優先すべきか自動で決められない場合は、人間確認へ切り替える

## 人間確認を優先するケース
- 仕様判断が必要
- 設計方針の変更を伴う
- セキュリティ影響がある
- 認証、権限、課金、CI/CD、インフラ、DB 構造変更に関わる

## 修正時のルール
- CodeRabbit 対応は可能な限り別コミットに分ける
- CodeRabbit 対応と無関係な変更を混ぜない
- 必要に応じてテストを追加する
- 修正後は lint / build / test を再実行する
- CodeRabbit の指摘だけを直す push では、push 前に `@coderabbitai pause` を実行する
- `pause` 後の CodeRabbit 指摘修正は、そのまま push し、再レビューを待たずにマージする
- CodeRabbit 指摘修正の push では、原則として `@coderabbitai review` を追加で投げない

## 再レビューの使い分け
- `.coderabbit.yaml` では `auto_incremental_review: false` を維持し、push ごとの自動再レビューは行わせない
- 初回 PR レビューは通常どおり確認する
- レビュー対応の修正が軽微な場合は、そのまま push し、`@coderabbitai review` を追加で投げない
- 軽微な修正とは、命名、コメント、狭い条件分岐、既存指摘を満たすだけの小さな追加テストのように、仕様や挙動を広げない変更を指す
- 仕様、制御フロー、診断条件、公開 API、CI 設定などに影響する修正は軽微扱いにせず、push 後に `@coderabbitai review` を手動で実行する
- PR 単位で CodeRabbit の反応自体を止めたい場合は `@coderabbitai pause`、再開したい場合は `@coderabbitai resume` を使う
- PR の状況確認時は、status context や review 一覧だけでなく、PR 上の最新の `coderabbitai` コメントも必ず確認する
- 手動の `@coderabbitai review` 後は、最新の `coderabbitai` コメントを見て、再レビューが受理されたのか、弾かれたのかを必ず確認する
- 手動再レビューが受理された場合は、新しい `review in progress`、新しい run id、または更新された要約コメントを確認してから次へ進む

## PR サイクル完了条件
- 初回 CodeRabbit レビューを確認済みである
- 未解決の妥当な指摘が残っていない
- 新規修正が軽微でない場合、または手動で再レビューを依頼した場合は再レビュー結果も確認済みである
- レート制限が発生した場合は、待機時間経過後の再レビュー確認まで完了している
- 最新の `coderabbitai` コメントを確認し、再レビューが未実行のまま完了扱いしていない
- `review in progress` のまま完了扱いしていない
- CodeRabbit 指摘のみを直した push では、`pause` 済みであることを確認している

## 自動停止条件
- 指摘同士が矛盾する
- 修正により設計方針が変わる
- 同種の指摘が 2 回以上収束しない
- 指摘内容の妥当性を自動で判定できない
