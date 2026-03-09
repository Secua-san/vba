# CodeRabbit Review Workflow

## 確認対象
- PR review
- review comment
- review thread
- 差分へのインラインコメント
- 要約コメント内の重要指摘
- 最新の `coderabbitai` コメント

## 基本方針
- CodeRabbit の指摘を無条件で採用しない
- まず重複、ノイズ、誤検知の可能性を整理する
- 妥当な指摘のみ修正対象にする
- 修正理由または非採用理由を説明できる状態を保つ

## 実測メモ
- 2026-03-08 時点で、PR #11 から PR #24 の実測を確認した
- 初回の `coderabbitai` コメントは平均 0.24 分、中央値 0.22 分で、ほぼ 15 秒前後で到着している
- 一方で、`review in progress` のまま進行していた PR は、コメント更新完了まで平均 9.41 分、中央値 11.88 分かかっていた
- レート制限コメントの明示待機時間は平均 14 分 04 秒で、極端に短い 7 秒を除くと平均 16 分 24 秒だった
- 初回確認は 1 分後、その後は 2 から 3 分間隔の確認で十分だった

## 待機ルール
- 初回レビュー待ちでは、PR 作成直後に 10 秒単位で詰めて確認しない。最初の確認は 1 分後に行う
- `@coderabbitai review` を手動で実行した場合も、同じく最初の確認は 1 分後に行う
- 最新の `coderabbitai` コメントが `review in progress` を示している間は、CodeRabbit のレビュー完了を待つ
- `review in progress` の間は 2 から 3 分間隔で再確認し、マージしない
- `CodeRabbit` status context が `PENDING` の間も、最新コメントと合わせて継続確認する
- `CodeRabbit` status context が `SUCCESS` になり、最新の `coderabbitai` コメントにも未解決指摘や待機時間が無いことを確認してから次へ進む

## トリアージ分類
- 採用して修正する
- 人間確認が必要
- 誤検知または非採用

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
- `@coderabbitai review` の直後に `CodeRabbit is an incremental review system and does not re-review already reviewed commits...` という最新コメントが付いた場合、そのコメントは再レビューが弾かれており未実行であることを示す
- 手動の `@coderabbitai review` 後は、最新の `coderabbitai` コメントを見て、再レビューが受理されたのか、弾かれたのかを必ず確認する
- 手動再レビューが弾かれていた場合は、再レビュー完了扱いにしない
- 手動再レビューが受理された場合は、新しい `review in progress`、新しい run id、または更新された要約コメントを確認してから次へ進む

## レート制限時の運用
- CodeRabbit が待機時間を提示した場合は、その時間を確認してから次の行動を決める
- この待機ポリシーは、手動で再レビューを依頼した場合、または PR 単位で再レビューが必要だと判断した場合に適用する
- 待機時間の長短に関わらず、提示された待機時間はそのまま待つ
- 再レビュー前にマージしない
- 待機時間の上限は設けず、待機後に必要なら `@coderabbitai review` を再実行する
- 待機後に再レビューした結果も確認し、妥当な未解決指摘がないことを確認してから次へ進む
- 待機後の確認では、status context の変化に加えて最新の `coderabbitai` コメントを見て、再レビュー開始、進行中、未実行のどれかを判定する
- 待機後の手動再レビューでも、`incremental review system...` のコメントが返った場合は再レビュー未実行として扱う

## PR サイクル完了条件
- 初回 CodeRabbit レビューを確認済みである
- 未解決の妥当な指摘が残っていない
- 新規修正が軽微でない場合、または手動で再レビューを依頼した場合は再レビュー結果も確認済みである
- レート制限が発生した場合は、待機ポリシーに従った再レビュー確認まで完了している
- 最新の `coderabbitai` コメントを確認し、再レビューが未実行のまま完了扱いしていない
- `review in progress` のまま完了扱いしていない
- CodeRabbit 指摘のみを直した push では、`pause` 済みであることを確認している

## 自動停止条件
- 指摘同士が矛盾する
- 修正により設計方針が変わる
- 同種の指摘が 2 回以上収束しない
- 指摘内容の妥当性を自動で判定できない
