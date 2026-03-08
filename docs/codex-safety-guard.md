# Codex Safety Guard

自動コミットや自動 PR を判断するときの入口ドキュメント。

## 先に読む参照先
- ブランチ運用とコミット / PR の粒度: [process/git-workflow.md](process/git-workflow.md)
- 自動化許可条件、品質ゲート、停止条件: [process/automation-policy.md](process/automation-policy.md)
- CodeRabbit の確認フロー: [process/coderabbit-review.md](process/coderabbit-review.md)
- コミットメッセージ形式: [commit-message.md](commit-message.md)
- PR 本文テンプレート: [../.github/pull_request_template.md](../.github/pull_request_template.md)

## 即停止する条件
- `main` / `master` 直作業、または detached HEAD のまま進めようとしている
- 仕様が曖昧、または差分が複数目的で混在している
- lint / build / test の失敗理由が説明できない
- 認証、権限、課金、シークレット、CI/CD、インフラ、DB、マイグレーション、破壊的変更、依存関係のメジャー更新を含む
- 現在の変更内容を 1 文で説明できない
