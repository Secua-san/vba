# Codex Safety Guard

自動コミットや自動 PR を判断するときの軽い入口チェック。詳細ルールは `docs/process/` の正本を読む。

## 先に読む参照先
- 運用ドキュメントの入口: [process/README.md](process/README.md)
- docs 全体の入口: [README.md](README.md)
- PR 本文テンプレート: [../.github/pull_request_template.md](../.github/pull_request_template.md)

## 即停止する条件
- `main` / `master` 直作業、または detached HEAD のまま進めようとしている
- 仕様が曖昧、または差分が複数目的で混在している
- lint / build / test の失敗理由が説明できない
- 認証、権限、課金、シークレット、CI/CD、インフラ、DB、マイグレーション、破壊的変更、依存関係のメジャー更新を含む
- 現在の変更内容を 1 文で説明できない
