---
name: auto-commit-pr
description: Safely prepare commits and pull requests for this repository. Use when implementation is complete and Codex needs to inspect repository state, choose or create a compliant branch, split diffs into logical commits, run quality gates, draft PR text, or triage CodeRabbit feedback.
---

# Auto Commit PR

このスキルでは、コミットと PR 作成の骨子だけを扱う。詳細ルールは参照ドキュメントから必要なものだけ読む。

## 読む参照先
- ブランチ運用、コミット粒度、PR の範囲: [../../docs/process/git-workflow.md](../../docs/process/git-workflow.md)
- 自動コミット / 自動 PR の許可条件、品質ゲート、停止条件、リポジトリ作成ルール: [../../docs/process/automation-policy.md](../../docs/process/automation-policy.md)
- CodeRabbit の確認と再レビューサイクル: [../../docs/process/coderabbit-review.md](../../docs/process/coderabbit-review.md)
- 判断が空転した場合のサブエージェント利用: [../../docs/process/sub-agent-escalation.md](../../docs/process/sub-agent-escalation.md)
- コミットメッセージ形式と例: [../../docs/commit-message.md](../../docs/commit-message.md), [../../docs/process/commit-message-examples.md](../../docs/process/commit-message-examples.md)
- PR 本文テンプレート: [../../.github/pull_request_template.md](../../.github/pull_request_template.md)

## 手順
1. 現在のブランチ、変更済みファイル、未追跡ファイル、detached HEAD、`main` / `master` 直作業、リポジトリ初期化状態を確認する。
2. 既存の適切なブランチを再利用できるか判断し、必要なら `<type>/<short-summary-kebab-case>` 形式で新規ブランチを作成して切り替える。
3. 差分を意図単位で分類し、実装、リファクタ、整形、生成物、設定、ドキュメント、無関係修正を混在させない。
4. `../../docs/process/automation-policy.md` に従って品質ゲートを実行し、失敗理由が不明または高リスクなら停止する。
5. 実差分に基づいて `type(scope): summary` を作成し、1 コミット 1 意図でコミットする。
6. 自動 PR 条件を満たす場合だけ PR を作成し、テンプレートの必須項目を埋める。
7. PR 作成後は CodeRabbit を確認し、妥当な低リスク指摘のみ採用し、必要なら別コミットで修正して品質ゲートを再実行する。
8. 差分分類、リスク判断、または CodeRabbit トリアージが空転したら、論点を絞ってサブエージェントへ意見を求める。

## 即停止する条件
- 仕様、ブランチ戦略、または差分分割方針が曖昧
- 禁止領域に触れている
- lint / build / test の失敗理由を説明できない
- リポジトリ状態が不整合で安全に自動化できない
