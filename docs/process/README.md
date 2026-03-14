# 運用ガイド

`docs/process/` の入口。コミット、PR、レビュー、再生成、運用変更では最初にここを読む。

## 最短動線

| 作業 | 読む順 |
| --- | --- |
| 実装を始める | [要件書](../requirements/000-overview.md) -> [ADR 一覧](../adr/README.md) -> [Git Workflow](git-workflow.md) -> 関連手順書 -> [TASKS](../../TASKS.md) |
| 自動コミット / 自動 PR 判断 | [Automation Policy](automation-policy.md) -> [Git Workflow](git-workflow.md) -> [CodeRabbit Review](coderabbit-review.md) |
| コミットメッセージを書く | [コミットメッセージ規約](../commit-message.md) -> [例](commit-message-examples.md) |
| PR を作る / 閉じる | [Git Workflow](git-workflow.md) -> [PR テンプレート](../../.github/pull_request_template.md) -> [CodeRabbit Review](coderabbit-review.md) -> [レビュー要約ログ案内](coderabbit-review-summaries.md) |
| Microsoft Learn 署名を更新する | [Microsoft Learn 組み込み署名再生成メモ](mslearn-signature-regeneration.md) |
| 判断が空転した | [Sub-Agent Escalation](sub-agent-escalation.md) |

## 正本文書

- [git-workflow.md](git-workflow.md): ブランチ、コミット粒度、PR の必須事項
- [automation-policy.md](automation-policy.md): 自動コミット / 自動 PR の許可条件、品質ゲート、停止条件
- [coderabbit-review.md](coderabbit-review.md): CodeRabbit の待機、トリアージ、再レビュー、完了条件
- [coderabbit-review-summaries.md](coderabbit-review-summaries.md): レビュー要約ログの入口と当月ログの案内
- [mslearn-signature-regeneration.md](mslearn-signature-regeneration.md): Microsoft Learn 由来署名データの更新手順
- [dialogsheet-interop-source-feasibility.md](dialogsheet-interop-source-feasibility.md): DialogSheet interop 補助ソースの可否調査
- [dialogsheet-control-collection-feasibility.md](dialogsheet-control-collection-feasibility.md): DialogSheet control collection の導入段階と owner 設計の整理
- [worksheet-chart-control-collection-feasibility.md](worksheet-chart-control-collection-feasibility.md): Worksheet / Chart control collection を横展開する前提と優先順位の整理
- [sub-agent-escalation.md](sub-agent-escalation.md): サブエージェント利用の切り替え基準

## 取りこぼしを減らす順番

1. 先に要件か ADR で「何を守るか」を確認する。
2. 次にこのディレクトリで「どう進めるか」を確認する。
3. 実作業後に `TASKS.md` とレビュー要約ログへ必要な更新だけを残す。

## レビュー履歴の扱い

- CodeRabbit の履歴は [coderabbit-review-summaries.md](coderabbit-review-summaries.md) を入口にして、当月ログだけを開く。
- 月をまたぐときは `docs/process/coderabbit-review-logs/YYYY-MM.md` を追加し、入口からリンクする。
- ルールと履歴を同じファイルへ混ぜない。運用判断は `coderabbit-review.md`、PR ごとの学びは月次ログへ残す。
