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
- [worksheet-chart-control-entrypoint-feasibility.md](worksheet-chart-control-entrypoint-feasibility.md): Worksheet / Chart 上の control をどの entry point から実装するかの整理
- [worksheet-chart-control-identity-feasibility.md](worksheet-chart-control-identity-feasibility.md): `OLEObject.Object` と control code name 支援に必要な metadata source の整理
- [shape-oleformat-object-promotion-feasibility.md](shape-oleformat-object-promotion-feasibility.md): `Shape.OLEFormat.Object` を control owner へ昇格できる条件と除外条件の整理
- [shape-oleformat-object-explicit-sheet-root-feasibility.md](shape-oleformat-object-explicit-sheet-root-feasibility.md): `Worksheets("Sheet1")` / `ThisWorkbook.Worksheets("Sheet1")` 系 root を sidecar へ結べる範囲と join key の整理
- [explicit-sheet-name-broad-root-feasibility.md](explicit-sheet-name-broad-root-feasibility.md): `ActiveWorkbook.Worksheets("Sheet1")` の runtime gating 済み境界と、unqualified `Worksheets("Sheet1")` / `Application.Worksheets("Sheet1")` を同じ broad root family として扱う条件整理
- [workbook-binding-manifest-feasibility.md](workbook-binding-manifest-feasibility.md): broad root 再評価の前提となる workbook binding manifest の置き場所、key、保守条件の整理
- [active-workbook-identity-provider-contract.md](active-workbook-identity-provider-contract.md): host / extension / server で共有する active workbook identity snapshot と gating 条件の整理
- [worksheet-chart-shapes-root-feasibility.md](worksheet-chart-shapes-root-feasibility.md): `Shapes(Index)` / `Shapes.Item(Index)` を generic `Shape` surface として扱う境界の整理
- [worksheet-chart-control-metadata-source-poc.md](worksheet-chart-control-metadata-source-poc.md): workbook package を起点にした control metadata PoC と sidecar 方針の整理
- [worksheet-control-metadata-sidecar-artifact.md](worksheet-control-metadata-sidecar-artifact.md): loose files と併用する sidecar schema、保存場所、lookup ルールの整理
- [sub-agent-escalation.md](sub-agent-escalation.md): サブエージェント利用の切り替え基準

## 取りこぼしを減らす順番

1. 先に要件か ADR で「何を守るか」を確認する。
2. 次にこのディレクトリで「どう進めるか」を確認する。
3. 実作業後に `TASKS.md` とレビュー要約ログへ必要な更新だけを残す。

## レビュー履歴の扱い

- CodeRabbit の履歴は [coderabbit-review-summaries.md](coderabbit-review-summaries.md) を入口にして、当月ログだけを開く。
- 月をまたぐときは `docs/process/coderabbit-review-logs/YYYY-MM.md` を追加し、入口からリンクする。
- ルールと履歴を同じファイルへ混ぜない。運用判断は `coderabbit-review.md`、PR ごとの学びは月次ログへ残す。
