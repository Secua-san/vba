# 運用ガイド

`docs/process/` の入口。最初はコア文書だけを開き、必要になったときだけ機能別メモを追加で開く。

## コア文書

| 作業 | 開く文書 |
| --- | --- |
| ブランチ / コミット / PR | [git-workflow.md](git-workflow.md) |
| コミットメッセージ | [コミットメッセージ規約](../commit-message.md) |
| コミットメッセージ例が必要 | [commit-message-examples.md](commit-message-examples.md) |
| CodeRabbit 対応 | [coderabbit-review.md](coderabbit-review.md) |
| 自動コミット / 自動 PR 判断 | [automation-policy.md](automation-policy.md) |
| Microsoft Learn 由来データの再生成 | [mslearn-signature-regeneration.md](mslearn-signature-regeneration.md) |
| 判断が空転した | [sub-agent-escalation.md](sub-agent-escalation.md) |

## 機能別メモ

- DialogSheet 系: [dialogsheet-interop-source-feasibility.md](dialogsheet-interop-source-feasibility.md), [dialogsheet-control-collection-feasibility.md](dialogsheet-control-collection-feasibility.md)
- Worksheet / Chart control 系: [worksheet-chart-control-collection-feasibility.md](worksheet-chart-control-collection-feasibility.md), [worksheet-chart-control-entrypoint-feasibility.md](worksheet-chart-control-entrypoint-feasibility.md), [worksheet-chart-control-identity-feasibility.md](worksheet-chart-control-identity-feasibility.md), [worksheet-chart-control-metadata-source-poc.md](worksheet-chart-control-metadata-source-poc.md), [worksheet-control-metadata-sidecar-artifact.md](worksheet-control-metadata-sidecar-artifact.md), [worksheet-chart-shapes-root-feasibility.md](worksheet-chart-shapes-root-feasibility.md), [worksheet-control-shared-spec-family-candidate-feasibility.md](worksheet-control-shared-spec-family-candidate-feasibility.md), [worksheet-control-shape-name-path-vocabulary-feasibility.md](worksheet-control-shape-name-path-vocabulary-feasibility.md), [worksheet-control-shape-name-path-case-spec-feasibility.md](worksheet-control-shape-name-path-case-spec-feasibility.md)
- Shape / explicit sheet-name root 系: [shape-oleformat-object-promotion-feasibility.md](shape-oleformat-object-promotion-feasibility.md), [shape-oleformat-object-explicit-sheet-root-feasibility.md](shape-oleformat-object-explicit-sheet-root-feasibility.md), [explicit-sheet-name-broad-root-feasibility.md](explicit-sheet-name-broad-root-feasibility.md)
- Workbook binding / active workbook 系: [workbook-binding-manifest-feasibility.md](workbook-binding-manifest-feasibility.md), [active-workbook-identity-provider-contract.md](active-workbook-identity-provider-contract.md), [application-workbook-root-feasibility.md](application-workbook-root-feasibility.md)
- Workbook root family test 系: [workbook-root-family-case-table-policy.md](workbook-root-family-case-table-policy.md), [workbook-root-family-server-mirror-policy.md](workbook-root-family-server-mirror-policy.md), [workbook-root-family-server-mirror-cross-family-preconditions.md](workbook-root-family-server-mirror-cross-family-preconditions.md), [workbook-root-shadow-fixture-topology-feasibility.md](workbook-root-shadow-fixture-topology-feasibility.md), [workbook-root-shadow-fixture-split-poc.md](workbook-root-shadow-fixture-split-poc.md), [workbook-root-shadow-text-source-canonicalization-feasibility.md](workbook-root-shadow-text-source-canonicalization-feasibility.md), [shadow-fixture-split-cross-family-observation.md](shadow-fixture-split-cross-family-observation.md), [worksheet-broad-root-shadow-extension-matrix-feasibility.md](worksheet-broad-root-shadow-extension-matrix-feasibility.md), [worksheet-broad-root-shadow-server-helper-feasibility.md](worksheet-broad-root-shadow-server-helper-feasibility.md), [worksheet-broad-root-shadow-coverage-feasibility.md](worksheet-broad-root-shadow-coverage-feasibility.md), [worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md](worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md), [worksheet-broad-root-extension-only-interaction-shared-spec-feasibility.md](worksheet-broad-root-extension-only-interaction-shared-spec-feasibility.md), [application-workbook-root-extension-only-interaction-shared-spec-feasibility.md](application-workbook-root-extension-only-interaction-shared-spec-feasibility.md), [application-workbook-root-extension-only-completion-shared-spec-feasibility.md](application-workbook-root-extension-only-completion-shared-spec-feasibility.md), [application-workbook-root-completion-server-scope-feasibility.md](application-workbook-root-completion-server-scope-feasibility.md), [application-workbook-root-interaction-semantic-server-scope-feasibility.md](application-workbook-root-interaction-semantic-server-scope-feasibility.md)

## 読み方

- まずコア文書を 1 つ開き、必要になった機能別メモだけを追加で開く。
- 機能別メモは、その論点専用の制約整理であり、リポジトリ全体の運用ルールではない。
- 通常タスクの進め方と Done の定義は [../../AGENTS.md](../../AGENTS.md) を正本とし、機能別メモの整理や追加は実装を進めるための補助に留める。
- `docs/process/coderabbit-review-logs/` は記録専用で、通常の参照導線には入れない。

## 更新ルール

- 共通ルールを変えるときはコア文書を更新する。
- 特定機能の制約整理を変えるときは対応する機能別メモか ADR を更新する。
- 作業進捗は [`TASKS.md`](../../TASKS.md) に残す。
