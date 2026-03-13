# ドキュメントガイド

`docs/` の入口。まずこのページで作業導線を選び、必要な詳細文書だけを開く。

## 最短動線

| 作業 | 読む順 |
| --- | --- |
| 実装を始める | [要件書](requirements/000-overview.md) -> [ADR 一覧](adr/README.md) -> [運用ガイド](process/README.md) -> [TASKS](../TASKS.md) |
| コミット / PR / レビュー | [運用ガイド](process/README.md) -> [Git Workflow](process/git-workflow.md) -> [CodeRabbit Review](process/coderabbit-review.md) -> [PR テンプレート](../.github/pull_request_template.md) |
| Microsoft Learn 由来データを更新する | [運用ガイド](process/README.md) -> [Microsoft Learn 組み込み署名再生成メモ](process/mslearn-signature-regeneration.md) |
| 設計判断を確認する | [ADR 一覧](adr/README.md) -> 対象 ADR |
| ドキュメント自体を更新する | このページ -> [運用ガイド](process/README.md) -> 対象の正本文書 |

## ディレクトリの役割

- [requirements/](requirements/000-overview.md): プロダクト要件、対象範囲、マイルストーン
- [adr/](adr/README.md): 長く効く設計判断の正本
- [process/](process/README.md): コミット、PR、レビュー、再生成、運用ルール

## 読み方の原則

- 入口ページから 1 段ずつ辿り、最初から全部開かない。
- 要件、設計判断、運用ルールを同じ文書へ混ぜない。
- 履歴系の文書は「必要な月だけ」開く。CodeRabbit ログは月次分割を前提にする。

## 更新ルール

- 恒久ルールは正本文書へ集約し、別文書では要約とリンクだけを置く。
- 設計判断は `docs/adr/`、進捗は [`TASKS.md`](../TASKS.md)、PR 単位の学びは CodeRabbit 月次ログへ残す。
- 長くなり続ける履歴は月次やテーマ単位で分割し、入口ファイルは軽く保つ。
