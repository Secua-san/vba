# ドキュメントガイド

`docs/` の入口。普段はこのページ、[`TASKS.md`](../TASKS.md)、今回の作業に直接関係する正本文書だけを開く。過去の履歴や長い補足が必要なときだけ [`TASKLOG.md`](../TASKLOG.md) を追加で開く。

## まず開くもの

| 状況 | 開く文書 |
| --- | --- |
| 今回の作業範囲を確認する | [`TASKS.md`](../TASKS.md) |
| 過去の完了履歴や docs-only 判断を遡る | [`TASKLOG.md`](../TASKLOG.md) |
| 機能範囲や優先度が曖昧 | [要件書](requirements/000-overview.md) |
| 既存設計の制約を確認したい | [ADR 一覧](adr/README.md) -> 対象 ADR |
| コミット / PR / レビュー / 再生成 | [運用ガイド](process/README.md) -> 対象の運用文書 |
| ドキュメントを更新する | このページ -> 対象の正本文書 |

## ディレクトリの役割

- [requirements/](requirements/000-overview.md): プロダクト要件、対象範囲、マイルストーン
- [adr/](adr/README.md): 複数パッケージにまたがる長期の設計判断
- [process/](process/README.md): ブランチ、レビュー、再生成、運用手順と機能別メモ

## 読み方の原則

- 入口から 1 段で止め、必要な正本文書だけを開く。
- `requirements`、`adr`、`process` を同時に全部読まない。今必要な系統だけ選ぶ。
- 同じルールは 1 つの正本文書にだけ置き、入口ページは導線だけに保つ。
- `TASKS.md` は直近の状況と次タスクの参照用、`TASKLOG.md` は履歴参照用として分ける。
- `docs/process/coderabbit-review-logs/` は記録専用で、通常の参照導線に含めない。

## 更新ルール

- 恒久ルールは正本文書へ集約する。
- 設計判断は `docs/adr/`、日常参照する進捗要点は [`TASKS.md`](../TASKS.md)、詳細な完了履歴や判断ログは [`TASKLOG.md`](../TASKLOG.md) に分ける。
- 入口ページに履歴、テンプレート、長い補足を持ち込まない。
