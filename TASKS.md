# TASKS

`TASKS.md` は日常参照用のサマリ。直近の状況、次に行うタスク、重要事項だけを残す。詳細な完了履歴や長い補足は [`TASKLOG.md`](TASKLOG.md) を参照する。

## 運用ルール

- 通常タスクの進め方と Done の定義は [AGENTS.md](AGENTS.md) の「実装優先ルール」と「Done の定義」を正本とする
- `TASKS.md` には直近の状況、次に行うタスク、重要事項だけを書く
- 詳細な完了履歴、docs-only の判断記録、長い補足は [`TASKLOG.md`](TASKLOG.md) に移す
- 進行中タスクを `完了` 扱いにするのは、実コード変更と検証を伴ったときに限る
- `整理`、`指針整理`、`レビュー容易化` は副次効果であり、単独では通常タスクの主成果物にしない
- 文書更新は実装差分に直接付随する最小限に留め、よく読まれる入口文書へ長い履歴を持ち込まない

## 重要事項

- 現在の主タスクは `ProcedureStatementNode` の block statement structured AST 拡張
- `assignment` / `call`、主要 block statement、label target statement、termination statement の structured AST slice は完了済みで、次は downstream の raw text 依存を減らす段階
- Codex 作業では [AGENTS.md](AGENTS.md) の「最小変更ガード」「テスト選択ルール」「出力ルール」を優先し、承認なしのコード変更、全体テスト、E2E テスト、無関係修正を避ける
- 過去の完了履歴や docs-only 更新の経緯は [`TASKLOG.md`](TASKLOG.md) を参照する

## 進行中

- [ ] ProcedureStatementNode の block statement structured AST を広げる
  - 主要 block statement は node kind 化済みなので、`range` / `text` 互換を維持したまま downstream の diagnostics / references / formatter を AST kind ベースへ寄せる
  - block validation の first slice は AST kind ベースへ移行済みなので、残る raw text 依存を局所的に削る
  - 既存 diagnostics、references / rename、semantic token、formatter の回帰を崩さないことを完了条件にする

- [ ] Codex 作業制御を強化する
  - テスト高速化の候補を調査する
  - 重いテストを分類し、明示承認が必要なテストを分ける
  - 最小テスト選択ルールを作成する
  - Codex 作業チェックリストを作成する

## 次に行うタスク

- structured statement を使う diagnostics / symbol 連携の次 slice を進める
- `formatModuleIndentation.ts`、references / rename / semantic token に残る block text 判定を段階的に AST kind ベースへ寄せる
- core / server / extension の回帰を維持しながら structured AST 利用箇所を増やす
- Codex 作業制御の改善タスクは、アプリ本体コードに触れず、文書とルール整備だけで小さく分割して進める

## 直近の更新

- [運用] `TASKS.md` を参照用サマリ、[`TASKLOG.md`](TASKLOG.md) を履歴ログに分離した
- [完了] extension host の active workbook identity snapshot と shared semantic test の安定化を完了した
- [完了] `ProcedureStatementNode` の `assignment` / `call` structured AST first slice を導入した
- [完了] `unreachable-code` diagnostics の block boundary 判定を structured statement metadata へ寄せた
- [完了] `Exit` / `End` termination statement を structured AST 化し、`unreachable-code` diagnostics の判定へ接続した
