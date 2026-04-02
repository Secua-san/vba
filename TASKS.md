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
- `assignment` / `call` の first slice は完了済みで、次は block statement 群を node kind 化する段階
- 過去の完了履歴や docs-only 更新の経緯は [`TASKLOG.md`](TASKLOG.md) を参照する

## 進行中

- [ ] ProcedureStatementNode の block statement structured AST を広げる
  - `If` / `Select Case` / `For` / `For Each` / `Do` / `While` / `With` / `On Error` を text 判定ではなく node kind で持てるようにする
  - block validation の stack 判定を raw text regex から AST kind ベースへ移し、`range` / `text` は formatter / diagnostics / navigation 互換のため維持する
  - 既存 diagnostics、references / rename、semantic token、formatter の回帰を崩さないことを完了条件にする

## 次に行うタスク

- `If` / `Select Case` / `For` / `For Each` の structured node 追加を先に入れる
- `Do` / `While` / `With` / `On Error` を続けて node kind 化する
- block validation を AST kind ベースへ移したうえで core / server 回帰を確認する

## 直近の更新

- [運用] `TASKS.md` を参照用サマリ、[`TASKLOG.md`](TASKLOG.md) を履歴ログに分離した
- [完了] extension host の active workbook identity snapshot と shared semantic test の安定化を完了した
- [完了] `ProcedureStatementNode` の `assignment` / `call` structured AST first slice を導入した
