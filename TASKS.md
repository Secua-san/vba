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

- Phase 3 の AST 安定化・構文情報整備は完了済み
- Phase 4 のシンボルテーブル・スコープ解析は完了済み
- Phase 5 の名前解決・基本型推論は完了済み
- Phase 6 の Diagnostics は完了済み
- `assignment` / `call`、member call、主要 block statement、label target statement、termination statement の structured AST slice は core / server 回帰で固定済み
- Codex 作業では [AGENTS.md](AGENTS.md) の「最小変更ガード」「テスト選択ルール」「出力ルール」を優先し、承認なしのコード変更、全体テスト、E2E テスト、無関係修正を避ける
- 過去の完了履歴や docs-only 更新の経緯は [`TASKLOG.md`](TASKLOG.md) を参照する

## 進行中

- [x] Phase 3 の AST 安定化・構文情報整備を進める
  - Phase 2 structured AST coverage は維持し、未構造化 `executableStatement` 互換 fallback を局所的に削った
  - downstream の type inference / references / semantic token を AST segment 優先へ寄せた
  - 既存 diagnostics、references / rename、semantic token、formatter の回帰を崩さないことを確認した

- [x] Phase 6 の Diagnostics を完了する
  - diagnostics の text fallback を、実装済み AST segment から外れた未構造化 `executableStatement` に限定した
  - structured leading label を AST に保持し、到達不能判定の label reset を text 走査から切り離した

- [ ] Codex 作業制御を強化する
  - テスト高速化の候補を調査する
  - 重いテストを分類し、明示承認が必要なテストを分ける
  - 最小テスト選択ルールを作成する
  - Codex 作業チェックリストを作成する

## 次に行うタスク

- 補完の文脈絞り込みと誤候補抑制を強化する
- core / server / extension の回帰を維持しながら structured AST と symbol resolution 利用箇所を保守する
- Codex 作業制御の改善タスクは、アプリ本体コードに触れず、文書とルール整備だけで小さく分割して進める

## 直近の更新

- [運用] `TASKS.md` を参照用サマリ、[`TASKLOG.md`](TASKLOG.md) を履歴ログに分離した
- [完了] extension host の active workbook identity snapshot と shared semantic test の安定化を完了した
- [完了] `ProcedureStatementNode` の `assignment` / `call` structured AST first slice を導入した
- [完了] `unreachable-code` diagnostics の block boundary 判定を structured statement metadata へ寄せた
- [完了] `Exit` / `End` termination statement を structured AST 化し、`unreachable-code` diagnostics の判定へ接続した
- [完了] Phase 2 の structured AST coverage を完了扱いにし、formatter compressed block 判定と local rename target range を structured kind / segment 優先へ寄せた
- [完了] Phase 3 の AST 安定化として、type inference の assignment fallback を structured assignment に限定し、member call を structured call statement として references / semantic token 経路へ接続した
- [完了] Phase 4 のシンボルテーブル・スコープ解析として、procedure symbol が同 kind の module symbol を shadow する解決を core / server 回帰で固定した
- [完了] Phase 5 の名前解決・基本型推論として、`CreateObject("WScript.Shell")` 既知 ProgID 解決を core / server 回帰で固定した
- [完了] Phase 6 の Diagnostics として、structured label metadata を parser / diagnostics に接続し、text fallback を未構造化 `executableStatement` に限定した
