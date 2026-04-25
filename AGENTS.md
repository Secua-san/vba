# AGENTS.md

このリポジトリでは、Codex を使って Excel VBA 向け VS Code 拡張機能を開発する。

固定前提:
- Excel VBA のみ
- `Option Explicit` 前提
- Win64 / `PtrSafe` 必須

## 基本ルール
- 推測で実装しない。不明点、要件不足、影響範囲不明の箇所は確認事項または TODO にする
- 変更はユーザー依頼を満たす最小差分に限定する
- 無関係な修正、命名変更だけの変更、フォーマット変更だけの変更を混ぜない
- 新規抽象化、共通化、リファクタリングは原則禁止する。必要な場合は実装せず提案に留める
- 実装前に、対象ファイル、変更理由、影響範囲、最小変更案を提示する
- ユーザーの承認なしにコード変更しない
- 1 回の作業は 1 つの論理単位に絞る。広がる場合は別タスクまたは別 PR に分離する

## 作業ガード skill
- `.codex/skills/minimal-change`、`.codex/skills/no-speculation`、`.codex/skills/test-budget` は作業ガード skill として扱う
- 実装前または検証前に使ってよい
- 承認、計画提示、対象ファイル・変更理由・影響範囲の明示を省略する理由にしない
- 実装や整理の代替成果物にしない

## サブエージェント
- 親エージェントは実装責任を持ち、最終判断とコード変更を担当する
- 非実装作業は積極的に subagent へ切り出す
- 1 つの subagent に複数の責務を抱えさせない。役割ごとに小さく依頼する
- 調査が 2〜3 ステップ続く見込みなら、先に適切な subagent へ渡す
- subagent の結果待ちだけで止まらず、親エージェントは非重複の準備を進める
- subagent に実装代行、方針文書の主導、大規模整理を任せない
- 役割の正本は [docs/process/sub-agent-escalation.md](docs/process/sub-agent-escalation.md) とする

主な切り出し先:
- 優先度確認: `task-priority-auditor`
- branch / main / 未コミット差分確認: `branch-state-checker`
- 1 PR slice 比較: `slice-scout`
- 既存パターン探索: `pattern-investigator`
- 最小再現確認: `repro-prober`
- 整理系 skill 実行可否判定: `skill-gatekeeper`
- PR 前自己レビュー: `reviewer`（diff-reviewer 役割）

## 実装フロー
1. 目的を短く定義する
2. 必要な非実装作業を subagent へ切り出す
3. 実装対象と最小変更案を決める
4. 承認後に最小実装する
5. 関連する検証だけを行う
6. 実装差分に直接必要な最小文書だけ更新する
7. 差分を確認し、必要なら軽量レビューを行う
8. 完了条件を確認する

## テストルール
- 変更箇所に直接関係するテストだけ実行する
- ルートの `npm test` / `npm run test` は禁止する。明示指示がある場合だけ実行する
- `npm run test:host` と `npm run test --workspace vba-extension` は E2E / 重いテストとして扱い、明示指示がある場合だけ実行する
- 新規テストは変更内容を確認する最小ケースだけにする
- 「念のため」のテスト追加、無関係領域のスナップショット更新は禁止する

テスト選択:
- `scripts/`: `npm run test:scripts` または `node --test scripts/test/<file>.test.mjs`
- `packages/core/`: `npm run test --workspace @vba/core`
- `packages/server/`: `npm run test --workspace @vba/server` または `node --test packages/server/test/<file>.test.js`
- `packages/extension/`: まず `npm run build --workspace vba-extension` を優先し、extension host は明示承認時のみ
- docs / ルールのみ: `git diff --check -- <changed-files>`

## 整理系 skill
- `.codex/skills/doc-minimum-update` と `.codex/skills/lightweight-review` は無条件実行しない
- 実装後、または同一タスク内で実装完了が確実な場合だけ使う
- 対象は実装差分に直接関係する最小範囲に限定する
- 判断に迷う場合は実行しない。必要なら `skill-gatekeeper` に確認する
- skill 実行だけで完了扱いにしない

## Done の定義
- 通常タスクは、実コード変更、関連検証、必要最小限の文書更新が揃ったときだけ完了にする
- ドキュメント整理のみでは通常タスクの完了にしない
- ユーザーが明示的に docs-only / ルール整備を依頼した場合は例外とする
- 完了時は、変更ファイル、実行テスト、未実行テストと理由、リスクを報告する

## 固定ルール
- 解析対象は `.bas` / `.cls` / `.frm` / `.frx` を主軸とする
- XLAM 連携は `resources/vbac/vbac.wsf` を補助機能として使う
- `combine` は破壊的操作。バックアップ、確認、検証、エラーログを必須にする
- 開発は Windows ネイティブ環境で行い、`npm` と Node LTS を前提にする
- build / test / package は Windows 直で通る状態を維持する
- 解析は手書きパーサで実装し、`lexer -> parser -> AST -> symbol resolution -> type inference -> diagnostics/completion` を基本パイプラインとする
- `Declare PtrSafe`、`LongPtr`、`#If VBA7 Then` などの条件付きコンパイルを優先する
- 重要な設計判断は `docs/adr/` に記録する
- タスク管理は `PLAN.md`、`TASKS.md`、`TASKLOG.md` に分ける
- `TASKS.md` の完了扱いは、実コード変更と検証を伴う通常タスクだけにする
- よく読まれる入口文書には要点と導線だけを置く
- 外部 MCP 呼び出しは共通の retry / rate-limit 層を通し、`429`、`Retry-After`、backoff、失敗ログを扱う

## リポジトリ構成
- `packages/extension/`: VS Code 拡張本体
- `packages/server/`: Language Server
- `packages/core/`: 解析コア
- `resources/vbac/`: 同梱する `vbac.wsf`
- `docs/adr/`: ADR
- `.codex/skills/`: repo-local skill
- `.github/`: PR テンプレートと自動化設定

## 参照ドキュメント
- フェーズ進捗とロードマップ: [PLAN.md](PLAN.md)
- 直近タスク: [TASKS.md](TASKS.md)
- 完了履歴と長い補足: [TASKLOG.md](TASKLOG.md)
- docs 入口: [docs/README.md](docs/README.md)
- 要件: [docs/requirements/000-overview.md](docs/requirements/000-overview.md)
- ADR 入口: [docs/adr/README.md](docs/adr/README.md)
- 運用入口: [docs/process/README.md](docs/process/README.md)
- skill 入口: [.codex/skills/README.md](.codex/skills/README.md)
- PR テンプレート: [.github/pull_request_template.md](.github/pull_request_template.md)

## 運用メモ
- 新規タスク開始時は `PLAN.md` と `TASKS.md` を確認する
- `next` のみ指示された場合は、`task-priority-auditor` で優先候補を確認してから進める
- 実装前に、対象機能に対応する要件書または ADR を確認する
- コミットや PR 前は、必要な運用ドキュメントだけ確認する
- 通常タスクまたは明示された docs-only タスク完了後は、`skills/auto-commit-pr/SKILL.md` を参照し、停止条件に当たらなければ対象差分だけ commit / PR まで進める
- PR 前に `reviewer` で自己レビューする
- CodeRabbit review は [docs/process/coderabbit-review.md](docs/process/coderabbit-review.md) に従う
- CodeRabbit review 完了前に merge しない
- CodeRabbit review 記録は `docs/process/coderabbit-review-logs/YYYY-MM.md` に残す
- `reviewer` が使えない場合は `C:\Users\tagi0\.codex\config.toml` と `C:\Users\tagi0\.codex\agents\reviewer.toml` を確認する
- ルール変更時は重複記載を増やさず、正本ドキュメントを更新する
