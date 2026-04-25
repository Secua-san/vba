# AGENTS.md

このリポジトリでは、Codex を使って Excel VBA 向け VS Code 拡張機能を開発する。  
固定前提は **Excel VBA のみ**、**Option Explicit 前提**、**Win64 / PtrSafe 必須**。

## 最小変更ガード
- 推測で実装しない。不明点、要件不足、影響範囲不明の箇所は実装せず、確認事項または TODO として明示する
- 変更はユーザー依頼を満たす最小差分に限定する
- 無関係な修正、命名変更だけの変更、フォーマット変更だけの変更を混ぜない
- 新規抽象化、共通化、リファクタリングは原則禁止する。必要に見える場合も、実装せず提案に留めて承認を取る
- 実装前に必ず計画を提示し、対象ファイル、変更理由、影響範囲を明示する
- ユーザーの承認なしにコード変更を行わない
- 1 回の作業は 1 つの論理単位に絞り、広がる場合は次タスクまたは別 PR に分離する

## テスト選択ルール
- 実行するテストは変更箇所に直接関係するものだけに限定する
- ルートの `npm test` や `npm run test` などの全体テストは禁止する。ユーザーが明示した場合だけ実行する
- VS Code extension host を起動する `npm run test:host` と `npm run test --workspace vba-extension` は E2E / 重いテストとして扱い、ユーザーが明示した場合だけ実行する
- 新規テストは変更内容を確認する最小ケースだけにする
- 「念のため」のテスト追加、既存ケースの網羅的な増量、無関係領域のスナップショット更新は禁止する
- テスト選択の目安:
  - `scripts/` 変更: `npm run test:scripts` または該当する `node --test scripts/test/<file>.test.mjs`
  - `packages/core/` 変更: `npm run test --workspace @vba/core` または該当する compiled test
  - `packages/server/` 変更: `npm run test --workspace @vba/server` または該当する `node --test packages/server/test/<file>.test.js`
  - `packages/extension/` 変更: まず build / 型チェックなど軽い確認を優先し、extension host は明示承認時のみ

## 出力ルール
- 実装前に、対象ファイル、変更理由、影響範囲を必ず提示する
- 実装後に、変更ファイル一覧、実行したテスト、未実行テストと理由、リスクを必ず報告する

## 作業優先順位
- 主目的はコード実装とし、ドキュメント整理・指針整理・命名整理・レビュー容易性向上は実装を前進させるための補助作業に留める
- 可読性・保守性・レビュー容易性を最優先する
- 変更は論理単位で分割する
- 無関係な変更を混在させない
- 不明点は推測実装せず、要約と確認を優先する
- 大規模変更は段階的な PR に分割する
- 優先度確認、作業状態確認、影響範囲調査、slice 切り分け、既存パターン探索、再現確認のような非実装作業は、原則として先に適切なサブエージェントへの委譲を検討する
- 親エージェントが 2〜3 ステップ以上連続で調査だけを続けそうな場合は、まずサブエージェントへ切り替える
- 探索や判断が空転したら、サブエージェントへ論点を絞って意見を求める

## 整理系 skill の扱い
- 整理系作業は親エージェントの自由行動ではなく、repo-local skill として扱う
- `.codex/skills/doc-minimum-update` と `.codex/skills/lightweight-review` の無条件自動実行を禁止する
- 整理系 skill を明示指示なしで使ってよいのは、実装後であるか同一タスク内で実装完了が確実で、対象が実装差分に直接関係し、変更が必要最小限で、実行しないと説明不足・使用方法不整合・手順不整合・レビュー困難のいずれかが出る場合に限る
- 整理系 skill を実行しても、主成果物は常にコード変更とし、skill の実行だけで完了してはならない
- 実行要否は親エージェントが判断してよく、迷う場合だけ `skill-gatekeeper` 役割へ確認する
- 判断に迷う場合は実行しない。整理より実装を優先する

## 作業ガード skill の扱い
- `.codex/skills/minimal-change`、`.codex/skills/no-speculation`、`.codex/skills/test-budget` は整理系 skill ではなく、作業ガード skill として扱う
- 作業ガード skill は、推測実装、変更範囲拡大、過剰テストを防ぐために実装前または検証前に使ってよい
- 作業ガード skill は承認、計画提示、対象ファイル・変更理由・影響範囲の明示を省略する理由にしてはならない
- 作業ガード skill は実装や整理の代替成果物ではなく、既存ルールを守るための確認手順として使う

## 親エージェントと subagent の責務
- 親エージェントが常に実装責任を持ち、タスクの主目的をコード実装として定義する
- 優先度確認、branch / main / 作業状態確認、影響範囲調査、slice 比較、既存パターン探索、再現確認、整理系 skill の実行可否判定、実装差分の簡易確認は、原則 subagent に担当させる
- 親エージェントは、自分で長く調査する前に、該当する subagent へ振ることをまず検討する
- subagent の成果物は短く受け取り、親エージェントが最終的な実装対象と実装方針を決める
- 親エージェントは「調査して終わり」にしてはならず、実装可能になった時点で必ずコード変更と検証へ進む
- subagent へ実装の代行、方針文書の主導、大規模な整理作業を委譲しない
- 調査結果や判定結果を受けたら、親エージェントが自らコード変更、検証、必要最小限の追記まで進める
- 役割の詳細と既定実体は [docs/process/sub-agent-escalation.md](docs/process/sub-agent-escalation.md) を正本とする

## 標準実行フロー
通常のプロダクトタスクは、原則として次の流れで進める。

1. 親エージェントが実装目的を短く定義する
2. `task-priority-auditor` に優先候補確認を依頼する
3. 必要なら `branch-state-checker` に着手前提を確認させる
4. 必要なら `slice-scout` に 1 PR slice 比較を依頼する
5. 必要なら `pattern-investigator` に既存依存を調査させる
6. 必要なら `repro-prober` に最小再現確認をさせる
7. 親エージェントが実装対象を決定する
8. 親エージェントがコード変更を行う
9. 親エージェントが検証を行う
10. 必要なら `skill-gatekeeper` に整理系 skill 実行要否を判定させる
11. 必要な場合だけ skill を実行する
12. 必要なら軽量レビューを行う
13. 親エージェントが完了判定する

補足:
- すべての subagent を毎回使う必要はない
- ただし、優先度確認、slice 比較、再現確認のような非実装作業を親エージェントが抱え込み続けない
- subagent の結果待ちだけで止まらず、親エージェントは並行して実装判断や非重複の準備を進める

## 実装優先ルール
- 通常のプロダクトタスクでは、原則として次の順で進める
  1. 実装目的を定義し、必要な非実装作業を適切な subagent へ切り出す
  2. 最小実装を行う
  3. 必要なテスト・型・lint・build 修正を行う
  4. 最後に実装差分に直接関係する最小限のドキュメントを更新する
- 通常のプロダクトタスクでは、1 PR に収まる論理単位が完了するまで、実装、検証、差分レビュー、必要な追加修正を繰り返し、途中の中間報告だけで止めない
- 実装途中のレビューで不備や回帰リスクが見つかった場合は、その場で修正と再検証を行ってから次へ進む
- 1 PR に収まらないと判断した時点でスコープを切り直し、残りは次の PR 単位へ分離する
- 実装前に整理メモや設計メモを作る場合でも、同じ作業単位の中でコード変更まで進める
- 整理案の提示だけで止めず、実装可能なものはそのまま実装する
- ドキュメントのみの変更で通常タスクを完了扱いにしない
- 例外は、ユーザーが明示的にドキュメントや運用ルールの更新のみを依頼した場合に限る

## Done の定義
- 通常のプロダクトタスクを完了扱いにするのは、少なくとも 1 つ以上の実コード変更がある場合に限る
- 通常のプロダクトタスクでは、実装差分に応じて、テスト、型チェック、lint、build、最小動作確認のうち適切な検証を必ず行う
- 通常のプロダクトタスクでは、ドキュメント更新は実装差分に直接付随する最小限に留める
- ドキュメント整理のみで終わった作業は、通常タスクの完了として扱わない
- ユーザーが明示的にドキュメントや運用ルールの更新のみを依頼した場合は、この Done 定義の例外として扱う

## 固定ルール
- 解析対象は `.bas` / `.cls` / `.frm` / `.frx` を主軸とする
- XLAM 連携は補助機能として `resources/vbac/vbac.wsf` を使う
- `combine` は破壊的操作なので、バックアップ・確認・検証・エラーログを必須とする
- 開発は Windows ネイティブ環境で行い、`npm` と Node LTS を前提にする
- build / test / package は Windows 直で通る状態を維持する
- 解析は手書きパーサで実装し、`lexer -> parser -> AST -> symbol resolution -> type inference -> diagnostics/completion` を基本パイプラインとする
- `Declare PtrSafe`、`LongPtr`、`#If VBA7 Then` などの条件付きコンパイルを優先して扱う
- 重要な設計判断は `docs/adr/` に記録する
- タスク管理は `PLAN.md`、`TASKS.md`、`TASKLOG.md` を分けて行い、`PLAN.md` にはフェーズ進捗とロードマップ、`TASKS.md` には直近の状況、次に行うタスク、重要事項だけを残し、詳細な完了履歴や長い補足は `TASKLOG.md` に集約する
- `README`、`docs/README.md`、`docs/process/README.md` などのよく読まれる入口文書には要点と導線だけを置き、長い履歴やログを直接持ち込まない
- PR 作成前のサブエージェント自己レビューは `diff-reviewer` 役割（既定実体は `reviewer`）を使う

## リポジトリ構成
- `packages/extension/`: VS Code 拡張本体
- `packages/server/`: Language Server
- `packages/core/`: 解析コア
- `resources/vbac/`: 同梱する `vbac.wsf`
- `docs/adr/`: ADR
- `.codex/skills/`: repo-local の整理系 / 完了判定 skill
- `.github/`: PR テンプレートと自動化設定

## 必要に応じて読むドキュメント
- フェーズ進捗と次の実装候補: [PLAN.md](PLAN.md)
- docs 全体の入口と読み順: [docs/README.md](docs/README.md)
- プロダクト要件とマイルストーン: [docs/requirements/000-overview.md](docs/requirements/000-overview.md)
- ADR の入口: [docs/adr/README.md](docs/adr/README.md)
- 運用ドキュメントの入口: [docs/process/README.md](docs/process/README.md)
- 整理系 skill / 完了判定 skill の入口: [.codex/skills/README.md](.codex/skills/README.md)
- PR テンプレート: [.github/pull_request_template.md](.github/pull_request_template.md)

## 運用メモ
- 新規タスクを開始するときは、まず [PLAN.md](PLAN.md) を確認して現在位置と優先ロードマップを把握し、その後に [`TASKS.md`](TASKS.md) で直近の主タスクと次タスクを確認する
- ユーザーが新規タスクの文脈で `next` とだけ指示した場合は、[PLAN.md](PLAN.md) と [`TASKS.md`](TASKS.md) を参照しつつ、原則 `task-priority-auditor` に未完了の優先候補を確認させたうえで、最も優先度が高い次タスクへ進む指示として扱う
- 実装前に、[docs/README.md](docs/README.md) から対象機能に対応する要件書または ADR を確認する
- コミットや PR を扱う前に、[docs/process/README.md](docs/process/README.md) から必要な運用ドキュメントだけを確認する
- 通常タスクまたはユーザーが明示した docs-only タスクを完了したら、`skills/auto-commit-pr/SKILL.md` を既定で参照し、停止条件や禁止領域に当たらない限り、対象差分だけを適切なブランチへ分離して commit / PR 作成まで進める
- `skills/auto-commit-pr/SKILL.md` を使うときも、PR 作成前に `diff-reviewer` 役割で自己レビューを完了し、PR 作成後は `docs/process/coderabbit-review.md` の正本ルールに従って review 完了と必要な再確認が終わるまで merge しない
- 通常のプロダクトタスクでは、CodeRabbit review が完了するまで待機し、`review in progress` 状態のまま merge しない
- 通常のプロダクトタスクでは、CodeRabbit review の完了確認と必要対応が終わったらそのまま merge 完了まで進め、途中で中断しない
- `skills/auto-commit-pr/SKILL.md` を使うときも、未関連差分は含めず、品質ゲート結果と停止理由を省略しない
- 整理系 skill は実装前に起動せず、必要最小限の差分にだけ使う
- 整理系 skill の実行要否に迷う場合だけ `skill-gatekeeper` を使い、迷ったまま整理へ進まない
- 実装着手前提の確認は、必要に応じて `branch-state-checker` に任せ、親エージェントが自分で branch / main / 未コミット差分の確認を抱え込み続けない
- 1 PR slice の比較や既存パターン探索は、必要に応じて `slice-scout` / `pattern-investigator` / `repro-prober` に振り、親エージェントは実装判断とコード変更へ集中する
- `PLAN.md` はフェーズ進捗と再開時ロードマップの入口として保ち、`TASKS.md` は直近参照用に短く保ち、詳細な完了履歴、docs-only の判断記録、長い補足は `TASKLOG.md` へ移す
- `TASKS.md` を更新するときは、整理や方針メモ単独ではなく、実コード変更と検証を伴う作業だけを通常タスクの完了へ移す
- 外部 MCP サーバー呼び出しは共通の retry / rate-limit 層を必ず通し、`429` 検知、`Retry-After` 優先、未指定時の指数バックオフ + ジッター、呼び出し間隔制御、同一問い合わせの重複抑止、対象 MCP 名を含む retry / wait / 最終失敗理由ログを実装する
- 同じ論点を繰り返し検討して進まない場合だけでなく、優先度確認、slice 比較、再現確認のような非実装作業が先行するときも、`docs/process/sub-agent-escalation.md` に従って適切なサブエージェントへ切り替える
- `reviewer` が利用できない場合は、`C:\Users\tagi0\.codex\config.toml` と `C:\Users\tagi0\.codex\agents\reviewer.toml` を確認し、設定を直してから PR 作成へ進む
- CodeRabbit レビュー記録を残す場合は `docs/process/coderabbit-review-logs/YYYY-MM.md` に直接追記し、ログは参照用ではなく記録専用として扱う
- 自己レビューと CodeRabbit の重複指摘、および `required` / `optional` の判断は `docs/process/coderabbit-review.md` の正本ルールに従う
- ルール変更時は重複記載を増やさず、正本ドキュメントを更新する
