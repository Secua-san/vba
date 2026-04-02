# AGENTS.md

このリポジトリでは、Codex を使って Excel VBA 向け VS Code 拡張機能を開発する。  
固定前提は **Excel VBA のみ**、**Option Explicit 前提**、**Win64 / PtrSafe 必須**。

## 作業優先順位
- 主目的はコード実装とし、ドキュメント整理・指針整理・命名整理・レビュー容易性向上は実装を前進させるための補助作業に留める
- 可読性・保守性・レビュー容易性を最優先する
- 変更は論理単位で分割する
- 無関係な変更を混在させない
- 不明点は推測実装せず、要約と確認を優先する
- 大規模変更は段階的な PR に分割する
- 探索や判断が空転したら、サブエージェントへ論点を絞って意見を求める

## 整理系 skill の扱い
- 整理系作業は親エージェントの自由行動ではなく、repo-local skill として扱う
- `.codex/skills/doc-minimum-update` と `.codex/skills/lightweight-review` の無条件自動実行を禁止する
- 整理系 skill を明示指示なしで使ってよいのは、実装後であるか同一タスク内で実装完了が確実で、対象が実装差分に直接関係し、変更が必要最小限で、実行しないと説明不足・使用方法不整合・手順不整合・レビュー困難のいずれかが出る場合に限る
- 整理系 skill を実行しても、主成果物は常にコード変更とし、skill の実行だけで完了してはならない
- 実行要否は親エージェントが判断してよく、迷う場合だけ `skill-gatekeeper` 役割へ確認する
- 判断に迷う場合は実行しない。整理より実装を優先する

## 親エージェントと subagent の責務
- 親エージェントが常に実装責任を持ち、タスクの主目的をコード実装として定義する
- subagent は調査、整理系 skill の実行可否判定、実装差分の簡易確認にだけ使う
- subagent へ実装の代行、方針文書の主導、大規模な整理作業を委譲しない
- 調査結果や判定結果を受けたら、親エージェントが自らコード変更、検証、必要最小限の追記まで進める
- 役割の詳細と既定実体は [docs/process/sub-agent-escalation.md](docs/process/sub-agent-escalation.md) を正本とする

## 実装優先ルール
- 通常のプロダクトタスクでは、原則として次の順で進める
  1. 変更対象コードと影響範囲を特定する
  2. 最小実装を行う
  3. 必要なテスト・型・lint・build 修正を行う
  4. 最後に実装差分に直接関係する最小限のドキュメントを更新する
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
- タスク管理は `TASKS.md` と `TASKLOG.md` を分けて行い、`TASKS.md` には直近の状況、次に行うタスク、重要事項だけを残し、詳細な完了履歴や長い補足は `TASKLOG.md` に集約する
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
- docs 全体の入口と読み順: [docs/README.md](docs/README.md)
- プロダクト要件とマイルストーン: [docs/requirements/000-overview.md](docs/requirements/000-overview.md)
- ADR の入口: [docs/adr/README.md](docs/adr/README.md)
- 運用ドキュメントの入口: [docs/process/README.md](docs/process/README.md)
- 整理系 skill / 完了判定 skill の入口: [.codex/skills/README.md](.codex/skills/README.md)
- PR テンプレート: [.github/pull_request_template.md](.github/pull_request_template.md)

## 運用メモ
- 実装前に、[docs/README.md](docs/README.md) から対象機能に対応する要件書または ADR を確認する
- コミットや PR を扱う前に、[docs/process/README.md](docs/process/README.md) から必要な運用ドキュメントだけを確認する
- 通常タスクまたはユーザーが明示した docs-only タスクを完了したら、`skills/auto-commit-pr/SKILL.md` を既定で参照し、停止条件や禁止領域に当たらない限り、対象差分だけを適切なブランチへ分離して commit / PR 作成まで進める
- `skills/auto-commit-pr/SKILL.md` を使うときも、PR 作成前に `diff-reviewer` 役割で自己レビューを完了し、PR 作成後は `docs/process/coderabbit-review.md` の正本ルールに従って review 完了と必要な再確認が終わるまで merge しない
- `skills/auto-commit-pr/SKILL.md` を使うときも、未関連差分は含めず、品質ゲート結果と停止理由を省略しない
- 整理系 skill は実装前に起動せず、必要最小限の差分にだけ使う
- 整理系 skill の実行要否に迷う場合だけ `skill-gatekeeper` を使い、迷ったまま整理へ進まない
- `TASKS.md` は直近参照用に短く保ち、詳細な完了履歴、docs-only の判断記録、長い補足は `TASKLOG.md` へ移す
- `TASKS.md` を更新するときは、整理や方針メモ単独ではなく、実コード変更と検証を伴う作業だけを通常タスクの完了へ移す
- 外部 MCP サーバー呼び出しは共通の retry / rate-limit 層を必ず通し、`429` 検知、`Retry-After` 優先、未指定時の指数バックオフ + ジッター、呼び出し間隔制御、同一問い合わせの重複抑止、対象 MCP 名を含む retry / wait / 最終失敗理由ログを実装する
- 同じ論点を繰り返し検討して進まない場合は、`docs/process/sub-agent-escalation.md` に従ってサブエージェントへ切り替える
- `reviewer` が利用できない場合は、`C:\Users\tagi0\.codex\config.toml` と `C:\Users\tagi0\.codex\agents\reviewer.toml` を確認し、設定を直してから PR 作成へ進む
- CodeRabbit レビュー記録を残す場合は `docs/process/coderabbit-review-logs/YYYY-MM.md` に直接追記し、ログは参照用ではなく記録専用として扱う
- 自己レビューと CodeRabbit の重複指摘、および `required` / `optional` の判断は `docs/process/coderabbit-review.md` の正本ルールに従う
- ルール変更時は重複記載を増やさず、正本ドキュメントを更新する
