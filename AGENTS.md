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
- タスク管理を `TASKS.md` で行い、進捗に合わせ随時更新する
- PR 作成前のサブエージェント自己レビューは `reviewer` を既定とする

## リポジトリ構成
- `packages/extension/`: VS Code 拡張本体
- `packages/server/`: Language Server
- `packages/core/`: 解析コア
- `resources/vbac/`: 同梱する `vbac.wsf`
- `docs/adr/`: ADR
- `.github/`: PR テンプレートと自動化設定

## 必要に応じて読むドキュメント
- docs 全体の入口と読み順: [docs/README.md](docs/README.md)
- プロダクト要件とマイルストーン: [docs/requirements/000-overview.md](docs/requirements/000-overview.md)
- ADR の入口: [docs/adr/README.md](docs/adr/README.md)
- 運用ドキュメントの入口: [docs/process/README.md](docs/process/README.md)
- PR テンプレート: [.github/pull_request_template.md](.github/pull_request_template.md)

## 運用メモ
- 実装前に、[docs/README.md](docs/README.md) から対象機能に対応する要件書または ADR を確認する
- コミットや PR を扱う前に、[docs/process/README.md](docs/process/README.md) から必要な運用ドキュメントだけを確認する
- `TASKS.md` を更新するときは、整理や方針メモ単独ではなく、実コード変更と検証を伴う作業だけを通常タスクの完了へ移す
- 外部 MCP サーバー呼び出しは共通の retry / rate-limit 層を必ず通し、`429` 検知、`Retry-After` 優先、未指定時の指数バックオフ + ジッター、呼び出し間隔制御、同一問い合わせの重複抑止、対象 MCP 名を含む retry / wait / 最終失敗理由ログを実装する
- 同じ論点を繰り返し検討して進まない場合は、`docs/process/sub-agent-escalation.md` に従ってサブエージェントへ切り替える
- `reviewer` が利用できない場合は、`C:\Users\tagi0\.codex\config.toml` と `C:\Users\tagi0\.codex\agents\reviewer.toml` を確認し、設定を直してから PR 作成へ進む
- CodeRabbit レビュー記録を残す場合は `docs/process/coderabbit-review-logs/YYYY-MM.md` に直接追記し、ログは参照用ではなく記録専用として扱う
- 自己レビューと CodeRabbit の重複指摘、および `required` / `optional` の判断は `docs/process/coderabbit-review.md` の正本ルールに従う
- ルール変更時は重複記載を増やさず、正本ドキュメントを更新する
