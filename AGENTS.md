# AGENTS.md

このリポジトリでは、Codex を使って Excel VBA 向け VS Code 拡張機能を開発する。  
固定前提は **Excel VBA のみ**、**Option Explicit 前提**、**Win64 / PtrSafe 必須**。

## 作業優先順位
- 可読性・保守性・レビュー容易性を最優先する
- 変更は論理単位で分割する
- 無関係な変更を混在させない
- 不明点は推測実装せず、要約と確認を優先する
- 大規模変更は段階的な PR に分割する
- 探索や判断が空転したら、サブエージェントへ論点を絞って意見を求める

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

## リポジトリ構成
- `packages/extension/`: VS Code 拡張本体
- `packages/server/`: Language Server
- `packages/core/`: 解析コア
- `resources/vbac/`: 同梱する `vbac.wsf`
- `docs/adr/`: ADR
- `.github/`: PR テンプレートと自動化設定

## 必要に応じて読むドキュメント
- プロダクト要件とマイルストーン: [docs/requirements/000-overview.md](docs/requirements/000-overview.md)
- パーサ方針: [docs/adr/0001-parser-strategy.md](docs/adr/0001-parser-strategy.md)
- vbac 安全方針: [docs/adr/0002-vbac-integration-safety.md](docs/adr/0002-vbac-integration-safety.md)
- ブランチ・コミット・PR ルール: [docs/process/git-workflow.md](docs/process/git-workflow.md)
- 自動コミット / 自動 PR と品質ゲート: [docs/process/automation-policy.md](docs/process/automation-policy.md)
- CodeRabbit 対応: [docs/process/coderabbit-review.md](docs/process/coderabbit-review.md)
- サブエージェントへのエスカレーション: [docs/process/sub-agent-escalation.md](docs/process/sub-agent-escalation.md)
- コミットメッセージ規約: [docs/commit-message.md](docs/commit-message.md)
- PR テンプレート: [.github/pull_request_template.md](.github/pull_request_template.md)

## 運用メモ
- 実装前に、対象機能に対応する要件書または ADR を確認する
- コミットや PR を扱う前に、`docs/process/` 配下の運用ドキュメントを確認する
- 外部 MCP サーバー呼び出しは共通の retry / rate-limit 層を必ず通し、`429` 検知、`Retry-After` 優先、未指定時の指数バックオフ + ジッター、呼び出し間隔制御、同一問い合わせの重複抑止、対象 MCP 名を含む retry / wait / 最終失敗理由ログを実装する
- 同じ論点を繰り返し検討して進まない場合は、`docs/process/sub-agent-escalation.md` に従ってサブエージェントへ切り替える
- ルール変更時は重複記載を増やさず、正本ドキュメントを更新する
