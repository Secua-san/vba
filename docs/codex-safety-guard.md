# Codex Safety Guard

自動コミットや自動 PR を判断するときの軽い入口チェック。詳細ルールは `docs/process/` の正本を読む。

## 先に読む参照先
- 運用ドキュメントの入口: [process/README.md](process/README.md)
- docs 全体の入口: [README.md](README.md)
- PR 本文テンプレート: [../.github/pull_request_template.md](../.github/pull_request_template.md)

## 実装前チェック
- `PLAN.md` / `TASKS.md` と対象機能の要件書または ADR を確認する
- 対象ファイル、変更理由、影響範囲、最小変更案を先に出す
- `.codex/skills/minimal-change` と `.codex/skills/no-speculation` の停止条件に当たらないことを確認する
- 1 つの論理単位に閉じ、実装、文書、整理、生成物を不要に混ぜない

## テスト選択チェック
- 検証前に `.codex/skills/test-budget` の対応表で最小の関連テストを選ぶ
- `scripts/` は `node --test scripts/test/<file>.test.mjs` または `npm run test:scripts`
- `packages/core/` は `npm run test --workspace @vba/core`
- `packages/server/` は `npm run test --workspace @vba/server` または該当 `node --test packages/server/test/<file>.test.js`
- `packages/extension/` はまず `npm run build --workspace vba-extension`
- `npm run test --workspace vba-extension` と `npm run test:host` は重い E2E として扱い、明示指示時だけ実行する
- `npm test` / `npm run test` は全体テストとして扱い、明示指示時だけ実行する

## コミット / PR 前チェック
- コミット前は現在差分だけを簡易自己レビューし、関連テストが通った小単位だけを commit する
- PR 前は `reviewer` の自己レビューを行う。ユーザーが PR 前 full gate を明示した場合は `npm run lint`、`npm test`、`npm run test:host` を通す
- PR は draft にせず、CodeRabbit review の対象になる通常 PR とする

## 即停止する条件
- `main` / `master` 直作業、または detached HEAD のまま進めようとしている
- 仕様が曖昧、または差分が複数目的で混在している
- lint / build / test の失敗理由が説明できない
- 認証、権限、課金、シークレット、CI/CD、インフラ、DB、マイグレーション、破壊的変更、依存関係のメジャー更新を含む
- 現在の変更内容を 1 文で説明できない
