# コミットメッセージ規約

本リポジトリでは Conventional Commits を使う。

## 基本形式

```text
type(scope): summary
```

## ルール
- `type` は `feat` / `fix` / `refactor` / `test` / `docs` / `chore` / `perf` / `style` を使う
- `scope` は対象領域を短く示す。不要なら省略可
- `summary` は **英語**・**命令形**・簡潔を基本とし、実際の差分内容を正確に表す
- `update`、`fix bug`、`misc` のような曖昧な表現は避ける
- 本文は必要な場合のみ追加し、`Why`、トレードオフ、制約、既知の限界を書く
- 破壊的変更は `!` または `BREAKING CHANGE:` で明示する
- 1 コミット = 1 意図を守る

## よく使う scope
- `lexer`
- `parser`
- `ast`
- `symbols`
- `inference`
- `diagnostics`
- `lsp`
- `extension`
- `commands`
- `xlam`
- `vbac`
- `ui`
- `config`
- `build`
- `deps`
- `docs`

詳細な例は [process/commit-message-examples.md](process/commit-message-examples.md) を参照する。
