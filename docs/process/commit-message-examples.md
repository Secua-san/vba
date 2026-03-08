# Commit Message Examples

## 良い例
- `feat(parser): add property procedure parsing`
- `fix(lsp): avoid completion crash on invalid tokens`
- `refactor(symbols): split local and module scopes`
- `test(parser): add PtrSafe declaration cases`
- `docs(agents): clarify auto pr safety rules`
- `chore(deps): update vscode-languageserver`

## 避ける例
- `fix: bug fix`
- `docs: update`
- `chore: misc`
- `refactor: improve parser`
- `style(core): tidy up`

## 本文が必要なケース
- 変更理由が差分だけでは分からない
- トレードオフや制約を残す必要がある
- 既知の限界や後続タスクを明示したい

## 本文例

```text
feat(inference): support Variant numeric promotion

Align numeric promotion with common VBA behavior so diagnostics and
completion stay closer to runtime expectations.

Keep the change limited to arithmetic expressions for now.
```

## BREAKING CHANGE 例

```text
feat(parser)!: remove legacy error recovery path

BREAKING CHANGE: parser no longer accepts missing End If recovery.
```
