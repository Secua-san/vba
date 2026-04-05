このディレクトリを、このリポジトリの repo-local skill 置き場として扱う。

- [doc-minimum-update](doc-minimum-update/SKILL.md): 実装差分に直接必要な最小文書更新だけを行う
- [lightweight-review](lightweight-review/SKILL.md): 実装差分だけを短く確認する
- [task-close-check](task-close-check/SKILL.md): 通常タスクを完了扱いにしてよいかを最終確認する

共通ルール:
- 実装前に整理系 skill を起動しない
- 優先度確認、branch 状態確認、slice 比較、既存パターン探索、再現確認のために skill を使わず、まず対応する subagent を使う
- `doc-minimum-update` と `lightweight-review` は実装後の補助に限定し、必要性が曖昧なら `skill-gatekeeper` に判定させる
- 迷ったら起動しない
- 主成果物は常にコード変更とし、skill や subagent の実行だけで完了しない

subagent の役割対応:
- `task-priority-auditor` = `C:\Users\tagi0\.codex\agents\task-priority-auditor.toml`
- `branch-state-checker` = `C:\Users\tagi0\.codex\agents\branch-state-checker.toml`
- `slice-scout` = `C:\Users\tagi0\.codex\agents\slice-scout.toml`
- `pattern-investigator` = `C:\Users\tagi0\.codex\agents\pattern-investigator.toml`
- `repro-prober` = `C:\Users\tagi0\.codex\agents\repro-prober.toml`
- `skill-gatekeeper` = `C:\Users\tagi0\.codex\agents\skill-gatekeeper.toml`
- `diff-reviewer` = `C:\Users\tagi0\.codex\agents\reviewer.toml`

詳細な使い分けと標準フローは [../../docs/process/sub-agent-escalation.md](../../docs/process/sub-agent-escalation.md) を正本とする。
