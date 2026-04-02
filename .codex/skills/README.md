このディレクトリを、このリポジトリの repo-local skill 置き場として扱う。

- [doc-minimum-update](doc-minimum-update/SKILL.md): 実装差分に直接必要な最小文書更新だけを行う
- [lightweight-review](lightweight-review/SKILL.md): 実装差分だけを短く確認する
- [task-close-check](task-close-check/SKILL.md): 通常タスクを完了扱いにしてよいかを最終確認する

共通ルール:
- 実装前に整理系 skill を起動しない
- 迷ったら起動しない
- 主成果物は常にコード変更とし、skill や subagent の実行だけで完了しない

subagent の役割対応:
- `impact-investigator` = `explorer`
- `skill-gatekeeper` = `default`
- `diff-reviewer` = `reviewer`
