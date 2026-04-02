---
name: task-close-check
description: Check whether a normal implementation task is actually ready to be marked complete. Require real code changes, appropriate validation, and only minimal documentation updates tied to the implementation diff.
---

# Task Close Check

## skill 名
`task-close-check`

## 目的
通常タスクを完了扱いにしてよいかを、実装主導の観点で最終確認する。

## いつ使うか
- 実装、検証、必要最小限の追記まで終えたあと
- `TASKS.md` を更新する前
- PR 作成前に完了条件を明文化したいとき

## 完了条件
- 実コード変更がある
- テスト、型、lint、build、最小動作確認など、差分に応じた検証が行われている
- 文書更新がある場合は実装差分に直接関係する最小限に限定されている
- skill や subagent の成果だけで終わっていない
- README だけ更新、文書だけ更新、整理だけでは完了にしない

## 実行手順
1. コード変更の有無を確認する。
2. テスト、型、lint、ビルド、最小動作確認など適切な検証の有無を確認する。
3. 文書更新があれば、その範囲が最小か確認する。
4. 完了 / 未完了 を理由付きで返す。

## 出力ルール
- `完了可` または `未完了` を先に返す
- 理由はコード変更、検証、文書範囲の 3 点で短く示す
- 未完了の場合は、足りない要素だけを列挙する
- コード変更なしでは完了にしてはならない
- 文書のみ更新では完了にしてはならない

## 禁止事項
- skill や subagent の実行実績だけで完了扱いにする
- 通常タスクで docs-only を自己判断で例外化する
- 検証ゼロのまま完了扱いにする
- 文書更新の広がりを見逃す

迷ったら未完了と返し、実装と検証を優先する。

## 例
- README だけ更新されている場合は未完了
- コード変更はあるが検証ゼロなら未完了
- コード変更 + 必要検証 + 最小文書更新なら完了可
