---
name: lightweight-review
description: Perform a short, implementation-led review of the current diff only. Use after implementation to catch obvious defects, edge cases, naming drift, or unnecessary complexity without turning review into a redesign task.
---

# Lightweight Review

## skill 名
`lightweight-review`

## 目的
実装差分だけを対象に、短く修正可能な粒度で不具合候補を洗い出す。レビューで実装を止めず、必要な指摘だけ返す。

## いつ使うか
- 実装完了後
- PR 前の簡易セルフレビューをしたいとき
- 差分が狭く、全面レビューではなく確認だけで足りるとき

## 使ってはいけないケース
- 実装前
- 方針の再整理や大規模リファクタ提案をしたいとき
- リポジトリ全体の設計見直しを始めたいとき
- 実装停止を前提に長文講評をしたくなっているとき

## レビュー観点
- バグの可能性
- 型や null/undefined 相当の未考慮
- 境界条件の抜け
- 明らかな命名問題
- 不要な複雑性

## 実行手順
1. 差分の対象範囲を確認する。
2. 明らかな不具合候補だけを抽出する。
3. 軽微な改善点があれば付記する。
4. ブロッカーかどうかを区別して返す。

## 出力ルール
- レビュー対象は実装差分のみに限定する
- 指摘は短く、修正可能な粒度で返す
- ブロッカーと軽微な改善を分ける
- ブロッカーがなければ実装継続を妨げない
- 指摘が無ければ、その旨と残留リスクの有無だけを返す

## 禁止事項
- 設計全体の見直し
- 大規模リファクタ提案
- 方針文書の整理
- 差し戻し前提の過剰レビュー
- 差分外の論点拡張

迷ったら全面レビューへ広げず、実装差分の明確な問題だけに絞る。

## 例
- null/undefined 相当の未考慮
- 境界値の抜け
- 命名の一貫性欠如
- ただし全面的再設計は提案しない
