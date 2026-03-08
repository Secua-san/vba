# AGENTS.md

このリポジトリでは、VS Code 上の Codex（OpenAI coding agent）を用いて VS Code 拡張機能を開発する。  
対象は **Excel VBA** のみ。**Option Explicit 前提**、**Win64 / PtrSafe 必須**。

---

# 0. Codex 作業原則（最重要）

## 0.1 基本方針
- 可読性・保守性・レビュー容易性を最優先
- 変更は必ず「論理的単位」で分割する
- 無関係な変更を混在させない
- 不明点は推測実装せず要約と確認を優先する
- 大規模変更は段階的PRに分割する

## 0.2 リポジトリ作成規則
以下の場合、Codexはリポジトリ作成を提案または実行してよい：

- プロジェクトが未初期化
- 大規模機能を独立開発する必要がある
- モノレポ分割が必要になった場合
- 実験的機能を隔離する必要がある場合

### 作成時の規則
- GitHub 上に新規リポジトリを作成
- 初期テンプレート構成を自動生成
- README / LICENSE / .gitignore を生成
- main ブランチを保護対象とする

---

# 1. ゴール（What we build）
VBA（Excel）向けに以下を提供する VS Code 拡張機能を作る。

- IntelliSense（補完）
- 高度な型推論（段階的に精度を上げる）
- 構文チェック（Diagnostics）
- シンタックスハイライト（TextMate / 必要に応じて Semantic）

---

# 2. ソース取り扱い方針（Source strategy）

## 2.1 解析対象（主軸）
- **.bas / .cls / .frm** を直接解析する（これが主軸）

## 2.2 XLAM（バイナリ）取り扱い（補助機能）
- 拡張機能に **vbac.wsf** を同梱し、以下を提供する：
  - **XLAM からモジュール抽出（decombine）**
  - **ソースから XLAM へ結合（combine）**
- combine は破壊的操作のため、拡張機能側で **安全ガード（バックアップ・確認・検証）**を必須とする

---

# 3. リポジトリ構成（Monorepo）
拡張機能本体と言語サーバ（LSP）を **同一リポジトリ（monorepo）** で管理する。

推奨構成：

- `packages/extension/` : VS Code Extension（UI/コマンド/クライアント）
- `packages/server/`    : Language Server（解析・補完・診断）
- `packages/core/`      : 解析コア（lexer/parser/ast/symbol/type inference）
- `resources/vbac/`     : vbac.wsf（同梱資材）
- `docs/adr/`           : ADR
- `.github/`            : PR テンプレ、Codex タスクテンプレ

---

# 4. 開発環境（Windows native）
- 開発は **Windows 直**で行う（WSL / devcontainer 前提にしない）
- 依存管理は **npm**
- Node は **LTS** を使用（可能なら `.nvmrc` を配置）
- 拡張機能の build/test/package が **Windows 直で常に通る**ことを優先する

---

# 5. 解析方針（Parser strategy）
- 解析は **手書きパーサ**で実装する（recursive descent / Pratt 等）
- 基本パイプライン：
  字句解析 → 構文解析 → AST → シンボル解決 → 型推論 → 診断/補完
- 「壊れにくい最小構成」を先に作り段階的に拡張する

Excel VBA 特有の優先対応：

- `Declare PtrSafe`
- `LongPtr`
- 条件付きコンパイル（`#If VBA7 Then` 等）

---

# 6. 初期スコープ（MVP）

## (1) 宣言部解析
- Option Explicit
- Const / Enum / Type
- Declare PtrSafe
- 基本型

## (2) 手続き解析
- Sub / Function / Property
- 引数と戻り値型
- Dim / Static

## (3) 参照解決
- スコープ管理
- 識別子補完

## (4) 診断
- 未宣言検出
- 構文エラー
- Win64/PtrSafe警告
- 型不整合の可能性警告

---

# 7. XLAM 抽出/結合機能

## 提供コマンド
- VBA: Extract from XLAM
- VBA: Combine to XLAM

## 安全ガード（必須）
- 自動バックアップ作成
- 実行前確認ダイアログ
- 入力検証
- エラーログ出力

---

# 8. Git運用・ブランチ戦略

## 8.1 ブランチ規則
- main へ直接コミット禁止
- 必ず feature/fix/refactor/docs/chore/test ブランチを使用
- 形式：`種別/要約`

## 8.2 コミット粒度
- **1コミット＝1意図**
- 機能追加とリファクタを混在させない
- ロジック変更と整形を分離
- 自動生成物は別コミット
- 無関係差分は必ず分割

## 8.3 コミットメッセージ規約
形式：
`type(scope): summary`

種別：
feat / fix / refactor / test / docs / chore

規則：
- 英語・命令形
- 簡潔
- 曖昧表現禁止

---

## 8.4 PR規則
- 1PR = 1目的
- レビュー可能サイズ維持
- 無関係変更禁止

### PR本文必須項目
1. 背景・目的
2. 変更内容
3. 影響範囲
4. 動作確認
5. リスク
6. レビュー観点

---

# 9. Codex 自動コミット・自動PR規則

## 9.1 自動コミット許可条件
- 変更が論理単位で分割済み
- 品質ゲート違反なし
- 機密情報を含まない
- 破壊的変更でない

## 9.2 自動PR許可条件
- 低リスク変更
- 仕様が明確
- 差分がレビュー可能サイズ
- 禁止領域に該当しない

## 9.3 自動PR禁止領域
- 認証・権限管理
- 課金処理
- CI/CD設定
- インフラ
- DB構造変更
- マイグレーション
- 機密情報
- 依存関係メジャー更新
- 破壊的変更

## 9.4 CodeRabbit レビュー対応規則
- PR 作成後は CodeRabbit のレビュー結果を確認対象に含める
- CodeRabbit から指摘が入った場合、Codex は以下の順で対応する
  1. 指摘内容を収集する
  2. 重複・ノイズ・誤検知の可能性を整理する
  3. 修正が妥当な指摘のみ対応する
  4. 修正理由または非採用理由を PR 上で説明できる状態にする
- CodeRabbit 指摘対応の修正は、可能な限り別コミットに分ける
- CodeRabbit 指摘を受けたあとの修正でも、無関係な変更を混ぜない
- CodeRabbit 指摘のうち、以下は人間確認を優先する
  - 仕様判断が必要なもの
  - 設計方針の変更を伴うもの
  - セキュリティ影響があるもの
  - 認証、権限、課金、CI/CD、インフラ、DB構造変更に関わるもの

## 9.5 CodeRabbit 指摘対応時の原則
- 指摘を無条件ですべて採用しない
- まず再現性・妥当性・影響範囲を確認する
- 誤検知または方針不一致の場合は、修正ではなく理由整理を優先する
- 修正後は lint / build / test を再実行する
- 修正コミット push 後、CodeRabbit の再レビュー結果も確認対象とする

## 9.6 PR 後の運用規則
- PR 作成は完了ではなく、CodeRabbit 初回レビュー確認までを1サイクルとする
- CodeRabbit の未解決指摘が残っている場合、原則として PR を完了扱いしない
- CodeRabbit の指摘が解消されたら、最終的に人間レビューへ引き渡す

---

# 10. 品質ゲート
- lint 通過必須
- テスト実行必須
- Windows 直 build/test/package 成功必須
- 失敗時は自動コミット禁止
- テスト失敗を黙殺してPR作成禁止

---

# 11. ADR
重要な設計判断は docs/adr に記録する。