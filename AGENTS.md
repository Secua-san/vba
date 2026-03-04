# AGENTS.md

このリポジトリでは、VS Code 上の Codex（OpenAI coding agent）を用いて VS Code 拡張機能を開発する。
対象は **Excel VBA** のみ。**Option Explicit 前提**、**Win64 / PtrSafe 必須**。

---

## 1. ゴール（What we build）
VBA（Excel）向けに以下を提供する VS Code 拡張機能を作る。

- IntelliSense（補完）
- 高度な型推論（段階的に精度を上げる）
- 構文チェック（Diagnostics）
- シンタックスハイライト（TextMate / 必要に応じて Semantic）

---

## 2. ソース取り扱い方針（Source strategy）
### 2.1 解析対象（主軸）
- **.bas / .cls / .frm** を直接解析する（これが主軸）

### 2.2 XLAM（バイナリ）取り扱い（補助機能）
- 拡張機能に **vbac.wsf** を同梱し、以下を提供する：
  - **XLAM からモジュール抽出（decombine）**
  - **ソースから XLAM へ結合（combine）**
- combine は破壊的操作のため、拡張機能側で **安全ガード（バックアップ・確認・検証）**を必須とする

---

## 3. リポジトリ構成（Monorepo）
拡張機能本体と言語サーバ（LSP）を **同一リポジトリ（monorepo）** で管理する。

推奨構成：
- `packages/extension/` : VS Code Extension（UI/コマンド/クライアント）
- `packages/server/`    : Language Server（解析・補完・診断）
- `packages/core/`      : 解析コア（lexer/parser/ast/symbol/type inference）
- `resources/vbac/`     : vbac.wsf（同梱資材）
- `docs/adr/`           : ADR
- `.github/`            : PR テンプレ、Codex タスクテンプレ

---

## 4. 開発環境（Windows native）
- 開発は **Windows 直**で行う（WSL / devcontainer 前提にしない）
- 依存管理は **npm**
- Node は **LTS** を使用する（可能なら `.nvmrc` を置いて揃える）
- 重要：拡張機能の build/test/package が **Windows 直で常に通る**ことを優先する

---

## 5. 解析方針（Parser strategy）
- 解析は **手書きパーサ**で実装する（recursive descent / Pratt 等）。
- パイプラインは次を基本形とする：
  - 字句解析 → 構文解析 → AST → シンボル解決 → 型推論 → 診断/補完
- 「壊れにくい最小」を先に作り、対応文法・精度を段階的に拡張する。
- Excel VBA 特有の要素を優先して扱う：
  - `Declare PtrSafe`, `LongPtr`, 条件付きコンパイル（`#If VBA7 Then` 等）

---

## 6. 初期スコープ（MVP）
MVP は以下 (1)〜(4) まで到達する。

1) 宣言部の解析（最低限）
- `Option Explicit`
- `Const`, `Enum`, `Type`
- `Declare PtrSafe`（Win64/PtrSafe 前提で解釈）
- 主要な型トークン（`Long`, `LongPtr`, `Integer`, `String`, `Boolean`, `Variant`, `Object` など）

2) 手続きの解析（最低限）
- `Sub/Function/Property`（Public/Private/Friend）
- 引数（ByVal/ByRef/Optional/ParamArray）と戻り値型
- `Dim/Static` など変数宣言の最小対応

3) 参照解決（補完の土台）
- 未宣言/宣言済み識別子の管理（スコープ：プロシージャ内→モジュール→プロジェクト順に段階拡張）
- 変数/定数/プロシージャ呼び出しの補完（まずは同一ファイル/モジュール内中心）

4) 診断（Diagnostics）
- 未宣言（Option Explicit 前提）
- 明らかな構文エラー（パース失敗箇所の報告）
- Win64/PtrSafe 関連の警告（例：PtrSafe なし、LongPtr の不整合などを段階的に追加）
- 型不一致 “っぽい” 箇所の警告（推論が確定できない場合は過剰に断定しない）

---

## 7. XLAM 抽出/結合（vbac.wsf）仕様（Extension commands）
### 7.1 提供コマンド
- **VBA: Extract from XLAM (decombine)**
  - `.xlam` から `.bas/.cls/.frm` を抽出して管理フォルダへ出力
- **VBA: Combine to XLAM (combine)**
  - 管理フォルダのソースを `.xlam` に反映

### 7.2 安全ガード（必須）
- combine 実行前に必ずバックアップを作成する  
  - 例：`target.xlam.bak-YYYYMMDDHHmmss`
- combine 実行前に明示確認（対象パス、バックアップ先、上書きの注意）
- 入力検証：
  - 管理フォルダ（src）が空なら combine を拒否
  - 対象ファイルが存在しない/ロック中なら拒否
- ログ：
  - stdout/stderr は OutputChannel に必ず表示し、失敗理由を追えるようにする

---

## 8. ブランチ運用（feature ブランチ + PR）
- `main` は常にビルド可能な状態を保つ（壊れた状態を入れない）
- 作業は **必ず `feature/*` ブランチ**を切って行う
- 変更は PR で `main` に取り込む

### PR の粒度
- 1 PR = 1 目的（例：lexer の導入、diagnostic の追加、decombine コマンド追加）
- PR が大きくなる場合は段階 PR に分割する

### コミット規約（推奨）
- `feat:` 新機能
- `fix:` バグ修正
- `refactor:` 内部改善（外部仕様不変）
- `test:` テスト追加/修正
- `docs:` ドキュメント
- `chore:` 依存更新/雑務

---

## 9. コマンド（npm）
※実際の scripts 名称は `package.json` に合わせる。

- 依存導入：`npm ci`（初回/CI） / `npm install`（ローカル）
- ビルド：`npm run build`
- 監視ビルド：`npm run watch`
- テスト：`npm test`
- Lint：`npm run lint`
- パッケージ：`npm run package`（vsix 生成）

---

## 10. 品質ゲート（Quality gates）
- 変更のたびに最低限 lint を通す。
- 解析（lexer/parser/inference/diagnostics）はユニットテストを最小でも付ける。
- Windows 直で `build/test/package` が通ることを常に優先する。

---

## 11. ADR（Architecture Decision Records）
重要な意思決定は `docs/adr/` に ADR として残す。
例：
- なぜ手書きパーサか
- スコープ解決のルール
- 型推論のフェーズ設計
- VBA7 条件付きコンパイルの扱い
- vbac.wsf の同梱と安全ガード方針