# TASKS

## 進行中

- なし

## 完了

- [x] M0-M2 基盤実装のマージ
  - Core の lexer / parser / symbol / diagnostics パイプライン
  - 最小構成の LSP サーバーと VS Code 拡張接続
  - Windows での build / test / package フロー

- [x] M3 ワークスペースシンボル索引
  - モジュール名と標準モジュールの `Public` / `Friend` シンボルをワークスペース全体で索引化
  - ファイル横断の補完と定義ジャンプを追加
  - 閉じたファイルはディスク内容へ戻し、ファイル変更通知で索引を更新

- [x] ワークスペース考慮の診断
  - 他モジュールの公開シンボルに一意に一致する場合のみ `undeclared-variable` を抑制
  - あいまいな候補は診断を残す保守動作に固定

- [x] ワークスペース解析のテスト
  - server にファイル横断補完、定義ジャンプ、あいまい名の診断テストを追加
  - extension のスモークテストに複数ファイルの補完と定義ジャンプ確認を追加

- [x] 日本語コミットと PR 作成
  - コミット、PR 本文、レビュー対応を日本語で実施
  - PR #2 から PR #4 までをマージし、M3 と M4 の基盤機能を `main` へ反映

- [x] M3 ワークスペース参照検索
  - `Find References` のための最小参照索引を追加
  - 同一モジュールと標準モジュール公開シンボルの参照検索を実装
  - server / extension のテストでローカル参照とファイル横断参照を確認

- [x] M4 型推論基盤
  - 明示型、リテラル型、単純代入、戻り値代入の最小推論を追加
  - 推論結果を使った単純な型不一致診断を追加
  - completion detail に型情報を表示

- [x] 型連動補完
  - 代入先の推論型に基づいて completion 候補を絞り込む
  - 引数ヒントに現在の引数型を表示する

- [x] 型不一致診断の拡張
  - 複合式、`Set` 代入、Variant を含むケースへ warning 判定を広げる

- [x] CI ラベル自動付与設定の修正
  - `actions/labeler@v5` に合わせて `.github/labeler.yml` の形式不整合を解消

- [x] ByRef / ByVal 危険箇所の診断
  - 同一モジュール内の `ByRef` 呼び出しで、式渡しと型不一致を warning として追加
  - server でワークスペース公開手続きへの `ByRef` 警告も補完

- [x] Set 必須箇所検出
  - 参照型への代入で `Set` が必要なケースを warning として追加
  - `Set` を足せば整合するケースでは `type-mismatch` より `set-required` を優先

- [x] 重複定義の診断
  - モジュールスコープと手続きスコープで衝突する宣言を `duplicate-definition` error として追加
  - パラメータ、ローカル変数、ローカル定数、手続き宣言の重複を検出

- [x] 到達不能コードの診断
  - `Exit Sub` / `Exit Function` / `Exit Property` / `End` の後に続く同一到達領域の文を `unreachable-code` warning として追加
  - `Else` / `Case` / ループ終端 / ラベルで保守的に検出を打ち切り、誤検知を抑制

- [x] 未使用変数の診断
  - 手続きローカルの変数と引数について、実行文で一度も参照されない宣言を `unused-variable` warning として追加
  - 読み書きの区別はせず、書き込みだけの変数は今回の段階では警告対象に含めない

- [x] write-only 代入の診断
  - 代入されるが読み出されないローカル変数を `write-only-variable` warning として追加
  - `unused-variable` とは重複させず、ローカル変数のみを対象にする

- [x] 継続行の型不一致診断修正
  - 行継続 `_` を使った代入文も型推論対象に含める
  - core / server のテストで継続行から `type-mismatch` が返ることを確認

- [x] ローカル変数の安全リネーム
  - 同一手続き内の procedure-scope 変数だけを `prepareRename` / `rename` の対象にする
  - 新しい名前が不正、または同一手続きや可視シンボルと衝突する場合は保守的に拒否する

- [x] セマンティックハイライト
  - LSP の full document semantic tokens を追加し、変数、引数、定数、手続き、型、列挙体メンバーを色分け対象にする
  - server / extension のテストで legend と token 配列が返ることを確認する

- [x] VBA コードスニペット
  - `Sub` / `Function` / `Property` / `If` / `For` / `Select Case` / `Do While` / `With` の snippets を追加する
  - extension の smoke test で snippet completion が読み込まれることを確認する

- [x] 拡張機能開発ホストの起動導線
  - ルート `.vscode` に build 付きの `extensionHost` 起動設定を追加する
  - `npm run dev:host` と `npm run test:host` で CLI からも実機確認と拡張テストを起動できるようにする
  - README に開発ホストの確認手順を追記する

- [x] VBA 構文インデント
  - `core` に `If` / `Select Case` / `For` / `Do` / `With` / `Property` を基準にした最小インデント formatter を追加する
  - `.frm` のデザイナー領域は保持しつつ、コード領域だけを整形する
  - server の document formatting provider と extension の smoke test で整形結果を確認する

- [x] 継続行整形
  - `_` を使う代入、引数列、メソッドチェーンの hanging indent を formatter で安定化する
  - 引数列の閉じ括弧だけを base indent に戻し、継続行の `_` 前後も最小限正規化する
  - core / server / extension のテストで継続行専用 fixture の整形結果を確認する

- [x] CodeRabbit 待機時間の見直し
  - PR #11 から PR #24 の実測を確認し、初回反応時間、進行中コメントの更新完了時間、レート制限待機時間を整理する
  - `docs/process/coderabbit-review.md` に実測値と新しい待機基準を追記する
  - 初回レビュー待ちの標準待機と、レート制限時の自動停止しきい値を見直す

## 次候補

- [ ] ブロック整形
  - `ElseIf` / `Case` / `#If` 系を含むブロック境界の改行と整形ルールを整理する
  - 構文インデント formatter を土台に、継続行以外の block layout を段階的に整える

## メモ

- `docs/requirements/000-overview.md` にはユーザー管理の差分があるため、自動コミットに含めない
