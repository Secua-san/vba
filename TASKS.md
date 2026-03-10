# TASKS

## 進行中

- なし

## 完了

- [x] 組み込みメンバー署名データの拡張（第7弾）
  - 現行 Microsoft Learn を再確認し、`WorksheetFunction` には `XLookup` / `XMATCH` が未掲載のままであることを確認
  - `Range.Address` / `Range.AddressLocal` の署名を取り込み、`ActiveCell` / `Cells` のような Range 系組み込みルートからも解決できるように修正
  - `Address` 系の optional 引数メタデータと `XlReferenceStyle` 型情報を server / extension テストで回帰確認

- [x] 組み込みメンバー署名データの拡張（第6弾）
  - `WorksheetFunction.Choose` / `WorksheetFunction.Transpose` を Microsoft Learn 由来の署名抽出対象へ追加し、参照 JSON を再生成
  - `Choose` の可変長必須引数と `Transpose` の単一必須引数を server / extension テストで回帰確認
  - `XLookup` / `XMATCH` / `Address` は現行 Learn JSON で確認できなかったため、次候補で再整理する

- [x] CodeRabbit レビュー要約ログ運用の追加
  - `docs/process/coderabbit-review-summaries.md` を新規追加し、PR ごとのレビュー要約テンプレートと記録を追加
  - 要約には「この作業で当てはまりそうな内容（横展開候補）」を必須項目として定義
  - `docs/process/coderabbit-review.md` / `docs/process/sub-agent-escalation.md` / `AGENTS.md` にログ追記ルールを反映

- [x] PR 前サブエージェント設定の `reviewer` 既定化
  - `docs/process/coderabbit-review.md` の PR 前セルフレビュー担当を `explorer` から `reviewer` へ変更
  - `docs/process/sub-agent-escalation.md` の PR 前必須レビュー担当を `reviewer` へ変更し、`config.toml` / `reviewer.toml` の確認手順を追加
  - `AGENTS.md` にも同方針を反映し、PR 作成前は `reviewer` を使う運用に統一

- [x] 組み込み署名データ第5弾レビュー修正
  - extension の `BuiltInMemberSignature` テストで、追加メソッド以降の `vscode.Position` を文字列検索ヘルパー経由に変更し、fixture 行変更への耐性を向上
  - 参照 JSON 生成時の `generatedAt` を出力対象から外し、再生成時の差分ノイズを削減

- [x] PR 前サブエージェント自己レビュー運用の追加
  - `docs/process/coderabbit-review.md` に「PR作成前のセルフレビュー（サブエージェント）」を追加
  - `docs/process/sub-agent-escalation.md` に「PR前の必須レビュー」を追加
  - 次回以降は PR 作成前に `reviewer` を既定として差分レビューを実施し、結果要約後に PR を作成する

- [x] 組み込みメンバー署名データ拡張（第5弾）
  - `WorksheetFunction` の参照・統計系メソッド（`Match` / `Index` / `Lookup` / `HLookup`）を署名抽出対象へ追加し、Microsoft Learn 参照 JSON を再生成
  - server / extension テストに上記 4 メソッドの署名ヘルプ検証を追加し、`Match` / `Index` / `Lookup` / `HLookup` の省略可能引数メタデータを回帰監視
  - extension fixture に新規4メソッド呼び出しを追加し、署名ヘルプと fallback 抑止の既存ケースが崩れないことを確認

- [x] ariawase ライセンス表記の追加
  - `THIRD_PARTY_LICENSES.md` を新規追加し、`vbaidiot/ariawase`（MIT）の出典リンクとライセンス原文を記載
  - ルート `README.md` にサードパーティライセンス一覧への導線を追加
  - 拡張機能配布物に含まれる `packages/extension/README.md` にも `ariawase` のライセンス情報を追記

- [x] 組み込みメンバー署名データ拡張（第4弾レビュー修正）
  - `WorksheetFunction.Or` / `WorksheetFunction.Xor` の `Arg2` 以降で不足していた `dataType` / `description` / `isRequired` を再生成ロジック側で補完
  - 署名パラメータ展開で `Arg1-Arg30` / `Arg1...Arg30` / `Arg1…Arg30` の表記ゆれを扱えるようにして、可変引数判定の `…` も吸収
  - server / extension テストに `Or` / `Xor` 第2引数の `Variant` と省略可能フラグの回帰確認を追加

- [x] 組み込みメンバー署名データ拡張（第4弾）
  - `WorksheetFunction` の論理・集計系メソッド（`And` / `Or` / `Xor` / `CountA` / `CountBlank`）を署名抽出対象へ追加し、Microsoft Learn 参照 JSON を再生成
  - server / extension テストに上記 5 メソッドの署名ヘルプ検証を追加し、`And` / `CountA` の省略可能引数メタデータを回帰監視
  - `Application` 側 fallback の抑止ケースとして、`ActiveCell`（property）と `NewWorkbook`（event）の呼び出しでも署名を返さないことを fixture 単位で確認

- [x] 組み込みメンバー署名データ拡張（第3弾）
  - `WorksheetFunction` の日付/文字列/検索系メソッド（`EDate` / `EoMonth` / `Text` / `Find` / `Search` / `VLookup`）を署名抽出対象へ追加し、Microsoft Learn 参照 JSON を再生成
  - server / extension テストに上記 6 メソッドの署名ヘルプ検証を追加し、`Find` / `Search` の省略可能引数メタデータも回帰監視
  - `WorksheetFunction.Find` の誤った説明文を生成スクリプト側の override で補正
  - fallback signature help の抑止ケースとして、`Application.WorksheetFunction()`（property）と `Application.AfterCalculate()`（event）が署名対象にならないことを確認

- [x] 組み込みメンバー署名データ拡張（第2弾レビュー修正）
  - 可変引数展開時に parameter table 名との数値サフィックス対応を追加し、`Max` / `Min` の `Arg30` でも `dataType` / `description` / `label` を復元
  - 署名生成前に `...` を除外した parameter name 解決を追加し、`signatureLabel` と parameter metadata の不整合を防止

- [x] 組み込みメンバー署名データ拡張（第2弾）
  - `WorksheetFunction.Average` / `Count` / `Max` / `Median` / `Min` を署名抽出対象に追加し、Microsoft Learn 参照 JSON を再生成
  - ParamArray 系の `Arg1..ArgN` かつ `...` を含む署名では、個別メソッド分岐ではなく汎用ルールで `Arg2` 以降を省略可能へ補正
  - server / extension テストに `Average` の可変引数署名を追加し、fallback 表示との差分を回帰確認

- [x] 組み込みメンバー署名データ拡張（第1弾）
  - `signatureMemberAllowList` を拡張し、`Application.CalculateFull` / `CalculateFullRebuild` / `CalculateUntilAsyncQueriesDone` と `WorksheetFunction.Power` / `Round` の署名を Microsoft Learn から再生成
  - 署名未収録の built-in callable でも、`Application.OnTime()` のような fallback signature help を返す保守動作を追加
  - server / extension テストに署名拡張と fallback の回帰確認を追加

- [x] 組み込みメンバー署名のレビュー修正
  - `WorksheetFunction.Sum` の署名データで `Arg2` 以降を `省略可能` として扱うよう再生成ロジックを補正
  - `WorksheetFunction.Sum` と同名の公開手続きが存在する場合でも、signature help が組み込みメンバーを優先するように修正
  - server / extension テストに必須・省略可能引数の期待値と衝突ケースの回帰確認を追加

- [x] MCP サーバー呼び出しの 429 対策
  - 共通の retry / rate-limit ヘルパーを追加し、`429`、`Retry-After`、指数バックオフ + ジッター、最大再試行超過時の明確な失敗を実装
  - 呼び出し間隔の制御と in-flight 重複抑止を追加し、対象 MCP 名、retry 回数、待機時間、最終失敗理由を構造化ログへ出力
  - `scripts/generate-mslearn-vba-reference.mjs` を共通ヘルパーへ移行し、スクリプト用テストを root `npm test` に組み込む

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

- [x] ブロック整形
  - `If / ElseIf / Else / Select Case / Case / #If / #Else / #End If` の圧縮ブロックを formatter で複数行へ展開する
  - 通常の `:` 区切り文は維持し、ブロック境界に関わる行だけを安全側で分離する
  - core / server / extension の整形テストで block layout の結果を確認する

- [x] 宣言整列
  - 単一行の `Dim` / `Const` / `Declare` を対象に、連続する宣言ブロック内で `As` / `=` / `Lib` の位置を限定的に揃える
  - 複数宣言、継続行、通常文の `:` 区切りは対象外にし、既存の block layout formatter と競合しないようにする
  - core / server / extension の整形テストで declaration alignment の結果を確認する

- [x] コメント整形
  - 行頭コメントと末尾コメントを対象に、コメントマーカー前後の空白を formatter で最小限正規化する
  - コメント位置の移動は行わず、`'` と `Rem` の空白だけを保守的に整える
  - core / server / extension の整形テストで comment formatting の結果を確認する

- [x] CodeRabbit 待機時間の見直し
  - PR #11 から PR #24 の実測を確認し、初回反応時間、進行中コメントの更新完了時間、レート制限待機時間を整理する
  - `docs/process/coderabbit-review.md` に実測値と新しい待機基準を追記する
  - 初回レビュー待ちの標準待機と、レート制限時の自動停止しきい値を見直す

- [x] Microsoft Learn 参照一覧の取得基盤
  - Excel、VBA 言語リファレンス、Office library reference の一覧を Microsoft Learn から JSON 化する再生成スクリプトを追加
  - `resources/reference/mslearn-vba-reference.json` に、補完とハイライト用のオブジェクト、列挙、キーワード、定数カテゴリ、関数、文を保存する

- [x] Option Explicit 補完
  - `Option Explicit` が無いモジュールに対する quick fix code action を追加し、先頭付近へ安全に挿入できるようにする
  - `.frm` のデザイナー領域と属性行を壊さず、既存 option 行の直後へ重複なく挿入する
  - server / extension のテストで標準モジュールと `.frm` の挿入結果を確認する

- [x] 組み込み参照データの補完連携
  - Microsoft Learn 由来の参照 JSON を `core` の shared built-in index に正規化し、Excel / VBA / Office の組み込みオブジェクト、定数、キーワードを補完候補へ追加する
  - 未宣言診断と rename 禁止名に同じ reserved / built-in 判定を使い、`Application`、`xlAll`、`Beep` などの誤警告を抑制する
  - semantic token に built-in function / constant / keyword を追加し、server / extension のテストで legend と token を確認する

- [x] 組み込みメンバー補完
  - `Application.` や `WorksheetFunction.` のような member access に対して、Microsoft Learn 由来のメソッド / プロパティ候補を返す
  - `Application.WorksheetFunction.` のような既知 chain も、built-in member の型名を使って段階的に解決する
  - built-in member の簡易ドキュメントと semantic token を server / extension のテストで確認する

- [x] 組み込みメンバーのシグネチャ支援
  - `WorksheetFunction.Sum` と `Application.Calculate` について、Microsoft Learn 由来の署名と説明を参照 JSON へ追加
  - built-in callable の signature help を追加し、`Application.WorksheetFunction.Sum` の chain でも同じ署名を返す
  - built-in callable の hover を追加し、署名、要約、Microsoft Learn リンクを表示する

## 次候補

- [ ] 既存署名メタデータの横断点検
  - `WorksheetFunction` と `Range` を含む既存署名について、optional / variadic / return type の欠落がないかを横断確認する
  - Microsoft Learn 側の更新で `XLookup` / `XMATCH` が追加された場合に備え、次回再生成時の確認観点を整理する

## メモ

- `docs/requirements/000-overview.md` にはユーザー管理の差分があるため、自動コミットに含めない
