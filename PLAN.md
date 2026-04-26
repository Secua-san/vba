# PLAN

この文書は、フェーズ表と現在の実装実態を照合したロードマップである。  
タスクを切り替えたあとに再開するときは、まずこの文書で現在位置を確認し、その後に [`TASKS.md`](TASKS.md) を見る。詳細な履歴や判断経緯が必要な場合だけ [`TASKLOG.md`](TASKLOG.md) を開く。

## 使い方

1. まずこの文書の「現在位置」と「優先ロードマップ」を確認する
2. 次に [`TASKS.md`](TASKS.md) で直近の主タスクと次タスクを見る
3. 作業再開前に `git status --short` で未コミット差分の有無を確認する
4. 実装対象に対応する要件書 / ADR / process 文書だけを追加で読む
5. 1 PR に収まる粒度の小タスクを 1 つ選び、そのまま実装へ入る

## 現在位置

- 現在の主実装軸は Phase 3 の AST 安定化・構文情報整備へ移る
- 直近完了は Phase 2 の `ProcedureStatementNode` structured AST coverage、formatter compressed block 判定の structured kind 接続、local rename target range の structured segment 優先化、active workbook identity snapshot transport の安定化
- リポジトリ全体としては、LSP の user-facing 機能が先行しており、parser / AST を後追いで構造化して基盤を固めている状態
- したがって、単純な Phase 0 -> 12 の直列進行ではなく、Phase 2-3 を再度強化しながら Phase 6-10 の既存機能を壊さない進め方が必要

## フェーズ進捗サマリ

| Phase | 状態 | 現状 | 次のゲート |
| --- | --- | --- | --- |
| Phase 0: リポジトリ基盤整備 | 完了 | `core` / `server` / `extension` の分離、`build` / `test` / `lint` / `package`、fixture と test host が揃っている | 維持のみ |
| Phase 1: 字句解析 | 完了 | `lexDocument.ts` と token range があり、行継続、コメント、文字列、日付、directive、型サフィックス、属性行を扱えている | 維持のみ |
| Phase 2: 構文解析 MVP | 完了 | module / procedure / declare / enum / type / variable に加え、assignment / call、主要 block statement、label付き statement、termination statement の structured node を core / server 回帰で固定した | 維持のみ |
| Phase 3: AST 安定化・構文情報整備 | 進行中 | `range` / `headerRange` / `nameRange` はある。formatter compressed block 判定と local rename target range は structured segment 優先になったが、未構造化 `executableStatement` 互換 fallback は残る | 後続処理を AST kind ベースへさらに寄せる |
| Phase 4: シンボルテーブル・スコープ解析 | 部分実装 | module / procedure scope の symbol 抽出、定義ジャンプ、参照検索、rename の基盤はある | structured statement を前提に symbol 解決を安定化する |
| Phase 5: 名前解決・基本型推論 | 部分実装 | explicit / assignment / return 起点の型推論、built-in owner 解決、worksheet control 系の限定的解決がある | structured AST と一貫した式・文ベースの推論へ寄せる |
| Phase 6: Diagnostics | 部分実装 | syntax、未宣言、重複、未使用、write-only、到達不能、ByRef 系などの診断がある | text 依存の判定を AST / symbol / type ベースへ置き換える |
| Phase 7: IntelliSense / 補完 MVP | 部分実装 | completion、hover、signature help、semantic token、document symbol は user-facing に動いている | 文脈絞り込みと型起点の補完精度を上げる |
| Phase 8: 高度な型推論・実行時バインディング対応 | 部分実装 | workbook binding、active workbook snapshot、worksheet control sidecar による限定解決はある | `CreateObject` / ProgID 系など VBA 特有の runtime binding 戦略を追加する |
| Phase 9: 定義ジャンプ・参照検索・シンボルナビゲーション | 部分実装 | definition / references / rename / document symbol / workspace 連携の土台がある | parser / symbol 基盤強化後に安定度と coverage を上げる |
| Phase 10: vbac.wsf / xlam 連携 | 部分実装 | extract / combine コマンドはある | combine の安全性、エラーハンドリング、ログ、検証を強化する |
| Phase 11: 品質強化・回帰防止 | 進行中 | shared case spec、fixture、server / extension ミラー回帰の整備が進んでいる | parser / AST 強化に追随する回帰セットを足す |
| Phase 12: 最小ドキュメント整備 | 最小維持 | `TASKS.md` / `TASKLOG.md` / docs 入口は整理済み | 実装差分に直接関係する最小更新だけ行う |

## 優先ロードマップ

### 最優先トラック: Parser / AST の基盤固め

このリポジトリは user-facing 機能が先行しているため、当面の最優先は Phase 3 の AST 安定化である。
Phase 2 の structured AST coverage は完了済みとして維持し、raw text fallback を後続処理から段階的に減らす。

#### 直近に実装する順序

1. structured statement を使う diagnostics / symbol 連携の raw text 依存を段階的に減らす
2. navigation / references / semantic token に残る fallback 判定を AST segment 優先へ寄せる
3. structured AST 対応の core / server / extension 回帰を維持する
4. `CreateObject("WScript.Shell")` 既知 ProgID 解決の最小実装へ進む

#### このトラックの完了条件

- block statement の主要構文が text ではなく node kind で判定できる
- `range` / `text` は互換維持のため残しつつ、判断の主軸を AST へ移せる
- parser 強化のたびに server / extension の回帰で壊れない

### 次トラック: AST を後続フェーズへ接続

Parser / AST が安定したら、Phase 4-7 で既にある user-facing 機能を AST ベースへ寄せる。

#### 優先順

1. symbol / scope 抽出の structured statement 対応
2. diagnostics の text 走査依存を段階的に削減
3. type inference の statement / expression 粒度を上げる
4. 補完の文脈絞り込みと誤候補抑制を強化する

#### 狙い

- parser と language service の二重ロジックを減らす
- 診断・補完・ナビゲーションの誤検知を下げる
- 今後の `CreateObject` や Office object 連携を AST / symbol / type の通常パイプラインへ載せやすくする

### その次のトラック: VBA / Excel 固有機能の拡張

基盤が安定したら、既にある workbook / worksheet control 系の導線を広げる。

#### 候補

1. `CreateObject("ProgID")` の既知 ProgID マッピング first slice
2. `Variant` / `Object` の暫定補完戦略
3. workbook / sheet / control 系の既存 sidecar 連携の coverage 拡大
4. `vbac.wsf` の combine 安全性と失敗時ログの強化

## 次の実装候補

次タスクは、原則として次の順で選ぶ。

1. diagnostics / symbol 連携に残る `executableStatement` fallback の整理
2. references / semantic token の residual text 依存削減
3. structured AST 対応の core / server / extension 回帰維持
4. `CreateObject("WScript.Shell")` 既知 ProgID 解決の最小実装

## 1 PR 粒度の目安

以下を超えそうなら、新規タスクに切る。

- 対象責務が parser から diagnostics、または diagnostics から workbook integration へ跨ぐ
- 対象ディレクトリが `packages/core` 中心から `packages/server` / `packages/extension` 中心へ移る
- 3 ターン以上実コード変更が止まり、整理や議論が前に出る
- 同一チャットを続けるより、別 PR に分けた方がレビューしやすい

## タスク切替時の再開手順

### 続きから入るとき

1. [`PLAN.md`](PLAN.md) の「現在位置」「優先ロードマップ」「次の実装候補」を確認する
2. [`TASKS.md`](TASKS.md) で現在の主タスクと次タスクを確認する
3. `git status --short` を見て、未コミット差分を壊さない前提で着手する
4. 対象が parser 系なら `packages/core/src/parser/parseModule.ts` と `packages/core/src/types/model.ts` を先に開く
5. 対象が LSP 回帰なら `packages/server/src/lsp/documentService.ts` と該当 test fixture を先に開く
6. 実装後は対象差分に対応する最小テストだけ先に回し、最後に必要な全体検証を行う

### 新規タスクを切るべきタイミング

- 3 ターン以上コード変更が出ていない
- 整理や方針議論が主になっている
- フェーズ境界を跨ぐ
- 次の作業が別 PR 相当になっている

## 再開時にまず見るファイル

- 進捗サマリ: [`TASKS.md`](TASKS.md)
- 履歴: [`TASKLOG.md`](TASKLOG.md)
- parser / AST: [`packages/core/src/parser/parseModule.ts`](packages/core/src/parser/parseModule.ts), [`packages/core/src/types/model.ts`](packages/core/src/types/model.ts)
- diagnostics / symbol / inference: [`packages/core/src/diagnostics/analyzeModule.ts`](packages/core/src/diagnostics/analyzeModule.ts), [`packages/core/src/symbol/buildModuleSymbols.ts`](packages/core/src/symbol/buildModuleSymbols.ts), [`packages/core/src/inference/inferModuleTypes.ts`](packages/core/src/inference/inferModuleTypes.ts)
- LSP 接続点: [`packages/server/src/lsp/documentService.ts`](packages/server/src/lsp/documentService.ts)

## この文書の扱い

- `PLAN.md` はフェーズ進捗とロードマップの正面入口とする
- 直近の進行中タスクそのものは [`TASKS.md`](TASKS.md) を正本とする
- 詳細な完了履歴は [`TASKLOG.md`](TASKLOG.md) に寄せる
- この文書は、フェーズの進み具合か優先トラックが変わったときだけ最小更新する
