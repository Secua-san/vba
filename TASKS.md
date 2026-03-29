# TASKS

## 進行中

- [ ] application workbook root family の extension-only interaction / semantic entry を shared spec に残す境界を整理する
  - `applicationWorkbookRoot` には `worksheetBroadRoot` より多い extension-only scope entry があるため、family canonical anchor と package-local adapter の境界を同じ物差しで整理する
  - `hover` / `signature` / `semantic` の extension-only entry を local file へ戻す条件があるか、`scopes` 契約の適用例として shared spec に残すべきかを切り分ける

## 完了

- [x] worksheet broad root family の extension-only non-target hover / signature negative を shared spec に残す要否を整理する
  - [docs/process/worksheet-broad-root-extension-only-interaction-shared-spec-feasibility.md](docs/process/worksheet-broad-root-extension-only-interaction-shared-spec-feasibility.md) を追加し、`worksheetBroadRoot.hover.negative` / `signature.negative` は server scope が無くても canonical anchor source として shared spec に残す方針を整理した
  - [docs/process/workbook-root-family-case-table-policy.md](docs/process/workbook-root-family-case-table-policy.md) を更新し、`scopes` 配列を持つ shared spec では scope 非対称だけを理由に local file へ戻さず、anchor / reason が family canonical source かどうかで判断する共通ルールを追記した
  - 次タスクを broad root family から `applicationWorkbookRoot` の extension-only entry 境界整理へ切り替えた

- [x] worksheet broad root family の server non-target hover / signature negative を shared scope へ広げる要否を整理する
  - [docs/process/worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md](docs/process/worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md) を追加し、server 側は non-target completion negative だけ shared 化し、hover / signature の non-target negative は extension E2E の shared entry に留める方針を整理した
  - server 側では positive hover / signature anchor に対する `snapshot unavailable` / `mismatch` の closed-state coverage を既に固定しているため、non-target interaction negative を足しても得られる保証が小さいことを整理した
  - 次タスクを「server scope へ広げるか」から、「extension-only interaction negative を shared spec に残すべきか」へ切り替えた

- [x] worksheet broad root family の shadow negative coverage を direct OLE route 以外へ広げる要否を整理する
  - [docs/process/worksheet-broad-root-shadow-coverage-feasibility.md](docs/process/worksheet-broad-root-shadow-coverage-feasibility.md) を追加し、broad root shadow は root identifier の built-in 判定で閉じるため、`Shapes("CheckBox1").OLEFormat.Object` や root `.Item("Sheet One")` を追加しても別分岐を踏まないことを整理した
  - direct `OLEObjects("CheckBox1").Object` route の 2 本で root gating coverage は十分とし、route coverage は `getWorkbookRootFamilyBuiltinContext()` より後段の分岐が増えた時だけ再評価する方針を明文化した
  - 次タスクを broad root shadow ではなく、server 側 non-target hover / signature negative の shared scope 要否整理へ切り替えた

- [x] worksheet broad root family の server-only shadow negative を helper / table へ寄せる要否を整理する
  - [docs/process/worksheet-broad-root-shadow-server-helper-feasibility.md](docs/process/worksheet-broad-root-shadow-server-helper-feasibility.md) を追加し、`Worksheets` shadow / `Application` shadow は現時点では 2 本の明示 test のまま維持し、server-local helper / table へは寄せない方針を整理した
  - broad root shadow test の重複量は小さく、shadow 宣言の違い自体が test の主語になっているため、抽象化を足すより明示性を維持する方が読みやすいと整理した
  - 再評価トリガーを「shadow variant 増加」「coverage 拡張」「レビュー指摘の連続」「triage 上の主語不足」に限定し、shared spec や extension matrix へ先回りしないことを明文化した

- [x] worksheet broad root family の symbol-shadowed case を extension matrix へ広げる要否を整理する
  - [docs/process/worksheet-broad-root-shadow-extension-matrix-feasibility.md](docs/process/worksheet-broad-root-shadow-extension-matrix-feasibility.md) を追加し、`Worksheets` shadow / `Application` shadow は現時点では server-only negative case のまま維持し、extension matrix へは広げない方針を整理した
  - broad root shadow は root identifier の built-in 判定で閉じる conservative negative であり、host 非同期や sidecar lookup の主要リスクではないため、server unit test を最短経路として維持する判断を記録した
  - `worksheetBroadRoot` shared spec に `state: "shadowed"` を追加せず、extension 側の dedicated shadow fixture も導入しないこと、再評価トリガーを broad root shadow の実不具合 / review drift / variant 増加に限定することを明文化した

- [x] unqualified worksheet broad root の root `.Item("SheetName")` 形式を再評価
  - `packages/server/src/lsp/documentService.ts` の broad root root-selector 解釈を拡張し、`Worksheets.Item("SheetName")` / `Application.Worksheets.Item("SheetName")` を direct call form と同じ sheetName literal として扱えるようにした
  - `packages/server/test/documentService.test.js`、`packages/extension/test/suite/index.ts`、`packages/extension/test/fixtures/WorksheetBroadRootBuiltIn.bas` を更新し、snapshot 未一致では閉じ、manifest match 時には completion / hover / signature help が開くことを root `.Item` と child `.Item` の両方で固定した
  - ADR [0005](docs/adr/0005-explicit-sheet-name-root-policy.md) と [docs/process/explicit-sheet-name-broad-root-feasibility.md](docs/process/explicit-sheet-name-broad-root-feasibility.md) を更新し、`Worksheets.Item` は既定メンバーとして direct call form と同一 family で扱う根拠を Microsoft Learn ベースで追記した
  - `npm run lint`、`npm test`、`npm run package` を通して broad root family と既存の非対象境界が崩れていないことを確認した

- [x] broad root family の重複テスト補助を整理
  - `packages/server/test/documentService.test.js` に worksheet broad root 専用 fixture / snapshot / token lookup helper を追加し、`matched` / `mismatched` / `unavailable` / root `.Item` 非対象ケースの重複した setup と query を削減した
  - `packages/extension/test/suite/index.ts` に token 指定の completion / hover / signature helper を追加し、`WorksheetBroadRootBuiltIn.bas` の broad root 行列テストを状態別の配列駆動へ寄せた
  - shadow case はこの段階では server-only の単発 negative test のまま維持し、shared spec や extension matrix へは上げていない
  - `npm run lint --workspace @vba/server`、`npm run lint --workspace vba-extension`、`npm run test --workspace @vba/server`、`npm run test --workspace vba-extension`、`npm test`、`npm run package` を通して既存 broad root 挙動を維持した

- [x] ドキュメント導線と CodeRabbit 記録の整理
  - `docs/README.md` と `docs/process/README.md` を、「最初に読む正本」と「必要なときだけ開く機能別メモ」に整理し、入口からの参照階層を浅くした
  - `docs/process/coderabbit-review.md` を正本に寄せ、CodeRabbit 月次ログは `docs/process/coderabbit-review-logs/YYYY-MM.md` へ直接記録する方式に変え、案内用の中継ページを廃止した
  - `docs/process/mslearn-signature-regeneration.md`、`docs/process/sub-agent-escalation.md`、`AGENTS.md` からレビュー記録まわりの重複説明を削り、ログは記録専用で参照しない方針にそろえた

- [x] Worksheet control metadata sidecar artifact の仕様整理
  - sidecar 正本パスを `<bundle-root>/.vba/worksheet-control-metadata.json` に固定し、nearest ancestor lookup と workspace root での打ち切り規則を [docs/process/worksheet-control-metadata-sidecar-artifact.md](docs/process/worksheet-control-metadata-sidecar-artifact.md) に整理した
  - `owners[]` に `worksheet` / `chartsheet` を統一し、未対応 owner を `status: "unsupported"` と `reason` 付き record で表す schema v1 を定義した
  - `OLEObject.Object` 後段型付けと `Sheet1.CommandButton1` 支援の両方に必要な最小 field を `sheetCodeName` / `shapeName` / `codeName` / `controlType` / `progId or classId` として固定し、現行 probe 出力との差分を整理した

- [x] Worksheet workbook package control metadata probe の最小実装
  - `scripts/lib/workbookControlMetadata.mjs` を追加し、Open XML package の workbook / worksheet / drawing / control data part をたどって `sheetCodeName`、`shapeName`、`codeName`、`shapeId`、`ProgID`、`classId` を抽出できるようにした
  - `scripts/probe-workbook-control-metadata.mjs` と `npm run probe:worksheet-control-metadata` を追加し、workbook package から JSON を標準出力または `--out` で書き出せるようにした
  - `scripts/test/workbookControlMetadata.test.mjs` で synthetic workbook package を組み立て、`shape name != code name`、`shapeId` 結合、`ProgID` / `classId`、chart sheet 除外、CLI 出力を固定した

- [x] Worksheet / Chart control metadata source の PoC
  - workbook package を第 1 PoC source とし、worksheet 側では `sheetPr@codeName`、`controls/control@name|shapeId|r:id`、drawing の `xdr:cNvPr@id|name`、`oleObject@progId` / Embedded Control Data part の `classid` を組み合わせる経路があることを整理
  - chart sheet では `sheetPr@codeName` と `drawing` part までは確認できるが、Open XML docs 上で `controls` / `oleObjects` 相当の経路は未証明のため、worksheet と同列には扱わず保留とした
  - 抽出ツール artifact と将来 manifest は consumer format 候補と位置付け、最小経路は「workbook package probe -> sidecar JSON」とする方針を [docs/process/worksheet-chart-control-metadata-source-poc.md](docs/process/worksheet-chart-control-metadata-source-poc.md) に記録

- [x] Worksheet / Chart control identity source の整理
  - `OLEObject.Object` 後段型付けと `Sheet1.CommandButton1` 支援は、どちらも `shape name` / `code name` / `ProgID` を結び付けた control inventory を必要とすることを整理
  - Office VBA の `Using ActiveX Controls on Sheets`、`OLEObjects object`、`OLEObject.progID` を確認し、collection access は shape name、event / direct access は code name、control type 判定には `ProgID` が使えることを明文化
  - 現行リポジトリの静的入力 `.bas` / `.cls` / `.frm` / `.frx` だけでは worksheet / chart sheet 上の control inventory を復元できないため、`.Object` と control code name は当面未解決のまま維持する方針を [docs/process/worksheet-chart-control-identity-feasibility.md](docs/process/worksheet-chart-control-identity-feasibility.md) に記録

- [x] Worksheet / Chart OLEObjects.Item の型解決
  - `scripts/lib/supplementalReferenceConfig.mjs` と `scripts/generate-mslearn-vba-reference.mjs` に既存 owner member override を追加し、`OLEObjects.Item` を raw doc の `Object` ではなく `typeName: "OLEObjects"` として正規化した
  - `resources/reference/mslearn-vba-reference.json` を再生成し、`Sheet1.OLEObjects.Item(1)` / `Item("CheckBox1")` / `Item(i + 1)` / `Chart1.OLEObjects.Item(1)` から `OLEObject` member へ到達できるようにした
  - `Item(GetIndex())` は collection のまま維持し、script / server / extension テストで `Item(...)` の single/literal/function 境界を固定した

- [x] Worksheet / Chart OLEObjects root の最小プロトタイプ
  - `packages/core/src/reference/builtinReference.ts` の indexed collection owner map に `OLEObjects -> OLEObject` を追加し、`Sheet1.OLEObjects(1)` / `Sheet1.OLEObjects("CheckBox1")` / `Sheet1.OLEObjects(i + 1)` / `Chart1.OLEObjects(1)` から `OLEObject` member へ到達できるようにした
  - server / extension に専用 fixture と回帰テストを追加し、completion / hover / signature help で `OLEObject` method・property が出ることを確認した
  - `Sheet1.OLEObjects(GetIndex())` は collection のまま維持し、`.Object` の先は引き続き保守動作のまま据え置いた

- [x] DialogFrame 補助参照の最小プロトタイプ
  - `scripts/lib/supplementalReferenceConfig.mjs` に `DialogSheet.DialogFrame` の補助 property と `DialogFrame` owner の allow list を追加し、`resources/reference/mslearn-vba-reference.json` を再生成
  - `DialogSheets(1).DialogFrame` / `Application.DialogSheets(1).DialogFrame` / `ActiveWorkbook.DialogSheets(1).DialogFrame` / `ThisWorkbook.DialogSheets(1).DialogFrame` から `Caption` / `Text` / `Select` の completion / hover / signature help に到達できるようにした
  - `DialogSheets(Array(...))` は collection のまま維持し、`_Dummy*` / legacy member 混入防止と `DialogFrame` owner の型・署名監査を `scripts/test` / server / extension テストで固定した
  - CodeRabbit 指摘を受け、extension の negative signature help helper を空配列対応に修正し、supplemental property の `typeName` 欠落時は生成を fail-fast させた

- [x] DialogSheet control collection の single-selector 正規化
  - `DialogSheet.Buttons` / `CheckBoxes` / `OptionButtons` と各 collection / item owner を補助参照へ追加し、`resources/reference/mslearn-vba-reference.json` を再生成
  - 数値・文字列リテラル selector と `.Item(<literal>)` だけ `Button` / `CheckBox` / `OptionButton` へ落とし、式 selector や `Array(...)` selector は collection owner のまま維持する marker ルールを core / server に追加
  - collection owner は `Count` / `Item`、item owner は `Caption` / `Name` / `OnAction` / `Text` / `Select` と `Value` の最小構成に絞り、`scripts/test` / server / extension テストで `_Dummy*` 除外、type override、completion / hover / signature help / semantic token を固定した

- [x] Worksheet / Chart control collection への横展開方針整理
  - Office VBA の `Using ActiveX Controls on Sheets`、`Worksheet.OLEObjects`、`Chart.OLEObjects` を確認し、worksheet / chart sheet の正本道線は `OLEObjects` / `Shapes` / control code name であることを整理した
  - `WorksheetClass.Buttons` / `CheckBoxes` / `OptionButtons` と `ChartClass.*` は interop 側の `Optional Index As Object -> As Object` 補助ソース候補と位置付け、`DialogSheet` の literal-only selector 正規化を直ちに横展開しない方針を文書化した
  - `docs/process/worksheet-chart-control-collection-feasibility.md` を正本にし、変更系 member 抑止と `OLEObjects` / control name 導線優先を次段の前提として固定した

- [x] Worksheet / Chart OLEObjects / Shapes / control name 導線整理
  - Office VBA の `Using ActiveX Controls on Sheets`、`Worksheet.OLEObjects`、`Chart.OLEObjects`、`OLEObject.Object`、`Worksheet.Shapes`、`Chart.Shapes` を確認し、entry point ごとの向き不向きを比較した
  - 最初の実装候補は `Worksheet.OLEObjects(Index)` / `Chart.OLEObjects(Index)` を `OLEObject` owner へ落とす最小プロトタイプとし、`.Object` の先、`Sheet1.CommandButton1`、`Shapes` root は後続へ送る方針を文書化した
  - `docs/process/worksheet-chart-control-entrypoint-feasibility.md` を正本にし、`OLEObjects` 優先、`.Object` 未解決維持、control code name inventory 不足、`Shapes` 広域性を次段の前提として固定した
  - 後続で `Worksheet / Chart OLEObjects root` と `OLEObjects.Item` は実装済みだが、`.Object` 後段型付け、`Sheet1.CommandButton1`、`Shapes` root 実装は未着手のまま維持した

- [x] DialogSheet control collection の補助参照可否整理
  - Microsoft Learn の `DialogSheet.Buttons` / `CheckBoxes` / `OptionButtons` / `DialogFrame` と `Button` / `CheckBox` / `OptionButton` / `DialogFrame` interface を確認し、owner ごとの導入難所を整理
  - `DialogFrame` は direct property で先行候補、他 3 系統は `Optional Index As Object -> As Object` のため single-selector / collection selector 分岐が必要という結論を文書化
  - `docs/process/dialogsheet-control-collection-feasibility.md` を正本にし、`_Dummy*` 除外、`memberTypeOverrides`、grouped selector 抑止を次段実装の前提として固定

- [x] DialogSheet common callable の Workbook / Application root 展開
  - `scripts/lib/supplementalReferenceConfig.mjs` に `Application.DialogSheets` / `Workbook.DialogSheets` の補助 property 設定を追加し、`resources/reference/mslearn-vba-reference.json` を再生成
  - `Application.DialogSheets(1)` / `ActiveWorkbook.DialogSheets(1)` から既存 `DialogSheet` common callable へ到達できるようにし、`Array(...)` selector は collection のまま維持した
  - `scripts/test` / server / extension に root 展開の監査と回帰テストを追加し、`DialogSheet1.` document module root は引き続き未公開のまま保守した

- [x] ドキュメント導線とレビュー履歴の整理
  - `docs/README.md`、`docs/process/README.md`、`docs/adr/README.md` を追加し、入口から要件、ADR、運用詳細へ辿る導線を整理
  - CodeRabbit 実ログを `docs/process/coderabbit-review-logs/2026-03.md` へ分割し、レビュー履歴の肥大化を抑制
  - `README.md`、`AGENTS.md`、関連運用文書の参照先を新しい入口へ寄せ、重複記載を削減

- [x] DialogSheet common callable の最小プロトタイプ
  - `scripts/lib/supplementalReferenceConfig.mjs` に `DialogSheet` interop allow list と `DialogSheets` collection clone を追加し、`resources/reference/mslearn-vba-reference.json` を再生成
  - `DialogSheets(1)` / `DialogSheets("Dialog1")` / `DialogSheets.Item(1)` では `DialogSheet` member へ到達し、`DialogSheets(Array(...))` は collection のまま維持する built-in 解決を追加
  - `dummy` / legacy member 混入防止に加え、allow list member の署名抽出失敗や `DialogSheets.Item` の `typeName` 欠落を監査するテストを追加し、`DialogSheet1.` document module の保守動作は維持した

- [x] DialogSheet interop 由来 member source の導入可否整理
  - Office VBA 概念記事、Office VBA API、.NET interop `DialogSheet` page を突き合わせ、全面自動導入は不採用、補助ソースとしての限定利用は可能という結論を整理
  - interop page は `Reserved for internal use.`、`_Dummy*`、`_SaveAs` のような legacy member を含むため、導入するなら allow list と skip rule を先に固定する必要がある
  - `docs/process/dialogsheet-interop-source-feasibility.md` に source inventory、制約、推奨方針、最小候補 member を記録

- [x] DialogSheet document module root の扱い整理
  - Microsoft Learn の `Refer to Sheets by Name` で `DialogSheets("Dialog1").Activate` を確認し、dialog sheet 自体は VBA から参照可能であることを再確認
  - ただし Office VBA 側に `DialogSheet` object page は無く、ローカル参照 JSON も owner を持たないため、現時点では built-in owner へ昇格しない方針を ADR へ記録
  - .NET interop `DialogSheet` page は `Reserved for internal use.` かつ `dummy` member を含むため、補助ソース化は別タスクへ分離
  - server テストでは `DialogSheet1.` が built-in completion / signature help / hover / semantic token を返さない保守動作を既に固定している

- [x] Chart document module root の到達性改善
  - `VB_PredeclaredId = True` かつ `VB_Base = 0{00020821-0000-0000-C000-000000000046}` の chart document module を `Chart` root として扱い、`Chart1.` から built-in member completion / signature help / hover / semantic token へ到達できるようにした
  - Microsoft Learn の `Chart.CodeName` / `Chart object` と Windows registry の `Excel.Chart` CLSID を根拠に、chart sheet code name を document root として扱う条件を固定した
  - `DialogSheet` は Office VBA の object page / 参照 JSON が不足しているため今回は保守動作を維持し、次候補へ分離した

- [x] Sheet document module alias の到達性改善
  - `VB_PredeclaredId = True` かつ `VB_Base = 0{00020820-0000-0000-C000-000000000046}` の worksheet document module だけを `Worksheet` root として扱い、`Sheet1.SaveAs` / `Sheet1.Evaluate` の built-in member 到達性を追加
  - `ThisWorkbook` 専用だった document module root 判定を一般化し、completion / signature help / hover / semantic token で共通に使うよう整理
  - worksheet 以外の predeclared class module は保守動作のまま維持し、誤って `Worksheet` member を出さない回帰テストを追加

- [x] Worksheet callable の署名昇格
  - `ActiveSheet` を型付けせず、`Worksheets(1)` / `ActiveWorkbook.Worksheets(1)` のような indexed collection access から `Worksheet` member へ到達できるようにした
  - 現行 Microsoft Learn の `Worksheet.Evaluate` / `Worksheet.SaveAs` / `Worksheet.ExportAsFixedFormat` を署名抽出対象へ追加し、参照 JSON を再生成した
  - server / extension テストで indexed collection access 経由の completion / signature help / semantic token を回帰確認した
  - `Worksheets("A(1)")` と `Worksheets(i + 1)` は単一 `Worksheet` として扱い、`Worksheets(Array(...))` や `ActiveWorkbook.Worksheets(GetIndex())` は collection のまま維持する保守動作を回帰固定した

- [x] 組み込みメンバー署名データの拡張（第8弾）
  - 現行 Microsoft Learn の `Workbook.SaveAs` / `Workbook.Close` / `Workbook.ExportAsFixedFormat` を署名抽出対象へ追加し、参照 JSON を再生成
  - `Sub` 相当の `Workbook` callable は生成データへ `returnType: "Void"` を保持しつつ、表示ラベルは従来どおり `As Void` を出さない形へ補正した
  - `ActiveWorkbook` / `ThisWorkbook` 経由の signature help と hover を server / extension テストで回帰確認した
  - `Worksheet.Evaluate` / `Worksheet.SaveAs` は現行 root 到達性を踏まえて次候補へ残し、Workbook 側を優先した

- [x] Application / Workbook / Worksheet 系 inventory と workbook root 解決
  - 現行 Microsoft Learn の `Application` / `Workbook` / `Worksheet` object page とローカル参照 JSON を照合し、この 3 owner では object page 由来の未掲載 member が無いことを確認した
  - `ActiveWorkbook` / `ThisWorkbook` から `Workbook` member を引けるようにし、`Application.ActiveCell` のような alias property chain でも既存 `typeName` を継承できるようにした
  - server / extension テストで `ActiveWorkbook` / `ThisWorkbook` completion、`Application.ActiveCell.Address` signature、`ThisWorkbook.SaveAs` hover、`ThisWorkbook.SaveAs` / `Application.ActiveCell.Address` semantic token を回帰確認した

- [x] Microsoft Learn 監視対象 owner の拡張
  - `scripts/lib/referenceSignatureConfig.mjs` の watch list に `Range.HasSpill` / `SavedAsArray` / `SpillParent` を追加し、`WorksheetFunction` 以外も未掲載監視できるようにした
  - `docs/process/mslearn-signature-regeneration.md` に現在の watch list と owner 選定基準を追加し、`Range` 動的配列メンバーの更新導線を明記した
  - `WorksheetFunction` だけを前提にしない形へ手順書を更新し、次回の owner 拡張候補整理につなげた

- [x] Microsoft Learn 監視対象メンバーの自動検知見直し
  - `scripts/lib/referenceSignatureConfig.mjs` に owner 単位の `signatureMissingMemberWatchList` を追加し、未掲載監視を共有設定化
  - `scripts/test/mslearnReferenceAudit.test.mjs` を watch list ベースの監視と allow list 重複検知へ更新
  - `docs/process/mslearn-signature-regeneration.md` を watch list から allow list への移行手順に合わせて更新

- [x] Microsoft Learn 追加メンバーの再生成観点整理
  - `docs/process/mslearn-signature-regeneration.md` を追加し、allow list、再生成、built-in index、server / extension テスト、レビュー記録までの更新箇所を整理
  - `scripts/test/mslearnReferenceAudit.test.mjs` の監視失敗メッセージから手順書へ辿れるように修正
  - `AGENTS.md` に手順書への導線を追加

- [x] レビュー判断ルールの更新
  - PR 前自己レビューと CodeRabbit が同じ論点を指摘した場合は、原則として修正する運用へ変更
  - `required` / `optional` のような運用時挙動については、互換性、既存テスト、誤案内防止を基準に判断する方針を正本へ明記
  - 正本の `docs/process/coderabbit-review.md` に集約し、`sub-agent` / `AGENTS` / `TASKS` へ反映

- [x] 既存署名メタデータの横断点検
  - `WorksheetFunction.Max` / `Min` の第1引数 metadata 欠落と `Arg30` の required 誤判定を生成スクリプト側で修正し、参照 JSON を再生成
  - `WorksheetFunction` / `Range` の既存署名について、型・説明・必須/省略可能・戻り値型の欠落監査を `scripts/test` に追加
  - 現行 Microsoft Learn スナップショットでは `WorksheetFunction` に `XLookup` / `XMATCH` が未掲載であることを回帰確認

- [x] 組み込みメンバー署名データの拡張（第7弾）
  - 現行 Microsoft Learn を再確認し、`WorksheetFunction` には `XLookup` / `XMATCH` が未掲載のままであることを確認
  - `Range.Address` / `Range.AddressLocal` の署名を取り込み、`ActiveCell` / `Cells` のような Range 系組み込みルートからも解決できるように修正
  - `Address` 系の optional 引数メタデータと `XlReferenceStyle` 型情報を server / extension テストで回帰確認

- [x] 組み込みメンバー署名データの拡張（第6弾）
  - `WorksheetFunction.Choose` / `WorksheetFunction.Transpose` を Microsoft Learn 由来の署名抽出対象へ追加し、参照 JSON を再生成
  - `Choose` の可変長必須引数と `Transpose` の単一必須引数を server / extension テストで回帰確認
  - `XLookup` / `XMATCH` / `Address` は現行 Learn JSON で確認できなかったため、次候補で再整理する

- [x] CodeRabbit レビュー記録運用の追加
  - CodeRabbit の月次記録テンプレートを整備し、PR ごとのレビュー記録を開始した
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

- [x] Worksheet control metadata sidecar lookup の最小実装
  - `scripts/lib/worksheetControlMetadataSidecar.mjs` と `--format sidecar` / `--bundle-root` を `scripts/probe-workbook-control-metadata.mjs` へ追加し、probe 出力から `<bundle-root>/.vba/worksheet-control-metadata.json` を生成・配置できるようにした
  - `packages/core/src/reference/worksheetControlMetadataSidecar.ts` を追加し、nearest ancestor lookup、workspace root での打ち切り、schema v1 validation、`status: "unsupported"` owner の切り分け helper を実装した
  - `packages/server/src/lsp/documentService.ts` へ sidecar の read-only cache と log を追加し、各 `DocumentState` に読み込み結果を保持するようにした
  - scripts / core / server の回帰テストを追加し、初回は user-facing 型解決へ直結させず read-only state までで固定した

- [x] Worksheet control metadata sidecar を `OLEObject.Object` 型解決へ接続する
  - `Sheet1.OLEObjects("ShapeName").Object` と `Sheet1.OLEObjects.Item("ShapeName").Object` の string literal selector にだけ sidecar を適用し、`shapeName -> controlType` で `CheckBox` などの control owner へ進めるようにした
  - 数値 selector、dynamic selector、`Chart1` の unsupported owner、`ActiveSheet` root は従来どおり保守的に未解決のまま維持し、named worksheet selector だけを user-facing に昇格した
  - completion / hover / signature help / semantic token を server / extension test で回帰固定し、fixture sidecar も追加した

- [x] Worksheet control metadata sidecar を `Sheet1.ControlCodeName` 解決へ接続する
  - worksheet document module root の direct access にだけ sidecar を適用し、`sheetCodeName + codeName -> controlType` で `Sheet1.chkFinished.Value` / `Select` のような control code name 導線を解決するようにした
  - workbook / standard module からの非 document-module access、`ActiveSheet` root、supported/unsupported を問わない chartsheet owner、sidecar 未検出、`shape name != code name` の direct access は従来どおり保守的に未解決のまま維持した
  - server / extension に専用 fixture と回帰テストを追加し、`OLEObject.Object` 既存導線と衝突しないこと、semantic token / hover / signature help まで control owner へ進むことを固定した

- [x] Worksheet / Chart Shapes root の control 導線を整理する
  - `Worksheet.Shapes(Index)` / `Chart.Shapes(Index)` と `Shapes.Item(Index)` を generic `Shape` owner へ正規化し、`Name` や `OLEFormat` などの `Shape` surface までは user-facing に出すようにした
  - `Shape.OLEFormat.Object` の先は generic `Object` のまま維持し、worksheet control metadata sidecar があっても `CheckBox.Value` のような control-specific owner へは昇格させないことを server / extension test で固定した
  - `msoOLEControlObject`、`shape name != code name`、embedded / linked OLE object 混在時の扱いは正本 [docs/process/worksheet-chart-shapes-root-feasibility.md](docs/process/worksheet-chart-shapes-root-feasibility.md) に整理した

- [x] Shape.OLEFormat.Object の control owner 昇格条件を整理する
  - 昇格候補を `worksheet document module root + shape name string literal + sidecar 一致` に限定し、`Shape.Type = msoOLEControlObject` は runtime 条件、sidecar provenance を静的 resolver 条件として使う方針を正本 [docs/process/shape-oleformat-object-promotion-feasibility.md](docs/process/shape-oleformat-object-promotion-feasibility.md) に整理した
  - `Shapes(1)` / `Shapes.Item(1)`、dynamic selector、`Chart1` root、`ShapeRange` / grouped selector、code name 導線を昇格対象から外す判断を文書化し、server / extension test で `Shape.OLEFormat.Object` 非昇格の境界を補強した

- [x] Shape.OLEFormat.Object を worksheet literal shapeName selector に限定接続する
  - `Sheet1.Shapes("CheckBox1").OLEFormat.Object` と `Sheet1.Shapes.Item("CheckBox1").OLEFormat.Object` だけを sidecar 連携で control owner へ進め、`Value` / `Select` の completion / hover / signature help / semantic token が出るようにした
  - numeric / dynamic / chart / `ShapeRange` の負例と、`OLEObject.Object` / `Sheet1.ControlCodeName` 既存導線との非衝突を server / extension test で固定した

- [x] Shape.OLEFormat.Object の explicit sheet-name root 拡張可否を整理する
  - `Worksheets("Sheet1")` 系 root の join key は `sheetCodeName` ではなく `sheetName` を使うべきこと、`Sheet1` alias / control code name 導線とは別 key として扱うべきことを正本 [docs/process/shape-oleformat-object-explicit-sheet-root-feasibility.md](docs/process/shape-oleformat-object-explicit-sheet-root-feasibility.md) に整理した
  - unqualified `Worksheets("Sheet1")` と `ActiveWorkbook.Worksheets("Sheet1")` は active workbook 依存のため現段階では不採用、最初の候補は `ThisWorkbook.Worksheets("Sheet1")` に限定する方針を固定した
  - `OLEObject.Object` と `Shape.OLEFormat.Object` は `workbook root identity + sheetName + shapeName` lookup helper を将来共有できるが、`ActiveSheet` / chartsheet / `ShapeRange` は引き続き除外境界として維持する

- [x] Shape.OLEFormat.Object を ThisWorkbook.Worksheets("SheetName") root に限定接続する
  - `ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object` と `.Item("CheckBox1")` を sidecar 連携で control owner へ進め、completion / hover / signature help / semantic token を server / extension test で回帰固定した
  - shared helper により `ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object` と `.Item("CheckBox1").Object` も同じ workbook-qualified root から control owner へ進み、server / extension test で固定した
  - `sheetCodeName` と `sheetName` が一致しない fixture（`sheetCodeName = "Sheet1"`, `sheetName = "Sheet One"`）を導入し、`Worksheets("Sheet One")` のみが解決される境界を server / extension test で固定した
  - `ThisWorkbook` 起点の workbook root identity を `Worksheets("SheetName")` 連鎖でも保持して resolver へ伝播させ、generic `Worksheet` owner に降りた後も current bundle の sidecar を選べるようにした
  - `sheetName + shapeName` lookup は `OLEObject.Object` / `Shape.OLEFormat.Object` の共通 helper で扱える形に整理しつつ、`ActiveWorkbook` / unqualified `Worksheets` / `ThisWorkbook.Worksheets(1)` / `ActiveSheet` / chartsheet / `ShapeRange` は引き続き負例で固定した

- [x] workbook-qualified 以外の explicit sheet-name root 展開条件を整理する
  - `docs/process/explicit-sheet-name-broad-root-feasibility.md` を追加し、`ActiveWorkbook.Worksheets("Sheet1")` と unqualified `Worksheets("Sheet1")` は active workbook 依存のため current bundle の sidecar に静的接続しない方針を一次情報ベースで整理した
  - `ThisWorkbook.Worksheets("Sheet1")` だけを user-facing に維持し、broad root の境界は `OLEObject.Object` と `Shape.OLEFormat.Object` の両方でそろえる方針を固定した
  - ADR [0005](docs/adr/0005-explicit-sheet-name-root-policy.md) の status を `Accepted` に上げ、broad root 再評価は workbook binding / manifest のような明示的 workbook 同一性が入った場合だけと整理した

- [x] current bundle と target workbook を結ぶ明示バインディング条件を整理する
  - ADR [0006](docs/adr/0006-workbook-binding-policy.md) を追加し、broad root 再評価の前提となる workbook identity binding は専用 artifact `<bundle-root>/.vba/workbook-binding.json` で扱う方針を固定した
  - `docs/process/workbook-binding-manifest-feasibility.md` を追加し、manifest / config / workbook package mode の比較、`Workbook.FullName` を primary key にする理由、unsaved workbook と add-in workbook を broad root 対象外にする理由を整理した
  - `worksheet-control-metadata.json` には binding 情報を混ぜず、sidecar は control inventory、binding manifest は runtime workbook identity と責務分離する方針を文書化した

- [x] active workbook identity provider の契約を整理する
  - ADR [0007](docs/adr/0007-active-workbook-identity-provider-contract.md) を追加し、host / extension / server で共有する `ActiveWorkbookIdentitySnapshot`、custom LSP notification `vba/activeWorkbookIdentity`、resolver gating 条件を固定した
  - [docs/process/active-workbook-identity-provider-contract.md](docs/process/active-workbook-identity-provider-contract.md) を追加し、`Application.ActiveWorkbook` / `ActiveProtectedViewWindow` / `ProtectedViewWindow.Workbook` / `Workbook.Saved` / `Workbook.IsAddin` を根拠に、`available` / `unavailable` / `protected-view` / `unsupported` state と log 方針を整理した
  - [docs/process/workbook-binding-manifest-feasibility.md](docs/process/workbook-binding-manifest-feasibility.md) と [docs/process/explicit-sheet-name-broad-root-feasibility.md](docs/process/explicit-sheet-name-broad-root-feasibility.md) を更新し、manifest policy と runtime provider contract の責務分離を明確化した

- [x] active workbook identity snapshot transport の最小実装を整理する
  - `packages/core` に `ActiveWorkbookIdentitySnapshot` / `workbook-binding.json` の schema helper、validation、nearest ancestor lookup、`normalizeWorkbookFullNameForComparison` を追加し、runtime snapshot と manifest matching rule を共有できるようにした
  - `packages/server` に custom notification `vba/activeWorkbookIdentity`、snapshot cache、`binding-missing` / `binding-disabled` / `match` / `mismatch` / `invalid-payload` log を read-only で追加し、resolver へはまだ接続しない段階で状態観測できるようにした
  - `packages/extension` は host bridge 実装前に placeholder snapshot を送らず、runtime state 未初期化のまま待機する。notification 経路自体は server 側へ追加済みで、host 実装時に `vba/activeWorkbookIdentity` をそのまま流せるようにした

- [x] workbook-bound broad root gating の最小接続を追加する
  - `available` snapshot と `workbook-binding.json` の match がそろったときだけ、`ActiveWorkbook.Worksheets("SheetName")` 系 broad root を current bundle sidecar lookup の候補へ開く
  - 最初の接続対象は既存の `OLEObject.Object` / `Shape.OLEFormat.Object` と同じ worksheet control owner 導線に限定し、resolver 条件と負例を server / extension test で固定する
  - `packages/extension` に test mode 限定の `vba.test.setActiveWorkbookIdentitySnapshot` コマンドを追加し、fixture manifest と組み合わせて broad root gating の E2E を固定した

- [x] unqualified worksheet broad root gating の可否を整理する
  - ADR [0005](docs/adr/0005-explicit-sheet-name-root-policy.md) と [docs/process/explicit-sheet-name-broad-root-feasibility.md](docs/process/explicit-sheet-name-broad-root-feasibility.md) を更新し、unqualified `Worksheets("SheetName")` / `Application.Worksheets("SheetName")` は Office VBA 上で active workbook family とみなし、`ActiveWorkbook.Worksheets("SheetName")` と同じ runtime gating 条件でだけ将来 user-facing にできる方針を固定した
  - broad root family の v1 対象構文を `Worksheets("literal")` / `Application.Worksheets("literal")` に限定し、`Sheets` / `ActiveSheet` / numeric / dynamic / grouped selector と shadow case は sidecar lookup の対象外に残す境界を固定した
  - `OLEObject.Object` / `Shape.OLEFormat.Object` は同じ PR・同じ条件で開閉すること、manifest mismatch / unavailable / shadow case を受け入れ条件に含めることを docs に反映した

- [x] unqualified worksheet broad root の最小接続を追加する
  - `packages/server` の broad root resolver を拡張し、`available` snapshot と `workbook-binding.json` の match がそろったときだけ、unqualified `Worksheets("SheetName")` と `Application.Worksheets("SheetName")` を synthetic `Worksheet` root として sidecar lookup へ流すようにした
  - 最初の user-facing 対象は `OLEObject.Object` / `Shape.OLEFormat.Object` に限定し、`Sheets` / `ActiveSheet` / numeric / dynamic selector、`Worksheets` shadow、`Application` shadow は server / extension test で負例のまま固定した
  - broad-root 専用 fixture を extension test に追加し、snapshot 未一致と manifest mismatch では閉じたまま、match 時にだけ completion / hover / signature help が開くことを `npm run lint` / `npm test` / `npm run package` で確認した

- [x] unqualified worksheet broad root の Item selector 回帰を拡張する
  - `Worksheets("SheetName").OLEObjects.Item("ShapeName").Object`、`Application.Worksheets("SheetName").OLEObjects.Item("ShapeName").Object`、`Worksheets("SheetName").Shapes.Item("ShapeName").OLEFormat.Object`、`Application.Worksheets("SheetName").Shapes.Item("ShapeName").OLEFormat.Object` を broad root family の正例として server / extension test へ追加した
  - `Worksheets.Item("SheetName")` / `Application.Worksheets.Item("SheetName")` のような root `.Item` 形式は v1 非対象として [docs/adr/0005-explicit-sheet-name-root-policy.md](docs/adr/0005-explicit-sheet-name-root-policy.md) と [docs/process/explicit-sheet-name-broad-root-feasibility.md](docs/process/explicit-sheet-name-broad-root-feasibility.md) に明記し、snapshot match 中でも開かない負例を固定した
  - broad root family の `.Item` 回帰を追加したうえで、`npm run lint` / `npm test` / `npm run package` を通して direct call form の既存境界と root `.Item` 非対象境界が崩れていないことを確認した
  - この「v1 非対象」判断は後続の `unqualified worksheet broad root の root .Item("SheetName") 形式を再評価` で superseded 済み

- [x] workbook-qualified worksheet root の `.Item("SheetName")` 形式を整理する
  - `ThisWorkbook.Worksheets.Item("SheetName")` / `ActiveWorkbook.Worksheets.Item("SheetName")` は `Worksheets.Item` の既定メンバー規則に従って direct call form と同じ sheetName selector として扱い、`packages/server` の workbook-qualified root sidecar lookup を共通 helper へ寄せた
  - `OLEObject.Object` / `Shape.OLEFormat.Object` の両経路で `ThisWorkbook.Worksheets.Item("SheetName")` は静的 current bundle root、`ActiveWorkbook.Worksheets.Item("SheetName")` は manifest + snapshot gating 下だけ user-facing に開くことを server / extension test で固定した
  - `ThisWorkbook.Worksheets.Item("Sheet1")` / `ActiveWorkbook.Worksheets.Item("Sheet1")` の codeName 指定と、`ThisWorkbook.Worksheets.Item(1)` / `ActiveWorkbook.Worksheets.Item(1)` の numeric selector は、match 済み snapshot 下でも従来どおり control owner に昇格しない負例のまま固定した

- [x] `Application.ThisWorkbook` / `Application.ActiveWorkbook` root を整理する
  - [docs/process/application-workbook-root-feasibility.md](docs/process/application-workbook-root-feasibility.md) を追加し、`Application.ThisWorkbook.Worksheets("SheetName")` / `.Item("SheetName")` は `ThisWorkbook` direct root と同じ static current-bundle family、`Application.ActiveWorkbook.Worksheets("SheetName")` / `.Item("SheetName")` は `ActiveWorkbook` direct root と同じ manifest + snapshot gating family として扱う方針を一次情報ベースで整理した
  - ADR [0005](docs/adr/0005-explicit-sheet-name-root-policy.md) を更新し、`Application` qualifier を挟んでも workbook root identity の意味は変えないこと、ただし built-in `Application` として解決できる場合にだけ sidecar lookup を有効にする shadow 境界を固定した
  - [docs/process/explicit-sheet-name-broad-root-feasibility.md](docs/process/explicit-sheet-name-broad-root-feasibility.md)、[docs/process/workbook-binding-manifest-feasibility.md](docs/process/workbook-binding-manifest-feasibility.md)、[docs/process/README.md](docs/process/README.md) を更新し、既存 broad-root / binding メモとの参照導線をそろえた

- [x] `Application.ThisWorkbook` / `Application.ActiveWorkbook` root の最小接続を追加する
  - `packages/server/src/lsp/documentService.ts` の broad-root resolver を拡張し、built-in `Application` qualifier の直後に `ThisWorkbook` / `ActiveWorkbook` が続く場合は、それぞれ既存 `ThisWorkbook` direct root と `ActiveWorkbook` broad-root family へ寄せて `OLEObject.Object` / `Shape.OLEFormat.Object` の sidecar lookup へ流すようにした
  - `packages/server/test/documentService.test.js`、`packages/extension/test/suite/index.ts`、`packages/extension/test/fixtures/ApplicationWorkbookRootBuiltIn.bas` を追加・更新し、`Application.ThisWorkbook.Worksheets("SheetName")` / `.Item("SheetName")` は static current-bundle family、`Application.ActiveWorkbook.Worksheets("SheetName")` / `.Item("SheetName")` は manifest + snapshot match 下だけ開くこと、`Application` shadow、codeName selector、numeric selector は従来どおり負例のままであることを completion / hover / signature help / semantic token まで固定した
  - CodeRabbit 指摘対応として、indexed access の `ActiveWorkbook("...")` を direct workbook root と誤認しない `accessKind === "none"` ガードを `packages/server/src/lsp/documentService.ts` に追加し、server 側に専用の負例テストを追加した
  - [docs/process/application-workbook-root-feasibility.md](docs/process/application-workbook-root-feasibility.md) と [docs/process/workbook-binding-manifest-feasibility.md](docs/process/workbook-binding-manifest-feasibility.md) を更新し、方針整理で止まっていた `Application.ThisWorkbook` / `Application.ActiveWorkbook` root が user-facing 実装済みになったことを反映した
  - 検証として `npm run build`、`npm test --workspace @vba/server`、`npm run lint`、`npm run package` を通した。`npm run test --workspace vba-extension` はこの Codex セッション自体が Code を使用中のため CLI 実行が弾かれ、別インスタンスが無い環境での再確認が必要

- [x] workbook root family の resolver / test matrix を整理する
  - `packages/server/src/lsp/documentService.ts` に `getWorkbookRootFamilyBuiltinContext()`、`resolveWorkbookRootFamilyPath()`、`getWorkbookDocumentBuiltinContext()` を追加し、`ThisWorkbook` / `ActiveWorkbook` / `Application.ThisWorkbook` / `Application.ActiveWorkbook` / `Worksheets` / `Application.Worksheets` の workbook root family 判定を shared helper へ寄せた
  - workbook root family にだけ関係する type-only root 解決は built-in family 判定の手前で無視するようにし、module-level `Type Application` が `Application.ThisWorkbook` / `Application.ActiveWorkbook` を塞がないことを server / extension の matrix で固定した
  - `Application.Caller` / `Application.Range("A1")` は workbook root family に昇格しない負例を server / extension に追加し、今後 `Application.*` selector 解析を広げても broad root family へ誤接続しない境界を固定した
  - `packages/server/test/documentService.test.js` に workbook root 用の completion / hover / signature helper を追加し、`Application.ThisWorkbook` / `Application.ActiveWorkbook` と shadow case の matrix 実行を helper 経由へ寄せた
  - `packages/extension/test/suite/index.ts` に workbook root 用の async helper と診断しやすい completion summary を追加し、`WorksheetBroadRootBuiltIn.bas` と `ApplicationWorkbookRootBuiltIn.bas` の current-bundle / broad-root matrix を配列駆動 helper で回すように整理した
  - `packages/extension/test/fixtures/ApplicationWorkbookRootBuiltIn.bas` を server 側の matrix と揃え、negative hover / signature と shadow 用 token の不足を補った
  - 検証として `npm run build --workspace @vba/server`、`npm test --workspace @vba/server`、`npx tsc -p packages/extension/tsconfig.test.json`、`npm run build`、`npm run test --workspace vba-extension` を通した

- [x] workbook root family の semantic token / fixture helper を追加整理する
  - `packages/server/test/documentService.test.js` と `packages/extension/test/suite/index.ts` に workbook root family 用の semantic token helper を追加し、line number 固定ではなく anchor token ベースで `Value` / `Select` を検証する形へ寄せた
  - workbook-qualified root item と `Application.ThisWorkbook` / `Application.ActiveWorkbook` matrix の semantic token assert を token 名ベースへ置き換え、fixture 行追加だけで壊れにくい構成にした
  - `ApplicationWorkbookRootBuiltIn.bas` と `WorksheetBroadRootBuiltIn.bas` へ section comment を追加し、current-bundle family / broad-root family / non-target root の塊を読み分けやすくした
  - 挙動変更なしの整理として `npm run lint`、`npm test`、`npm run package` を通した

- [x] workbook root family の matrix case table 共通化方針を整理する
  - [docs/process/workbook-root-family-case-table-policy.md](docs/process/workbook-root-family-case-table-policy.md) を追加し、shared 化対象は helper ではなく repo root `test-support/` 配下の canonical case spec に限定し、server / extension は package-local adapter を維持する方針を固定した
  - server の CommonJS / 同期 assertion と extension の TypeScript / 非同期 assertion の差から、`resolveJsonModule` や copy step を増やさずに両方から読める `*.cjs` ベースを正本候補とする判断を整理した
  - `occurrenceIndex` を shared spec 側で明示管理し、shadow / duplicate anchor を package-local message と切り分ける方針を固定した

- [x] workbook root family の shared case spec を最小抽出する
  - `test-support/workbookRootFamilyCaseTables.cjs` を追加し、`WorksheetBroadRootBuiltIn.bas` / `ApplicationWorkbookRootBuiltIn.bas` の completion anchor と application workbook root family の semantic anchor を canonical spec として切り出した
  - `packages/server/test/documentService.test.js` では shared spec を fixture に存在する anchor だけへ絞る filter を通し、broad root / application workbook root の completion・semantic matrix を package-local helper へ変換する薄い adapter へ寄せた
  - `packages/extension/test/suite/index.ts` では `test-support/` の CJS 正本を runtime 解決し、`CompletionItem.detail` / blocked label / async wait 条件 / failure message は local に残したまま shared spec から tuple を組み立てる形へ整理した
  - 検証として `npm run test --workspace @vba/server`、`npx tsc -p packages/extension/tsconfig.test.json`、`npm run lint`、`npm run package` を通した。`npm run test --workspace vba-extension` はこの Codex セッションで Code を使用中のため CLI 実行が弾かれた

- [x] workbook root family の hover / signature anchor shared 化を再評価する
  - `test-support/workbookRootFamilyCaseTables.cjs` に `hover` / `signature` の shared spec を追加し、`WorksheetBroadRootBuiltIn.bas` と `ApplicationWorkbookRootBuiltIn.bas` の positive / non-shadow negative anchor を canonical spec へ切り出した
  - `packages/server/test/documentService.test.js` と `packages/extension/test/suite/index.ts` を更新し、hover / signature も shared spec から package-local message と assertion shape を組み立てる adapter へ寄せた
  - shadow hover / signature は extension と server で `occurrenceIndex` が異なるため local に残し、境界を [docs/process/workbook-root-family-case-table-policy.md](docs/process/workbook-root-family-case-table-policy.md) に明記した
  - 検証として `npm run test --workspace @vba/server`、`npx tsc -p packages/extension/tsconfig.test.json`、`npm run lint`、`npm run package` を通した

- [x] workbook root family の shadow hover / signature occurrence 差分を整理する
  - [docs/process/workbook-root-family-case-table-policy.md](docs/process/workbook-root-family-case-table-policy.md) を更新し、extension の shadow hover は duplicate `.Value` anchor のため `occurrenceIndex = 1`、extension の shadow signature と server 側 shadow fixture は `occurrenceIndex = 0` で足りる実態を文書化した
  - v1 では shared spec schema に per-scope / per-kind occurrence override を追加せず、shadow hover / signature は package-local のまま残す判断を固定した
  - shared 化を再検討する条件を「shadow fixture topology を寄せられたとき」または「同種の per-scope occurrence 差分が別 family でも増えたとき」として整理した

- [x] workbook root family の shadow fixture topology をそろえる条件を整理する
  - [docs/process/workbook-root-shadow-fixture-topology-feasibility.md](docs/process/workbook-root-shadow-fixture-topology-feasibility.md) を追加し、extension は `Demo()` + `ShadowedApplication()` の mixed fixture、server は shadow 専用 inline fixture という topology 差を比較した
  - `per-scope occurrence override` を shared spec schema に追加する案は不採用とし、shadow shared 化を再開するなら extension 側の shadow section を専用 fixture へ分離する案を第一候補として整理した
  - [docs/process/README.md](docs/process/README.md) に導線を追加し、fixture topology を寄せる前に切り分ける条件を明文化した

- [x] workbook root family の shadow 専用 fixture 分離 PoC を整理する
  - [docs/process/workbook-root-shadow-fixture-split-poc.md](docs/process/workbook-root-shadow-fixture-split-poc.md) を追加し、extension 側で `ShadowedApplication()` を別 fixture へ切り出した場合の document 数増加、helper 影響、shared spec への波及範囲を整理した
  - server 側は inline shadow text を維持し、canonical shadow text source の導入は PoC 完了後に別判断へ切り分ける方針を固定した

- [x] workbook root family の shadow 専用 fixture 分離を最小実装する
  - extension の `ApplicationWorkbookRootBuiltIn.bas` から `ShadowedApplication()` を [ApplicationWorkbookRootShadowed.bas](packages/extension/test/fixtures/ApplicationWorkbookRootShadowed.bas) へ切り出し、shadow 用 document を別に開く構成へ変更した
  - shadow semantic negative を `state: "shadowed"` へ寄せ、shadow hover / signature も server / extension で同じ shared entry を参照する形へ寄せた
  - `npm run lint`、`npm test`、`npm run package` で回帰確認する

- [x] workbook root family の shadow text source canonical 化要否を整理する
  - [docs/process/workbook-root-shadow-text-source-canonicalization-feasibility.md](docs/process/workbook-root-shadow-text-source-canonicalization-feasibility.md) を追加し、canonical 化対象は full text ではなく shared anchor spec に留める判断を整理した
  - dedicated shadow fixture と server inline shadow text の二重管理は、現時点では drift や review コストの主因ではないと整理し、file 正本 / generator 正本は再評価トリガー発火時のみ比較する方針を固定した

- [x] workbook root family と同種の dedicated shadow fixture 分離が別 family でも必要かを観測する
  - [docs/process/shadow-fixture-split-cross-family-observation.md](docs/process/shadow-fixture-split-cross-family-observation.md) を追加し、現時点で第 2 family は無く、dedicated shadow fixture 分離は workbook root family 固有の対処に留まると整理した
  - `worksheet broad root family` は最も近い候補だが、まだ server-only shadow case であり、extension + server の shared shadow matrix には達していないと整理した

## メモ

- `docs/requirements/000-overview.md` にはユーザー管理の差分があるため、自動コミットに含めない
