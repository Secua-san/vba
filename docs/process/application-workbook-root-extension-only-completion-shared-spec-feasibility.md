# application workbook root family の extension-only completion entry を shared spec に残す境界

## 結論

- 現時点では `applicationWorkbookRoot.completion.negative` に残っている extension-only 3 entry を shared spec に残す。
- `packages/extension/test/suite/index.ts` 側の local table へ戻さず、引き続き [test-support/workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) を family canonical anchor source として扱う。
- `CompletionItem.detail` fragment と blocked label は package-local adapter に残し、shared spec へは入れない。

## 目的

前段で `applicationWorkbookRoot` の extension-only `hover` / `signature` / `semantic` entry は shared spec に残す判断を固定した。  
このメモでは、その残件である extension-only `completion.negative` 3 entry も同じく shared spec に残してよいかを整理する。

## 現状

### 対象 entry

`applicationWorkbookRoot.completion.negative` のうち、scope が `["extension"]` だけの entry は次の 3 本である。

- `Application.ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.`
  - `reason: "numeric-selector"`
  - `state: "static"`
- `Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.`
  - `reason: "code-name-selector"`
  - `state: "static"`
- `Application.ActiveWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.`
  - `reason: "numeric-selector"`
  - `state: "matched"`

### shared spec / adapter の分担

- shared spec 側は `anchor` / `reason` / `state` / `scopes` までを持つ。
- extension 側は [packages/extension/test/suite/index.ts](../../packages/extension/test/suite/index.ts) で、shared entry を
  - `detailFragment`
  - `blockedLabel`
  - async wait
  - failure message
  へ変換している。
- server 側は [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) で `server-application-ole` / `server-application-shape` / `server-application-shadowed` scope だけを読んでおり、この 3 本は消費しない。

## 観察結果

### 1. 3 本とも family canonical anchor source である

- 3 本とも [ApplicationWorkbookRootBuiltIn.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootBuiltIn.bas) 上の実 anchor に 1 対 1 で対応している。
- `numeric-selector` / `code-name-selector` / `static` / `matched` という語彙は、`applicationWorkbookRoot` 全体で使っている `reason` / `state` taxonomy に沿っている。
- extension helper の都合で一時的に生えた local datum ではなく、family table の residual slice とみなせる。

### 2. completion でも package-local に残すべきものは既に分離されている

- extension completion の固有事情は、`Value` の `detail` に `CheckBox.Value` が含まれるか、そして `Activate` / `Delete` のような blocked label が出ないかである。
- これらは shared spec には入っておらず、`mapExtensionWorkbookRootPositiveCompletionCases()` / `mapExtensionWorkbookRootClosedCompletionCases()` 側に閉じている。
- したがって shared spec に残るのは anchor / reason / state だけであり、completion adapter の都合が `test-support/` に漏れるわけではない。

### 3. local 化すると completion だけ正本が分裂する

- `applicationWorkbookRoot` は completion / hover / signature / semantic を同じ family table から拾っている。
- ここで extension-only completion 3 本だけを local table へ戻すと、同じ fixture を参照する negative matrix が
  - `test-support/workbookRootFamilyCaseTables.cjs`
  - `packages/extension/test/suite/index.ts`
  に分裂する。
- 既に interaction / semantic では「scope 非対称だけでは shared spec から外さない」と整理済みであり、completion だけ逆方向へ戻す理由は薄い。

### 4. server が mirror しないことと、shared spec に残す価値は別論点である

- server がこれら 3 本を読まないのは、`applicationWorkbookRoot` の completion negative で mirror している slice を `server-application-ole` / `server-application-shape` に絞っているためである。
- これは coverage 配分の問題であり、anchor 自体が extension-local だからではない。
- 後続で server mirror を増やすかどうかは別タスクとして扱えるため、このメモでは shared spec から外す根拠にはしない。

## 判断

### extension-only completion entry は shared spec に残す

- 上記 3 本の extension-only `completion.negative` entry は shared spec に残す。
- canonical anchor source は shared spec 側で持ち、`detailFragment` と blocked label は package-local adapter に残す。
- `hover` / `signature` / `semantic` と同様に、scope 非対称だけでは local 化しない。

### `detailFragment` / blocked label は shared spec へ持ち込まない

- `CheckBox.Value` のような `CompletionItem.detail` fragment と、`Activate` / `Delete` のような blocked label は extension adapter の責務とする。
- shared spec に `detailFragment` / `blockedLabel` を足すと、canonical anchor source ではなく adapter expectation を正本に抱え込むことになるため、この段階では導入しない。

### server mirror の拡張は別論点として残す

- 今回は「shared spec に残すか」を判断し、server 側へ mirror を増やすかは次タスクで切り分ける。
- これにより、「canonical anchor の置き場」と「coverage をどこまで mirror するか」を混同しない。

## 再評価の条件

- `detailFragment` / blocked label が family ごとに大きく分岐し、shared anchor より local expectation の方が主語として重要になったとき
- extension-only completion entry が増え続け、`applicationWorkbookRoot` family table より extension local table の方が読みやすいと判断できるほど residual slice が肥大化したとき
- server mirror を拡張した結果、completion residual slice がほぼ消え、shared spec に残す理由を別文脈で見直した方がよくなったとき

## 推奨方針

### 維持するもの

- `applicationWorkbookRoot` の completion canonical anchor は shared spec に置く
- `reason` / `state` / `scopes` は shared spec 側で持つ
- `detailFragment` / blocked label / async wait / failure message は extension adapter 側に残す

### 今やらないこと

- extension-only completion 3 本だけを `packages/extension/test/suite/index.ts` の local table へ戻す
- shared spec に `detailFragment` / `blockedLabel` を追加する
- server mirror 拡張の判断までこのメモで一緒に片付ける

## 関連文書

- 共通 policy: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
- 近い判断: [application-workbook-root-extension-only-interaction-shared-spec-feasibility.md](./application-workbook-root-extension-only-interaction-shared-spec-feasibility.md)
- workbook root family の fixture / shared spec 背景: [workbook-root-shadow-fixture-topology-feasibility.md](./workbook-root-shadow-fixture-topology-feasibility.md)
