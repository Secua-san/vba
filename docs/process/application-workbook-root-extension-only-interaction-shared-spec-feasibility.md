# application workbook root family の extension-only interaction / semantic entry を shared spec に残す境界

## 結論

- 現時点では `applicationWorkbookRoot` の extension-only `hover` / `signature` / `semantic` entry を shared spec に残す。
- `packages/extension/test/suite/index.ts` 側の local table へ戻さず、引き続き [test-support/workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) を family canonical anchor source として扱う。
- ただし extension-only `completion` entry は別論点として残し、このメモでは扱わない。

## 目的

前段で `worksheetBroadRoot` の extension-only interaction negative は shared spec に残す判断を固定し、共通 policy へも `scope 非対称だけでは shared spec から外さない` ルールを追加した。  
このメモでは、その同じ物差しを `applicationWorkbookRoot` の extension-only `hover` / `signature` / `semantic` entry に適用してよいかを整理する。

## 現状

### `applicationWorkbookRoot` の extension-only interaction / semantic entry

- `hover.negative`
  - `Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Valu`
  - `Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu`
- `signature.negative`
  - `Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Select(`
  - `Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(`
  - `Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(`
- `semantic.negative`
  - `Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value`
  - `Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value`

### adapter 側

- [packages/extension/test/suite/index.ts](../../packages/extension/test/suite/index.ts) は `scope: "extension"` の shared entry をそのまま positive / negative / closed / shadowed matrix へ流している。
- [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) は `server-application-ole` / `server-application-shape` / `server-application-shadowed` scope の entry を読むが、上記 extension-only interaction / semantic entry は消費しない。

### fixture 構成

- 通常系は [ApplicationWorkbookRootBuiltIn.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootBuiltIn.bas) が持つ。
- shadow semantic negative は [ApplicationWorkbookRootShadowed.bas](../../packages/extension/test/fixtures/ApplicationWorkbookRootShadowed.bas) に載っている。
- どちらも `applicationWorkbookRoot` family table と 1 対 1 に対応する shared fixture であり、extension local helper のためだけに増えた anchor ではない。

## 観察結果

### 1. extension-only interaction / semantic entry も family の state / reason taxonomy に乗っている

- これらの entry は `code-name-selector`、`numeric-selector`、`shadowed-root` のように、`applicationWorkbookRoot` 全体で使っている reason / state 語彙に沿っている。
- つまり local helper の都合で生えた一時的な anchor ではなく、family table の中で意味を持つ residual slice である。
- 共通 policy の 4.5 で定義した「anchor / reason / state が family canonical source として意味を持つ」という条件を満たしている。

### 2. server が読まないことと、shared spec に残す価値は別論点である

- server 側が読まない理由は、coverage の mirror 先を completion / closed-state interaction / selected semantic に絞っているためであり、anchor 自体が extension local だからではない。
- たとえば `Application.ThisWorkbook.Worksheets("Sheet1")` の code-name selector negative は、completion では一部 server scope もあるが、hover / signature は extension-only で残っている。
- これは「どの API まで server で mirror するか」の差であって、「その anchor が family canonical source かどうか」の差ではない。

### 3. local file へ戻すと、同じ family の negative matrix が source 別に割れる

- `applicationWorkbookRoot` は state ごとに positive / negative / closed / shadowed を shared table から組み立てている。
- ここで extension-only interaction / semantic entry だけ local 化すると、同じ `ApplicationWorkbookRootBuiltIn.bas` / `ApplicationWorkbookRootShadowed.bas` を参照する anchor が `test-support/` と `packages/extension/` に分裂する。
- 特に shadow semantic negative は、shared `state: "shadowed"` 語彙を使っているため、hover / signature と semantic を別正本へ割ると review 時に追いにくい。

### 4. package-local adapter へ残るべきものは既に分離されている

- extension 側の package 固有事情は、hover / signature wait 条件、semantic token decode 後の assertion 形、failure message である。
- それらは shared spec に入っておらず、[packages/extension/test/suite/index.ts](../../packages/extension/test/suite/index.ts) の mapper / helper に閉じている。
- したがって、この entry 群を shared spec に残しても、extension 固有の手続きが `test-support/` へ漏れるわけではない。

### 5. 残る未整理は `completion` 側であり、このメモと混ぜない方が読みやすい

- `applicationWorkbookRoot` には extension-only `completion.negative` も残っている。
- ただし completion は blocked symbol や detail fragment の扱いが絡み、interaction / semantic とは adapter 境界の見え方が少し違う。
- 今回は interaction / semantic に対象を絞り、`completion` の extension-only entry 境界は次タスクへ分離する方が論点が混ざらない。

## 判断

### extension-only interaction / semantic entry は shared spec に残す

- `applicationWorkbookRoot` の extension-only `hover` / `signature` / `semantic` entry は shared spec に残す。
- これらは `ApplicationWorkbookRootBuiltIn.bas` / `ApplicationWorkbookRootShadowed.bas` の canonical anchor source であり、local helper 固有データではない。
- server mirror の有無だけを理由に local file へ戻さない。

### 今回は `completion` を一緒に片付けない

- extension-only `completion.negative` は次タスクへ分離し、このメモでは判断を出さない。
- 理由は、interaction / semantic と completion では blocked symbol / detail fragment まわりの adapter 境界が少し異なるため。

### 再評価の条件

- extension-only interaction / semantic entry が helper 固有の待機条件や message に強く依存し、canonical anchor source と呼びづらくなったとき
- `applicationWorkbookRoot` の server mirror 範囲が広がり、extension-only residual slice がほぼ消えたとき
- `ApplicationWorkbookRootBuiltIn.bas` / `ApplicationWorkbookRootShadowed.bas` の anchor を local file と shared spec の両方で管理した方が review しやすいと判断できるほど、family table の見通しが悪化したとき
- completion extension-only entry の整理結果が interaction / semantic まで波及し、family 全体を local 化した方がよいという新しい根拠が出たとき

## 推奨方針

### 維持するもの

- `applicationWorkbookRoot` family table 内での interaction / semantic canonical anchor 管理
- `state` / `reason` / `rootKind` を shared spec 側で持ち、package-local adapter は assertion shape に専念する構造
- extension 側の dedicated shadow fixture と shared `state: "shadowed"` の対応

### 今やらないこと

- extension-only interaction / semantic entry を `packages/extension/test/suite/index.ts` の local table へ戻す
- `completion` の論点まで同時に決める
- server が使わない entry を一律で shared spec から外す

## 関連文書

- 共通 policy: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
- 近い判断: [worksheet-broad-root-extension-only-interaction-shared-spec-feasibility.md](./worksheet-broad-root-extension-only-interaction-shared-spec-feasibility.md)
- workbook root family の fixture topology: [workbook-root-shadow-fixture-topology-feasibility.md](./workbook-root-shadow-fixture-topology-feasibility.md)
