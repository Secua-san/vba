# worksheet broad root family の extension-only non-target hover / signature negative を shared spec に残す要否

## 結論

- 現時点では `worksheetBroadRoot.hover.negative` / `signature.negative` の extension-only entry を shared spec に残す。
- server scope を持たないことだけを理由に、`packages/extension/test/suite/index.ts` 側の local table へ戻さない。
- `test-support/workbookRootFamilyCaseTables.cjs` は、引き続き `WorksheetBroadRootBuiltIn.bas` の canonical anchor source として扱う。

## 目的

前段で、`worksheet broad root family` の non-target hover / signature negative を server scope へ広げない判断を固定した。  
このメモでは、その結果 extension だけが使う shared entry を、なお shared spec に残す意味があるかを整理する。

## 現状

### `worksheetBroadRoot`

- [test-support/workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) の `worksheetBroadRoot.hover.negative` / `signature.negative` は `extension` scope だけを持つ。
- [packages/extension/test/suite/index.ts](../../packages/extension/test/suite/index.ts) はその entry を `assertWorkbookRootNoHoverCases()` / `assertWorkbookRootNoSignatureCases()` へ流している。
- server 側は、[worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md](./worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md) の判断どおり、これらを消費しない。

### workbook root family 全体

- `applicationWorkbookRoot` には、hover / signature / semantic に extension-only scope の shared entry が既に複数ある。
- つまり `test-support/workbookRootFamilyCaseTables.cjs` は、現時点でも「両 package 完全対称の entry だけを置く場所」ではなく、「family canonical anchor を scope 付きで持つ場所」として使われている。

## 観察結果

### 1. `scopes` 配列がある以上、scope 非対称は schema 上の想定内である

- shared case spec は最初から `scopes` を持ち、`server-application-ole` や `server-worksheet-broad-root-item` のように、どの package / slice がその entry を消費するかを表している。
- この設計なら、ある entry が `extension` だけを持つこと自体は schema 違反ではなく、単に「現時点では extension だけが使う canonical anchor」という意味になる。
- scope 非対称だけを理由に shared spec から外すと、`scopes` を持たせた設計意図と逆行する。

### 2. これらの entry は package 固有仕様ではなく、fixture anchor の正本である

- `worksheetBroadRoot.hover.negative` / `signature.negative` は、`WorksheetBroadRootBuiltIn.bas` に置かれた non-target root / selector の canonical anchor を表している。
- ここで shared 化しているのは async wait 条件や failure message ではなく、あくまで fixture 上の anchor と reason である。
- したがって、entry 自体は extension 固有の test helper ではなく、family canonical source として shared 側に置く意味がある。

### 3. local table へ戻すと、同じ family の anchor source が 2 つに割れる

- いま `worksheetBroadRoot` は completion positive / negative、interaction positive、interaction negative を同じ family table で管理している。
- interaction negative だけ extension local に戻すと、同じ `WorksheetBroadRootBuiltIn.bas` を参照する anchor 群が `test-support/` と `packages/extension/` に分かれる。
- その結果、fixture 編集時の drift check が review 上追いにくくなり、「正例は shared、負例は local」という境界だけが増える。

### 4. package-local に残すべきものは、すでに adapter 層へ隔離できている

- extension 側が持っている package 固有事情は、`waitForHoverAtToken()` の待機、`waitForNoSignatureHelpAtToken()` の判定、failure message の文言である。
- それらは [packages/extension/test/suite/index.ts](../../packages/extension/test/suite/index.ts) の mapper / helper に残っており、shared spec 側へ漏れていない。
- つまり「shared spec に残すと extension 固有事情が混ざる」という懸念は、現在の `worksheetBroadRoot.hover.negative` / `signature.negative` には当てはまらない。

## 判断

### extension-only entry のまま shared spec に残す

- `worksheetBroadRoot.hover.negative` / `signature.negative` は `extension` scope のまま shared spec に残す。
- server が使わない entry であっても、canonical fixture anchor と reason を表すなら shared spec から外さない。
- local へ戻すかどうかの判断軸は「両 package で使うか」ではなく、「anchor / reason が family canonical source か」「package 固有事情が adapter 層に隔離されているか」とする。

### 再評価の条件

- `worksheetBroadRoot` interaction negative が extension 側 helper の都合に強く依存し、anchor / reason より待機条件や message の方が主体になったとき
- `WorksheetBroadRootBuiltIn.bas` の anchor を shared spec と local table の両方で持つ必要が出て、shared spec の canonical 性が崩れたとき
- CodeRabbit や `reviewer` から「extension-only entry が shared spec の見通しを悪化させている」という指摘が複数回続いたとき
- `applicationWorkbookRoot` など他 family の extension-only entry も含めて、scope 非対称 entry を family 単位ではなく local file 単位で再編した方が読みやすいと判断できたとき

## 推奨方針

### 維持するもの

- `worksheetBroadRoot` family table 内での interaction negative canonical anchor 管理
- package-local adapter による async wait / failure message / assertion shape の分離
- `scopes` による package 適用範囲の明示

### 今やらないこと

- `worksheetBroadRoot.hover.negative` / `signature.negative` を extension local の配列へ戻す
- server scope が無い entry を一律で shared spec から排除する
- `WorksheetBroadRootBuiltIn.bas` の anchor 正本を `test-support/` と `packages/extension/` の両方に持つ

## 関連文書

- 前段判断: [worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md](./worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md)
- 共通 policy: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
- broad root family の機能正本: [explicit-sheet-name-broad-root-feasibility.md](./explicit-sheet-name-broad-root-feasibility.md)
