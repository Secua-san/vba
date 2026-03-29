# worksheet broad root family の server-only shadow negative を helper / table へ寄せる要否

## 結論

- 現時点では `Worksheets` shadow / `Application` shadow の server-only negative を helper / table へ寄せない。
- [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の 2 本の単発 test を、そのまま明示的に維持する。
- `test-support/workbookRootFamilyCaseTables.cjs` には broad root shadow 用の local scope や `state: "shadowed"` を追加しない。
- server 側の broad root shadow 専用 helper も、今の段階では追加しない。

## 目的

`worksheet broad root family` の shadow case は extension matrix に上げず、server-only negative として維持する方針を前段で固定した。  
このメモでは、その server-only negative について、少なくとも server 側だけでも配列駆動 helper や local table へ寄せる価値があるかを整理する。

## 現状

- server 側には次の 2 本がある。
  - `document service keeps unqualified worksheet broad root closed when Worksheets is shadowed`
  - `document service keeps Application worksheet broad root closed when Application is shadowed`
- どちらも `createWorksheetBroadRootFixture()` を使う以外は、shadow 用宣言、query token、completion / hover / signature の negative assert を test 本体へ直書きしている。
- broad root family の built-in positive / non-target / snapshot gating は、既存の `mapSharedWorkbookRoot*` helper と `test-support/workbookRootFamilyCaseTables.cjs` で配列駆動化されている。
- ただし broad root shadow は shared spec に載せず、local helper も持たない。

## 観察結果

### 1. いまの重複量は小さい

- shadow test は 2 本だけで、どちらも 3 assertion に閉じている。
- 重複しているのは `setActiveWorkbookIdentitySnapshot(createMatchedActiveWorkbookIdentitySnapshot())` と completion / hover / signature の negative assert の形だけである。
- この程度の重複量では、配列駆動 helper や local table を増やしても削減できる行数は小さい。

### 2. shadow 宣言の違いが test の主語になっている

- `Worksheets` shadow は `Private Function Worksheets(...) As String` で built-in collection を塞いでいる。
- `Application` shadow は `Private Type Application` と `Dim Application As Application` の組み合わせで qualifier を塞いでいる。
- この差は単なる parameter 差ではなく、「何が built-in 判定を壊しているか」を test 本文で読めること自体に意味がある。
- helper / table へ寄せすぎると、肝心の shadow topology が `declarationText` の長い文字列へ押し込まれ、失敗時に主語が追いにくくなる。

### 3. workbook root family で helper 化した理由とは性質が違う

- workbook root family は completion / hover / signature / semantic、static / matched / closed / shadowed、server / extension の組み合わせが多く、shared spec を置く実益が大きかった。
- broad root shadow の server-only negative は、「2 variants x 3 assertions」に留まっている。
- 同じ整理手法を早めに持ち込むと、matrix 化の受益が少ないのに helper だけが増える。

### 4. local helper を作っても shared 化にはつながらない

- 今回の候補は server 側だけの local helper / table であり、extension や shared spec とは切り離されたままである。
- つまり増えるのは `documentService.test.js` 内の抽象化だけで、package 間の正本統一には寄与しない。
- そのため「読みやすさが上がるか」が唯一の判断軸になるが、現状では明示 test の方が読みやすい。

### 5. 将来の拡張点は helper より coverage 側にある

- 現在の shadow negative は `OLEObjects("CheckBox1").Object` の direct route しか見ていない。
- 将来追加を考えるなら、まず `Shapes("CheckBox1").OLEFormat.Object` や root `.Item("Sheet One")` を shadow 下でも閉じるか、といった coverage の要否が先に来る。
- coverage が増えて 4 本以上の shadow variant になった時点ではじめて、local helper / table 抽出の受益が見えやすくなる。

## 判断

### 今回は単発 test のまま維持する

- `Worksheets` shadow / `Application` shadow は、現行の明示 test をそのまま維持する。
- broad root shadow 専用の server-local helper / table は追加しない。
- `test-support/workbookRootFamilyCaseTables.cjs` や shared spec policy へも広げない。

### 再評価の条件

- broad root shadow variant が 3 種以上に増え、同じ negative assert を 4 本以上で繰り返すようになったとき
- `Shapes` route や root `.Item("Sheet One")` を含む shadow coverage を server-only で追加し、declaration / query / expected result の組み合わせが増えたとき
- CodeRabbit や `reviewer` から「broad root shadow test の重複が読みづらい」という同種指摘が 2 PR 以上連続したとき
- 失敗 triage で「どの shadow variant が落ちたか」を helper message へ寄せた方が読みやすい、と実運用で判断できたとき

## 推奨方針

### 維持するもの

- [packages/server/test/documentService.test.js](../../packages/server/test/documentService.test.js) の 2 本の明示 test
- `createWorksheetBroadRootFixture()` を使った bundle / sidecar setup の共通化
- broad root shadow は server-only negative として扱う前段判断

### 今やらないこと

- `assertWorksheetBroadRootShadowCases()` のような専用 helper を増やす
- `worksheetBroadRootShadowCases` のような local table を新設する
- broad root shadow を shared spec の scope や state へ昇格させる

## 関連文書

- extension matrix へ広げない判断: [worksheet-broad-root-shadow-extension-matrix-feasibility.md](./worksheet-broad-root-shadow-extension-matrix-feasibility.md)
- route coverage の判断: [worksheet-broad-root-shadow-coverage-feasibility.md](./worksheet-broad-root-shadow-coverage-feasibility.md)
- broad root family の正本: [explicit-sheet-name-broad-root-feasibility.md](./explicit-sheet-name-broad-root-feasibility.md)
- shared spec 境界: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
