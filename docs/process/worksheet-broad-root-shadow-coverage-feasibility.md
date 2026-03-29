# worksheet broad root family の shadow negative coverage を direct OLE route 以外へ広げる要否

## 結論

- 現時点では broad root shadow negative を `Shapes("CheckBox1").OLEFormat.Object` や root `.Item("Sheet One")` へ広げない。
- server 側の shadow negative は、引き続き direct `OLEObjects("CheckBox1").Object` route の 2 本だけで維持する。
- broad root shadow の route coverage は、root built-in 判定より後ろの分岐が増えたときだけ再評価する。

## 目的

`worksheet broad root family` の shadow negative は今のところ `OLEObjects("CheckBox1").Object` の direct route だけを固定している。  
このメモでは、`Shapes("CheckBox1").OLEFormat.Object` や `Worksheets.Item("Sheet One")` / `Application.Worksheets.Item("Sheet One")` を shadow 下でも追加で固定すべきかを整理する。

## 現状

- server 側の shadow negative は次の 2 本のみ。
  - `document service keeps unqualified worksheet broad root closed when Worksheets is shadowed`
  - `document service keeps Application worksheet broad root closed when Application is shadowed`
- どちらも `Worksheets` または `Application` の root identifier を shadow し、`OLEObjects("CheckBox1").Object` の completion / hover / signature が開かないことを固定している。
- broad root family の非 shadow 側では、direct route / child `.Item("CheckBox1")` / root `.Item("Sheet One")` / `Shapes("CheckBox1").OLEFormat.Object` まで user-facing path を既に固定している。

## 観察結果

### 1. broad root shadow は root 解決の段階で閉じる

- [packages/server/src/lsp/documentService.ts](../../packages/server/src/lsp/documentService.ts) では、`getWorkbookRootFamilyBuiltinContext()` が `resolution` を受け取った時点で broad root の built-in family を開かない。
- つまり `Worksheets` や `Application` が user-defined symbol として解決された時点で、`resolveWorksheetControlOwnerFromSidecar()` に渡る前に処理が止まる。
- このため、shadow 下では `OLEObjects` route、`Shapes` route、root `.Item("Sheet One")` の違いは後段まで到達せず、同じ root gating で閉じる。

### 2. 追加 route を足しても別の条件分岐を踏まない

- `Shapes("CheckBox1").OLEFormat.Object` は非 shadow 側では `OLEObject.Object` と別 route を通るが、shadow case ではその route 分岐に入る前に root resolution で閉じる。
- root `.Item("Sheet One")` も同様に、built-in `Worksheets` family として解決できたときだけ `resolveWorkbookRootFamilyPath()` の selector 解釈へ進む。
- したがって shadow 状態でこれらを追加しても、「root が shadow されたら built-in family に入らない」という同じ predicate を重ねて見るだけになる。

### 3. coverage を増やすなら、先に実装側の分岐増加が必要

- もし将来、shadow 状態でも root 後段の access kind や route ごとに別の fallback を持つようになれば、`Shapes` や root `.Item` の shadow negative は意味を持ちやすい。
- しかし現行実装では、root identifier が built-in でない時点で broad root family の処理自体が打ち切られる。
- 今の構造では route coverage を増やすより、root gating の 2 variants を明示 test で持つ方が効率がよい。

### 4. test コストに対して増える保証が小さい

- `Shapes` shadow negative や root `.Item` shadow negative を足すには、shadow text を増やすか、1 本の text にさらに route を追加する必要がある。
- どちらの場合も test は長くなるが、失敗時に増える情報は「やはり root shadow で閉じた」以上のものになりにくい。
- 現状の broad root shadow は conservative negative であり、coverage の広さより主語の明瞭さが重要である。

## 判断

### 今回は direct OLE route 以外へ広げない

- `Shapes("CheckBox1").OLEFormat.Object` の shadow negative は追加しない。
- root `.Item("Sheet One")` 系の shadow negative も追加しない。
- broad root shadow coverage は、現在の 2 本の direct OLE route で十分とみなす。

### 再評価の条件

- `getWorkbookRootFamilyBuiltinContext()` より後段で route 別の fallback や access kind 分岐が追加されたとき
- broad root shadow の実不具合が `Shapes` route や root `.Item("Sheet One")` でだけ再現したとき
- `Worksheets` / `Application` shadow 以外の broad root shadow variant が増え、root gating だけでは説明しづらい挙動差が出たとき
- review や triage で「direct OLE route だけでは route 差分を十分に説明できない」という指摘が続いたとき

## 推奨方針

### 維持するもの

- direct `OLEObjects("CheckBox1").Object` route の broad root shadow negative 2 本
- `Worksheets` shadow と `Application` shadow の root gating coverage
- 非 shadow 側での `Shapes` / root `.Item("Sheet One")` の user-facing 固定

### 今やらないこと

- `Shapes("CheckBox1").OLEFormat.Object` の shadow negative を server-only で追加する
- `Worksheets.Item("Sheet One")` / `Application.Worksheets.Item("Sheet One")` の shadow negative を server-only で追加する
- broad root shadow の route 違いを shared spec や local table に押し込む

## 関連文書

- server-only helper 判断: [worksheet-broad-root-shadow-server-helper-feasibility.md](./worksheet-broad-root-shadow-server-helper-feasibility.md)
- extension matrix へ広げない判断: [worksheet-broad-root-shadow-extension-matrix-feasibility.md](./worksheet-broad-root-shadow-extension-matrix-feasibility.md)
- broad root family の正本: [explicit-sheet-name-broad-root-feasibility.md](./explicit-sheet-name-broad-root-feasibility.md)
