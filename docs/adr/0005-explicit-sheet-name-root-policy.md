# ADR 0005: Explicit Sheet-Name Root Policy

## Status

Accepted

## Context

- worksheet document module alias (`Sheet1`) からの sidecar lookup は、現在 `sheetCodeName + shapeName` または `sheetCodeName + codeName` を根拠に user-facing 解決している。
- 一方、`Worksheets("Sheet1")` や `ThisWorkbook.Worksheets("Sheet1")` のような explicit sheet-name root は、Office VBA の正本では worksheet 名を selector に使う。
- `Worksheet.CodeName` の正本では、code name は `Sheet1.Range("A1")` のような expression alias であり、sheet name とは独立に変更され得る。
- `ThisWorkbook.Worksheets("Sheet1")` と `ThisWorkbook.Worksheets.Item("Sheet1")` は current bundle の workbook identity を静的に固定できるため、既に user-facing に解決している。
- `ActiveWorkbook.Worksheets("Sheet1")` と `ActiveWorkbook.Worksheets.Item("Sheet1")` は active workbook identity snapshot と `workbook-binding.json` の match がそろったときだけ、current bundle sidecar lookup の候補へ開く実装が追加された。
- unqualified `Worksheets("Sheet1")` と `Application.Worksheets("Sheet1")` は Office VBA 上で active workbook を対象にするが、root が暗黙である分だけ broad root 境界を別途固定する必要がある。

## Decision

- explicit sheet-name root の join key は `sheetCodeName` ではなく `sheetName` を使う。
- `sheetCodeName` は引き続き worksheet document module alias (`Sheet1`) と control code name 導線 (`Sheet1.chkFinished`) の join key として維持する。
- workbook identity を静的に固定できる `ThisWorkbook.Worksheets("Sheet1")` と `ThisWorkbook.Worksheets.Item("Sheet1")` は、引き続き explicit sheet-name root の基本経路として扱う。
- `ActiveWorkbook.Worksheets("Sheet1")` と `ActiveWorkbook.Worksheets.Item("Sheet1")` は、`available` snapshot、manifest 存在、manifest match、対応 owner の 4 条件がそろったときだけ user-facing にしてよい。
- unqualified `Worksheets("Sheet1")` と `Application.Worksheets("Sheet1")` は、Office VBA 上は active workbook alias とみなし、将来 user-facing にする場合も `ActiveWorkbook.Worksheets("Sheet1")` と完全に同じ gating 条件を使う。
- unqualified broad root family は `Worksheets("literal sheetName")` / `Worksheets.Item("literal sheetName")` と `Application.Worksheets("literal sheetName")` / `Application.Worksheets.Item("literal sheetName")` を同一扱いにし、`available` snapshot と manifest match がそろったときだけ sidecar lookup を開く。
- `Worksheets.Item("literal sheetName")` と `Application.Worksheets.Item("literal sheetName")` は `Worksheets.Item property` の既定メンバー規則に従い、direct call form と同じ gating 条件・同じ導線で扱う。
- built-in broad root gating は root identifier が built-in `Worksheets` collection として解決できた場合にだけ適用し、同名の変数、関数、メンバーへ shadow されているときは user-defined symbol を優先する。
- `OLEObject.Object` と `Shape.OLEFormat.Object` の explicit sheet-name root 対応は、将来 `workbook root identity + sheetName + shapeName` lookup helper を共有する前提で進める。
- broad root の再評価は、current bundle と target workbook の同一性を静的または明示設定で保証できる仕組みが追加された場合に限る。

## Consequences

- `Sheet1` alias 導線と `Worksheets("Sheet1")` 導線で join key を混同しないため、sheet name と code name がずれた workbook でも誤解決を避けやすい。
- `ActiveWorkbook.Worksheets("Sheet1")` / `.Item("Sheet1")` と unqualified `Worksheets("Sheet1")` / `.Item("Sheet1")` を同じ active-workbook family として扱うことで、root の書き方だけで user-facing 挙動が分岐しにくくなる。
- `ThisWorkbook.Worksheets("Sheet1")` / `.Item("Sheet1")` 連鎖でも workbook root identity を保持して sidecar resolver へ伝播する経路が必要であり、この最小経路は実装済みである。
- `Worksheets.Item("Sheet1")` と `Application.Worksheets.Item("Sheet1")` を direct call form と同じ family に含めることで、既定メンバー省略の有無だけで user-facing 挙動が変わる不整合を避けやすい。
- `Sheets`、`ActiveSheet`、numeric / dynamic / grouped selector を broad root family から外すことで、sheet type 混在や join key 不在の誤補完を抑えられる。
- unqualified broad root は shadow 可能な識別子なので、built-in と user-defined symbol の優先順位を docs と test の両方で固定する必要がある。
- 調査詳細や除外境界の補足は `docs/process/shape-oleformat-object-explicit-sheet-root-feasibility.md` と `docs/process/explicit-sheet-name-broad-root-feasibility.md` を参照する。
