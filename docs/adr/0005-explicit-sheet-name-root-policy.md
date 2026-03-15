# ADR 0005: Explicit Sheet-Name Root Policy

## Status

Proposed

## Context

- worksheet document module alias (`Sheet1`) からの sidecar lookup は、現在 `sheetCodeName + shapeName` または `sheetCodeName + codeName` を根拠に user-facing 解決している。
- 一方、`Worksheets("Sheet1")` や `ThisWorkbook.Worksheets("Sheet1")` のような explicit sheet-name root は、Office VBA の正本では worksheet 名を selector に使う。
- `Worksheet.CodeName` の正本では、code name は `Sheet1.Range("A1")` のような expression alias であり、sheet name とは独立に変更され得る。
- current resolver は document module root にだけ workbook bundle identity を持たせており、generic `Worksheet` owner へ降りた時点では `which workbook / which worksheet` の provenance を保持していない。
- `ActiveWorkbook.Worksheets("Sheet1")` と unqualified `Worksheets("Sheet1")` は active workbook 依存であり、current bundle の sidecar に静的に結び付けると誤解決の余地がある。

## Decision

- explicit sheet-name root の join key は `sheetCodeName` ではなく `sheetName` を使う。
- `sheetCodeName` は引き続き worksheet document module alias (`Sheet1`) と control code name 導線 (`Sheet1.chkFinished`) の join key として維持する。
- explicit sheet-name root を user-facing に広げる最初の候補は、workbook identity を静的に固定できる `ThisWorkbook.Worksheets("Sheet1")` 限定とする。
- `ActiveWorkbook.Worksheets("Sheet1")` と unqualified `Worksheets("Sheet1")` は、この段階では user-facing にしない。
- `OLEObject.Object` と `Shape.OLEFormat.Object` の explicit sheet-name root 対応は、将来 `workbook root identity + sheetName + shapeName` lookup helper を共有する前提で進める。

## Consequences

- `Sheet1` alias 導線と `Worksheets("Sheet1")` 導線で join key を混同しないため、sheet name と code name がずれた workbook でも誤解決を避けやすい。
- broad root 展開は遅くなるが、active workbook 依存の誤補完を抑えられる。
- 次段の実装では、`ThisWorkbook.Worksheets("Sheet1")` 連鎖でも workbook root identity を保持して sidecar resolver へ伝播する必要がある。
- 調査詳細や除外境界の補足は `docs/process/shape-oleformat-object-explicit-sheet-root-feasibility.md` を正本とする。
