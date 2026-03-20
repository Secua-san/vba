# ADR 0006: Workbook Binding Policy

## Status

Accepted

## Context

- broad root (`ActiveWorkbook.Worksheets("Sheet1")` / unqualified `Worksheets("Sheet1")`) を将来 user-facing に開くには、current bundle と runtime target workbook の同一性を明示的に結ぶ必要がある。
- current product は `loose files + .vba/worksheet-control-metadata.json` を static input として扱っており、sidecar は control inventory を保持する artifact である。
- `Application.ActiveWorkbook` は active window の workbook を返し、`Application.ThisWorkbook` は current macro code が動いている workbook を返すため、特に add-in では一致しない。
- `Workbook.FullName` は path を含む workbook 名を返し、`Workbook.Path` は workbook が未保存のとき空文字列になり得る。
- `Workbook.IsAddin` が `True` の workbook は add-in として動作し、window visibility や caller workbook との関係が通常 workbook と異なる。

## Decision

- current bundle と runtime workbook を結ぶ transport は、専用 artifact `<bundle-root>/.vba/workbook-binding.json` とする。
- `worksheet-control-metadata.json` には workbook binding 情報を混ぜない。control inventory と runtime workbook identity は別責務として扱う。
- workbook binding の primary key は、host から取得する `Workbook.FullName` を正規化した値とする。
- `Workbook.FullName` の v1 正規化は、Windows 前提で大文字小文字を無視し、`/` を `\` へそろえ、UNC 形式は保持し、`FullNameURLEncoded` は使わない。
- manifest には `fullName` を必須で持たせ、`name` / `path` / `isAddIn` は診断と保守条件のために併記する。
- `Workbook.Path` が空の workbook、または `Workbook.IsAddin = True` の workbook は、broad root binding の対象にしない。
- v1 generator は、saved かつ non-addin workbook にだけ `workbook-binding.json` を生成し、manifest 不在は「binding disabled」として扱う。
- workbook package mode は binding の runtime transport ではなく、manifest を生成する source of truth 候補として扱う。
- broad root を実際に開くのは、host / extension / server 間で active workbook identity を渡す契約が定義された後に限る。

## Consequences

- control metadata sidecar は volatile な path 情報を持たず、再生成差分を最小限に保てる。
- workbook rename / move により `fullName` が変わると binding manifest は更新が必要になる。
- unsaved workbook と add-in workbook は broad root の対象外のまま残る。
- broad root の将来実装では、manifest lookup と host から渡される active workbook identity の両方が揃ったときだけ resolver を有効化する。
- 調査詳細と schema 案は `docs/process/workbook-binding-manifest-feasibility.md` を参照する。
