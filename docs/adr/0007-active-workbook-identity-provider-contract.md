# ADR 0007: Active Workbook Identity Provider Contract

## Status

Accepted

## Context

- broad root (`ActiveWorkbook.Worksheets("Sheet1")` / unqualified `Worksheets("Sheet1")`) を将来 user-facing に開くには、runtime の active workbook identity を current bundle の `workbook-binding.json` と照合する必要がある。
- ADR [0006](0006-workbook-binding-policy.md) で、binding の primary key は `Workbook.FullName` を正規化した値とし、saved かつ non-addin workbook だけを対象にする方針を固定した。
- `Application.ActiveWorkbook` は active window の workbook を返すが、window が無い場合や Info / Clipboard window が active の場合は `Nothing` を返す。また active Protected View window はこの property では取れない。
- Protected View workbook は `ProtectedViewWindow.Workbook` からは参照できるが、`Workbooks` collection の member ではなく、許可されない操作は error になる。
- `Workbook.Saved` の正本では、未保存 workbook は `Path` が空文字列になり得る。
- current product の役割分離は `extension = VS Code / 外部 host との接続`, `server = resolver / diagnostics / cache`, `core = 純粋な解析 helper` である。

## Decision

- runtime の active workbook identity は、host -> extension -> server で共通に使う単一 schema `ActiveWorkbookIdentitySnapshot` として扱う。
- extension は外部 host との通信と lifecycle を担当し、server には custom LSP notification `vba/activeWorkbookIdentity` で snapshot 全体を通知する。
- server は host と直接通信しない。受信した snapshot を cache し、manifest matching と resolver gating にだけ使う。
- host / extension は `Workbook.FullName` / `Name` / `Path` / `IsAddin` の raw 値を渡し、v1 正規化と manifest 照合は server 側で一元化する。
- snapshot の v1 state は `available` / `unavailable` / `protected-view` / `unsupported` とする。
- `unsupported` の reason は少なくとも `unsaved` と `addin` を持つ。`protected-view` は別 state とし、`SourcePath` / `SourceName` を log 用 metadata として保持してよいが、resolver は有効化しない。
- broad root resolver を有効化する条件は、`state = available`、manifest 存在、manifest match 成功、対応 owner であることの 4 条件がそろったときだけとする。
- server は snapshot を cache する前に schema validation を行い、不正 payload は broad root を閉じたまま reject する。
- server log は snapshot 受信結果と gating 結果を分けて記録し、`available` / `unavailable` / `protected-view` / `unsupported` / `match` / `mismatch` / `invalid-payload` を識別できる code を持たせる。

## Consequences

- host 実装が VBA / COM / 外部 process のどれであっても、extension と server の境界は同じ snapshot schema で保てる。
- `Workbook.FullName` の正規化規則は server 側 1 箇所に寄るため、host 実装ごとの差で broad root の挙動がぶれにくい。
- Protected View、unsaved workbook、add-in workbook は「identity を取得できたが broad root には使えない」状態として区別できる。
- resolver 実装前でも、extension / server 間の transport、cache、log を read-only に先行実装できる。
- 調査詳細、field 候補、state ごとの payload は `docs/process/active-workbook-identity-provider-contract.md` を参照する。
