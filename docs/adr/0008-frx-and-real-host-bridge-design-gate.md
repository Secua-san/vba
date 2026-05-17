# ADR 0008: `.frx` and Real Host Bridge Design Gate

## Status

Accepted

## Context

- 要件では `.frx` フォームバイナリと配置オブジェクト解析が入力候補に残っている。
- 既存の worksheet / chart control metadata 方針では、`.bas` / `.cls` / `.frm` / `.frx` だけでは worksheet / chart sheet 上の ActiveX control inventory を安定復元できないため、workbook package 由来の sidecar を static input として使っている。
- ADR [0007](0007-active-workbook-identity-provider-contract.md) は active workbook identity の snapshot schema と extension -> server notification 契約を固定している。
- ただし、実 Excel host と extension がどう接続するか、どの process / transport / lifecycle を採るかはまだ固定していない。
- `.frx` binary 解析と実 host bridge はどちらも、推測で実装すると parser / extension / user workflow の境界を広げやすい。

## Decision

- `.frx` 配置オブジェクト解析は、対象 binary format、対象 object、取得する metadata、failure mode、fixture 条件を要件化するまで product code へ入れない。
- `.frx` 解析は worksheet / chart control sidecar の代替 source として扱わない。UserForm `.frm` / `.frx` 由来の metadata と workbook package 由来の worksheet / chart sidecar metadata は別 source family として扱う。
- 実 Excel host bridge は、Excel host との接続方式を決める専用 ADR または process doc を先に作るまで product code へ入れない。
- ADR 0007 の `ActiveWorkbookIdentitySnapshot` schema と `vba/activeWorkbookIdentity` notification は維持するが、実 host 接続方式の決定を代替しない。
- server は引き続き host と直接通信しない。実 bridge を作る場合も extension が host lifecycle と transport を所有し、server は validated snapshot の consumer に留める。
- 設計ゲートの詳細な解除条件は [frx-and-real-host-bridge-design-gate.md](../process/frx-and-real-host-bridge-design-gate.md) を参照する。

## Consequences

- `.frx` と実 host bridge は、未確定仕様を product code で既成事実化しない。
- `.frx` support を進める場合は、UserForm binary metadata と worksheet / chart sidecar metadata の混線を避けられる。
- 実 host bridge を進める場合は、connection mode、lifecycle、timeout、error handling、multi-instance、logging を先にレビューできる。
- PR7 では parser / LSP / extension runtime / scripts を変更しない。
- 後続 PR は、ゲート解除条件を満たす小さな design PR か、決定済み境界に沿った implementation PR に分ける。
