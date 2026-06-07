# `.frx` and Real Host Bridge Design Gate

## 結論

- `.frx` 配置オブジェクト解析と実 Excel host bridge は、どちらも product code 実装前に設計ゲートを通す。
- `.frx` は対象 binary format と取得 metadata を固定するまで解析しない。
- 実 host bridge は [ADR 0009](../adr/0009-excel-host-bridge-connection.md) の接続方式に従う。最小 v1 は manual refresh command として helper payload、timeout、logging、validation を固定済みで、自動接続や常駐化は引き続き設計ゲート対象とする。
- 既存の `ActiveWorkbookIdentitySnapshot` schema は利用可能な contract とし、実 host への v1 接続方式は ADR 0009 の manual helper として扱う。

## 現行で確認済みの前提

- 要件には `.frx` フォームバイナリと配置オブジェクト解析が残っている。
- `resources/vbac/vbac.wsf` / extract / combine の source component 対象には `.frx` が含まれる。
- worksheet / chart control metadata は workbook package から生成した sidecar を static input として扱っている。
- 既存 docs では、`.frm` / `.frx` は UserForm には有効だが、worksheet / chart sheet 上の ActiveX control inventory source とは別であると整理している。
- active workbook identity は host -> extension -> server の snapshot schema と LSP notification までは固定済みである。
- 実 Excel host と extension の v1 接続方式は、extension-owned manual `cscript.exe` helper として ADR 0009 で固定済みである。
- v1 manual refresh は `vba.refreshActiveWorkbookIdentity` から 1 回だけ `cscript.exe` helper を起動し、valid `ActiveWorkbookIdentitySnapshot` だけを `vba/activeWorkbookIdentity` notification へ流す。

## `.frx` 解析ゲート

### 非目標

- `.frx` を worksheet / chart control sidecar の代替 source として扱わない。
- `.frm` text parser の拡張と `.frx` binary parser を同じ PR に混ぜない。
- binary record を推測で読み、取得できた field だけを product schema に流さない。

### 実装前に決めること

1. 対象 format
   - UserForm `.frx` のどの binary format / record を対象にするか。
   - Excel export 由来の `.frx` だけを対象にするか、互換 exporter の `.frx` も扱うか。
2. 対象 object
   - UserForm 配置 control だけを扱うか。
   - 画像や binary asset のような non-control payload を解析対象外にするか。
3. 最小 metadata
   - control name / type / container / index / caption など、product 側で使う field。
   - `.frm` text 側から取る field と `.frx` binary 側から取る field の責務分担。
4. 出力 schema
   - 既存 sidecar へ混ぜるか、UserForm 専用 artifact に分けるか。
   - unsupported / unknown record の表し方。
5. failure mode
   - parse failure、unknown record、partial metadata、encoding mismatch をどう扱うか。
   - diagnostic に出すか、log に閉じるか。
6. fixture 条件
   - 最小 real fixture または synthetic fixture を置けるか。
   - binary fixture の更新理由と review 方法をどう残すか。

### ゲート解除条件

- 対象 format と metadata field の一覧が docs に固定されている。
- `.frm` / `.frx` / sidecar の責務境界が固定されている。
- fixture 方針と最小 test command が決まっている。
- unknown / unsupported を product code で推測処理しない failure mode が決まっている。

## 実 Excel host bridge ゲート

### 非目標

- server から Excel host へ直接接続しない。
- host 未接続時に placeholder snapshot を送って resolver を開かない。
- ADR 0009 の manual helper v1 に、自動 polling、long-running helper、startup bridge を混ぜない。

### 実装前に決めること

最小 v1 manual refresh では以下を固定済みである。自動 polling、long-running helper、startup bridge、複数 instance 識別を追加する場合は、この節を更新してから実装する。

1. helper payload
   - helper は ADR 0007 の `ActiveWorkbookIdentitySnapshot` だけを stdout JSON として返す。
   - 読み取る Excel property は通常 workbook では `ActiveWorkbook.FullName` / `Name` / `Path` / `IsAddin`、Protected View では取得できる場合のみ `ActiveProtectedViewWindow.SourceName` / `SourcePath` に限定する。
2. lifecycle
   - `vba.refreshActiveWorkbookIdentity` 実行時に 10 秒 timeout / VS Code progress cancellation 境界で helper を起動する。
   - Excel 未起動は valid `unavailable` snapshot の `reason=host-unreachable` とする。複数 instance 識別、workbook 切替追跡、Excel 終了監視は v1 に含めない。
3. transport contract
   - extension は helper stdout を JSON parse し、`parseActiveWorkbookIdentitySnapshot()` で validation してから server へ送る。
   - retry と自動 stale 更新は v1 に含めない。helper failure、timeout、invalid payload、stale `observedAt` では extension が `unavailable` / `host-error` を送って cached binding を閉じる。cancellation では snapshot を送らない。
4. user control
   - 明示 command 実行中は VS Code progress、成功 / unsupported / unavailable / protected-view / 失敗は VS Code message で表示する。
   - log は `VBA Active Workbook` output channel に出す。
5. security / safety
   - helper は property read だけを行い、workbook や macro を実行せず、workbook を mutate しない。
   - helper の出力は snapshot payload 以外を resolver へ流さない。invalid payload は extension validation で止める。
6. validation
   - local validation は extension build、VSIX verifier、`npm run smoke:active-workbook-identity` で command/package surface と実 helper stdout schema を確認する。
   - 実 Excel workbook を開いた状態の manual scenario は `npm run smoke:active-workbook-identity -- --expect-state available --expect-full-name <workbook-full-name>` で検証する。
   - extension host test は明示承認時だけ実行する。

### ゲート解除条件

- connection mode は ADR 0009 に従う。
- extension / helper / Excel / server の責務境界が固定されている。
- timeout / retry / stale handling / logging が決まっている。
- user-facing resolver を開く条件が ADR 0006 / ADR 0007 と矛盾しない。
- 最小 validation command と、manual test が必要な場合の手順が決まっている。

## 現行で product code に入れないこと

- `.frx` binary parser は、この設計ゲートを解除するまで product code、parser、LSP、extension runtime、script に追加しない。
- automatic polling、long-running helper、startup bridge、複数 Excel instance 識別は、ADR 0009 の manual helper v1 には含めない。

## 次段候補

- `.frx` は、UserForm `.frm` text と `.frx` binary の責務分担を決める design PR を先に作る。
- 実 host bridge は、manual helper v1 の real Excel smoke と release verifier を維持する。自動接続へ進む場合は、先にこのゲートと ADR 0009 を更新する。
- どちらも design PR が通った後、最小 implementation PR へ分ける。
