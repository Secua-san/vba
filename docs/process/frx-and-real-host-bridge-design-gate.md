# `.frx` and Real Host Bridge Design Gate

## 結論

- `.frx` 配置オブジェクト解析と実 Excel host bridge は、どちらも product code 実装前に設計ゲートを通す。
- `.frx` は対象 binary format と取得 metadata を固定するまで解析しない。
- 実 host bridge は [ADR 0009](../adr/0009-excel-host-bridge-connection.md) の接続方式に従い、実装前に helper payload、timeout、logging、validation を固定する。
- 既存の `ActiveWorkbookIdentitySnapshot` schema は利用可能な contract とし、実 host への v1 接続方式は ADR 0009 の manual helper として扱う。

## 現行で確認済みの前提

- 要件には `.frx` フォームバイナリと配置オブジェクト解析が残っている。
- `resources/vbac/vbac.wsf` / extract / combine の source component 対象には `.frx` が含まれる。
- worksheet / chart control metadata は workbook package から生成した sidecar を static input として扱っている。
- 既存 docs では、`.frm` / `.frx` は UserForm には有効だが、worksheet / chart sheet 上の ActiveX control inventory source とは別であると整理している。
- active workbook identity は host -> extension -> server の snapshot schema と LSP notification までは固定済みである。
- 実 Excel host と extension の v1 接続方式は、extension-owned manual `cscript.exe` helper として ADR 0009 で固定済みである。

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

1. helper payload
   - helper が ADR 0007 の `ActiveWorkbookIdentitySnapshot` をどう生成するか。
   - 読み取る Excel property を最小化できているか。
2. lifecycle
   - `vba.refreshActiveWorkbookIdentity` 実行時にどの timeout / cancellation 境界で helper を起動するか。
   - Excel が未起動、複数 instance、workbook 切替、Excel 終了をどう扱うか。
3. transport contract
   - ADR 0007 の `ActiveWorkbookIdentitySnapshot` を host 側でどう生成するか。
   - timeout、retry、cancellation、stale snapshot の扱い。
4. user control
   - 明示 command 実行中 / 成功 / 失敗をどう表示するか。
   - status / log / error surface をどこに出すか。
5. security / safety
   - workbook や macro を実行しないことをどう保証するか。
   - helper の出力が snapshot payload 以外を resolver へ流さないことをどう保証するか。
6. validation
   - local unit test、mock host test、manual Windows test の境界。
   - extension host test が必要になる条件。

### ゲート解除条件

- connection mode は ADR 0009 に従う。
- extension / helper / Excel / server の責務境界が固定されている。
- timeout / retry / stale handling / logging が決まっている。
- user-facing resolver を開く条件が ADR 0006 / ADR 0007 と矛盾しない。
- 最小 validation command と、manual test が必要な場合の手順が決まっている。

## PR7 でやること

- この設計ゲートを docs と ADR 入口から参照できるようにする。
- product code、parser、LSP、extension runtime、script は変更しない。

## 次段候補

- `.frx` は、UserForm `.frm` text と `.frx` binary の責務分担を決める design PR を先に作る。
- 実 host bridge は、ADR 0009 の manual helper 境界に沿って payload、timeout、logging、validation を固定してから最小 implementation PR へ進む。
- どちらも design PR が通った後、最小 implementation PR へ分ける。
