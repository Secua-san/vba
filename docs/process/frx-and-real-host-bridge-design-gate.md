# `.frx` and Real Host Bridge Design Gate

## 結論

- `.frx` 配置オブジェクト解析と実 Excel host bridge は、どちらも product code 実装前に設計ゲートを通す。
- `.frx` は対象 binary format と取得 metadata を固定するまで解析しない。
- 実 host bridge は Excel host との接続方式を固定するまで実装しない。
- 既存の `ActiveWorkbookIdentitySnapshot` schema は利用可能な contract だが、実 host への接続方式そのものは未決として扱う。

## 現行で確認済みの前提

- 要件には `.frx` フォームバイナリと配置オブジェクト解析が残っている。
- `resources/vbac/vbac.wsf` / extract / combine の source component 対象には `.frx` が含まれる。
- worksheet / chart control metadata は workbook package から生成した sidecar を static input として扱っている。
- 既存 docs では、`.frm` / `.frx` は UserForm には有効だが、worksheet / chart sheet 上の ActiveX control inventory source とは別であると整理している。
- active workbook identity は host -> extension -> server の snapshot schema と LSP notification までは固定済みである。
- 実 Excel host と extension の接続方式はまだ固定していない。

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
- connection mode 未決のまま COM / script / helper process のいずれかを実装しない。

### 実装前に決めること

1. connection mode
   - extension process から直接接続するか。
   - helper process を起動するか。
   - 既存 `vbac.wsf` と同系統の script bridge にするか。
2. lifecycle
   - VS Code 起動時、command 実行時、file open 時、manual refresh 時のどこで接続するか。
   - Excel が未起動、複数 instance、workbook 切替、Excel 終了をどう扱うか。
3. transport contract
   - ADR 0007 の `ActiveWorkbookIdentitySnapshot` を host 側でどう生成するか。
   - timeout、retry、cancellation、stale snapshot の扱い。
4. user control
   - 自動接続するか、明示 command で接続するか。
   - status / log / error surface をどこに出すか。
5. security / safety
   - workbook や macro を実行しないことをどう保証するか。
   - host bridge が読み取る property を最小化できているか。
6. validation
   - local unit test、mock host test、manual Windows test の境界。
   - extension host test が必要になる条件。

### ゲート解除条件

- connection mode を選ぶ ADR または process doc がある。
- extension / helper / Excel / server の責務境界が固定されている。
- timeout / retry / stale handling / logging が決まっている。
- user-facing resolver を開く条件が ADR 0006 / ADR 0007 と矛盾しない。
- 最小 validation command と、manual test が必要な場合の手順が決まっている。

## PR7 でやること

- この設計ゲートを docs と ADR 入口から参照できるようにする。
- product code、parser、LSP、extension runtime、script は変更しない。

## 次段候補

- `.frx` は、UserForm `.frm` text と `.frx` binary の責務分担を決める design PR を先に作る。
- 実 host bridge は、connection mode を比較して 1 つ選ぶ ADR を先に作る。
- どちらも design PR が通った後、最小 implementation PR へ分ける。
