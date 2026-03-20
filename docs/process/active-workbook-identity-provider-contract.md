# Active Workbook Identity Provider Contract

## 結論

- runtime の active workbook identity は、host / extension / server で共通の snapshot schema `ActiveWorkbookIdentitySnapshot` として扱う。
- extension が host bridge を担当し、server には custom LSP notification `vba/activeWorkbookIdentity` で snapshot を通知する。
- host が返す値は raw の `ActiveWorkbook.FullName` / `Name` / `Path` / `IsAddin` とし、manifest の v1 matching rule への正規化は server 側で行う。
- `Application.ActiveWorkbook` が `Nothing` の場合、Protected View の場合、unsaved workbook、add-in workbook は、すべて別 state として transport し、broad root resolver は無効のまま維持する。
- broad root を将来 user-facing に開くのは、`available` snapshot と `workbook-binding.json` の match がそろったときだけとする。

## 確認した公式ソース

### Office VBA

- [Application.ActiveWorkbook property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.activeworkbook)
  - active window の workbook を返す。
  - window が無い場合や Info / Clipboard window が active の場合は `Nothing` を返す。
  - active Protected View window の document はこの property では取れず、`ProtectedViewWindow.Workbook` を使う必要がある。
- [Application.ActiveProtectedViewWindow property (Excel)](https://learn.microsoft.com/office/vba/api/excel.application.activeprotectedviewwindow)
  - active Protected View window が無い場合は `Nothing` を返す。
- [ProtectedViewWindow.Workbook property (Excel)](https://learn.microsoft.com/office/vba/api/excel.protectedviewwindow.workbook)
  - Protected View workbook へはアクセスできるが、`Workbooks` collection の member ではなく、許可されない操作は error になる。
- [Workbook.FullName property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.fullname)
  - path を含む workbook 名を返す。
- [Workbook.Saved property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.saved)
  - 一度も保存されていない workbook は `Path` が空文字列になる。
- [Workbook.IsAddin property (Excel)](https://learn.microsoft.com/office/vba/api/excel.workbook.isaddin)
  - add-in workbook は window visibility や macro visibility が通常 workbook と異なる。

## 現行前提

- current product は `ThisWorkbook.Worksheets("Sheet1")` のような workbook identity を静的に固定できる root だけを user-facing にしている。
- broad root (`ActiveWorkbook.Worksheets("Sheet1")` / unqualified `Worksheets("Sheet1")`) は、manifest と runtime identity の両方が無い限り閉じたままである。
- `workbook-binding.json` は disk artifact、active workbook identity provider は runtime artifact であり、役割を混ぜない。
- host 実装はまだ存在しないため、今決めるのは contract と gating rule までとする。

## なぜ extension が host bridge を持つのか

| 選択肢 | 利点 | 欠点 | 判断 |
| --- | --- | --- | --- |
| server が host と直接通信する | LSP の外へ出さずに済む | analyzer 層が host lifecycle と retry / timeout を抱える | 不採用 |
| extension が host と通信し server へ通知する | VS Code 側の lifecycle、command、log と整合しやすい | custom notification が 1 本増える | 採用 |
| sidecar / manifest だけで broad root を開く | 実装が軽い | runtime active workbook と bundle の誤結合を防げない | 不採用 |

- extension は VS Code window に属し、将来の refresh command、status 表示、host 再接続と相性がよい。
- server は notification で受け取った snapshot を cache するだけに留めた方が、resolver / diagnostics の責務が明確である。

## 単一 schema にする理由

- host -> extension と extension -> server で別 schema を持つと、state 名や required field がずれやすい。
- broad root gating は field 数が少なく、transport ごとの変換を増やす利点が薄い。
- したがって v1 は 1 つの JSON schema を共有し、extension は transport adapter に徹するのが最も保守しやすい。

## 推奨 schema v1

```json
{
  "version": 1,
  "providerKind": "excel-active-workbook",
  "state": "available",
  "observedAt": "2026-03-21T00:00:00.000Z",
  "identity": {
    "fullName": "C:\\Work\\Book1.xlsm",
    "name": "Book1.xlsm",
    "path": "C:\\Work",
    "isAddin": false
  }
}
```

## State ごとの payload

### 1. `available`

- broad root gating の唯一の候補 state。
- `identity.fullName` / `name` / `path` / `isAddin` を必須とする。
- `isAddin` は `false` でなければならない。`true` なら `unsupported` へ落とす。

### 2. `unavailable`

- `ActiveWorkbook` が `Nothing` のケース、host 未接続、取得失敗をまとめる state。
- `reason` は少なくとも以下を持つ。
  - `no-active-workbook`
  - `host-unreachable`
  - `host-error`
  - `non-workbook-window`
- broad root resolver は常に無効。

### 3. `protected-view`

- `Application.ActiveWorkbook` では取れず、`ActiveProtectedViewWindow` はあるケース。
- `protectedView.sourcePath` / `sourceName` を log 用に持ってよい。
- `ProtectedViewWindow.Workbook` は制約付き object であり、`Workbooks` collection に属さないため、v1 broad root resolver には使わない。

```json
{
  "version": 1,
  "providerKind": "excel-active-workbook",
  "state": "protected-view",
  "observedAt": "2026-03-21T00:00:00.000Z",
  "protectedView": {
    "sourcePath": "C:\\Downloads",
    "sourceName": "Book1.xlsm"
  }
}
```

### 4. `unsupported`

- workbook は取得できたが、manifest binding の対象外であることを示す。
- v1 の `reason` は以下に絞る。
  - `unsaved`
  - `addin`
- `unsaved` は `Workbook.Path = ""` と整合し、`addin` は `Workbook.IsAddin = True` と整合する。

```json
{
  "version": 1,
  "providerKind": "excel-active-workbook",
  "state": "unsupported",
  "reason": "unsaved",
  "observedAt": "2026-03-21T00:00:00.000Z",
  "identity": {
    "fullName": "Book1.xlsm",
    "name": "Book1.xlsm",
    "path": "",
    "isAddin": false
  }
}
```

## なぜ host は raw 値を返し、server が正規化するのか

| 方針 | 利点 | 欠点 | 判断 |
| --- | --- | --- | --- |
| host で正規化して送る | host 側だけで完結して見える | host 実装ごとに比較規則がずれる | 不採用 |
| extension で正規化する | server が軽くなる | extension と server の判定が分離する | 不採用 |
| server で manifest と同じ helper を使って正規化する | match rule を 1 箇所に集約できる | server に snapshot cache が必要 | 採用 |

- ADR [0006](../adr/0006-workbook-binding-policy.md) で決めた v1 matching rule は server 側の resolver が最終的に使う。
- したがって host / extension は raw 値 transport に徹し、`/` -> `\`、case-insensitive、UNC 保持などの規則は server の shared helper に寄せるべきである。

## Transport 境界

### host -> extension

- request / response か event push かは host 実装に委ねる。
- ただし payload schema は `ActiveWorkbookIdentitySnapshot` で固定する。
- host 実装は少なくとも `observedAt` を付与し、extension は stale 判定や log に使えるようにする。

### extension -> server

- custom LSP notification 名は `vba/activeWorkbookIdentity` とする。
- diff / patch ではなく snapshot 全体の置き換え通知にする。
- server は最新 snapshot だけを cache し、document 単位ではなく workspace window 単位の runtime state として扱う。
- cache 前に schema validation を行い、不正 payload は reject して broad root を閉じたままにする。

## Resolver 有効化条件

- 以下をすべて満たすときだけ broad root resolver を有効化する。
  - 最新 snapshot の `state` が `available`
  - `identity.isAddin = false`
  - `identity.path` が空でない
  - current bundle に `workbook-binding.json` が存在する
  - manifest の `workbook.fullName` と snapshot `identity.fullName` が v1 matching rule で一致する

- どれか 1 つでも欠ける場合、resolver は broad root を閉じたままにする。

## Log 方針

### snapshot 受信 log

- `active-workbook-identity.updated`
- `active-workbook-identity.unavailable`
- `active-workbook-identity.protected-view`
- `active-workbook-identity.unsupported`
- `active-workbook-identity.invalid-payload`

### gating log

- `active-workbook-identity.match`
- `active-workbook-identity.mismatch`
- `active-workbook-identity.binding-missing`
- `active-workbook-identity.binding-disabled`

- 受信 log と gating log を分けることで、「host からは来ているが manifest が無い」のか、「host 自体が unavailable」なのかを切り分けやすくする。

## v1 でやらないこと

- `ProtectedViewWindow.Workbook` を broad root 解決へ使う。
- `FullNameURLEncoded` の decode。
- short path / long path 展開、symlink / junction 解決。
- 複数 Excel instance の識別。
- per-document / per-editor ごとの snapshot 保持。

## 今回の完了条件

- host / extension / server の責務分離を決める。
- runtime state として必要な snapshot schema を決める。
- `unavailable` / `protected-view` / `unsupported` を区別し、broad root を開かない条件を固定する。
- manifest v1 matching rule と runtime provider contract を接続する。
- 次段の最小実装候補を `notification + cache + log` に絞る。

## 次段の候補

- extension / server 間の `vba/activeWorkbookIdentity` notification と server cache を read-only で追加する。
- resolver 連携前に、snapshot 受信 log と `binding-missing` / `binding-disabled` の区別だけを観測可能にする。
- `reason` / field 欠損 / 型不整合をどの validation error code へ落とすかを、transport 実装と同じ PR で固定する。
