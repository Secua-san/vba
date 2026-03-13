# ADR 0004: DialogSheet Document Module Policy

## Status

Proposed

## Context

- `Worksheet` と `Chart` の document module root は、Office VBA の object page、`CodeName` page、`VB_Base` GUID を根拠に built-in owner へ接続できる。
- 一方で `DialogSheet` は、Office VBA 側に object page が無く、ローカル参照 JSON にも owner として入っていない。
- Microsoft Learn の Excel 概念記事 `Refer to Sheets by Name` では `DialogSheets("Dialog1").Activate` が示され、dialog sheet 自体は VBA から参照可能であることを確認できる。
- Microsoft Learn の .NET interop `DialogSheet` page には property / method 一覧があるが、ページ自体は `Reserved for internal use.` とされ、`_Dummy*` を含むため、そのまま VBA 補完データの正本にはしづらい。
- Windows registry の `HKEY_CLASSES_ROOT\\Excel.Sheet\\CLSID` は `{00020830-0000-0000-C000-000000000046}` で、document module 側の `VB_Base` と照合できる。

## Decision

- `VB_PredeclaredId = True` かつ `VB_Base = 0{00020830-0000-0000-C000-000000000046}` の class module を、現時点では built-in owner に昇格しない。
- `DialogSheet` document module は workspace symbol としては扱うが、completion / signature help / hover / semantic token の built-in member 解決は保守動作のまま維持する。
- 今後 `DialogSheet` を公開するときは、以下のどちらかを満たすことを前提にする。
  - Office VBA 側に安定した object page / member list が追加される
  - interop page を二次ソースとして取り込む専用ルールを設計し、`Reserved for internal use` と `dummy` member を除外する基準が固まる

## Consequences

- `DialogSheet1.` では誤って `Worksheet` や `Chart` の member を補完しない。
- `DialogSheet` の built-in member 支援は未提供のままだが、誤案内のリスクを抑えられる。
- 将来対応する場合は、Microsoft Learn 参照生成フローに interop 補助ソースをどう統合するかを先に決める必要がある。
