# workbook root family の server mirror 共通 policy を他 family へ持ち出す前提条件

## 結論

- [workbook-root-family-server-mirror-policy.md](./workbook-root-family-server-mirror-policy.md) は、現時点では `workbook root family` 専用の policy として維持する。
- `worksheet/chart control` 系や `DialogSheet` 系へ同じ物差しを直ちに持ち出さない。
- 他 family へ持ち出す前に必要なのは、「server と extension が同じ family table / canonical anchor source / route taxonomy を共有していること」であり、実装機能の類似だけでは足りない。

## 目的

`workbook root family` では、

- shared case spec が `test-support/workbookRootFamilyCaseTables.cjs` にある
- `scopes` によって server / extension の消費先を 1 表で持てる
- `route-specific gap` と `surface duplication` を同じ family vocabulary で比較できる

という前提が整っている。  
このメモでは、その共通 policy を他 family へ流用できる条件と、まだ流用しない理由を固定する。

## workbook root family 側でそろっている前提

### 1. family canonical source がある

- [test-support/workbookRootFamilyCaseTables.cjs](../../test-support/workbookRootFamilyCaseTables.cjs) が、`worksheetBroadRoot` と `applicationWorkbookRoot` の canonical anchor source になっている。
- shared spec に `scopes`、`state`、`reason`、`route` があり、server / extension は adapter だけを local に持つ。

### 2. route taxonomy が resolver 分岐と 1 対 1 に近い

- root family
  - `ThisWorkbook`
  - `ActiveWorkbook`
  - `Application.ThisWorkbook`
  - `Application.ActiveWorkbook`
  - `Worksheets("...")`
  - `Application.Worksheets("...")`
- selector kind
  - literal sheet name
  - numeric selector
  - code-name selector
  - shadowed root
- route family
  - `OLEObjects(...).Object`
  - `Shapes(...).OLEFormat.Object`
  - root `.Item("...")`

この vocabulary がそのまま server 側の gating / resolver / fixture anchor に対応している。

### 3. mirror 追加コストが小さい

- scope 追加
- server fixture への anchor 追加
- 既存 helper 再利用

で閉じるため、`route-specific gap` を閉じる mirror と `surface duplication` の切り分けが意味を持つ。

## 他 family へ持ち出すための前提条件

### 1. shared spec が先にあること

最低限、次が必要になる。

- family 単位の canonical anchor source
- `scopes` を持つ shared table
- package-local adapter と shared spec の責務分離

shared spec が無い状態では、`shared spec に残すか` と `server に mirror するか` を分ける意味が薄い。  
まず shared 正本を置くかどうかの論点が先に立つ。

### 2. root / selector / route / state の vocabulary が family 内で固定されていること

workbook root family では、同じ family table の中で

- root
- selector
- route
- state
- API surface

を並べて比較できる。  
他 family でも同じ粒度で vocabulary が固定されていないと、`route-specific gap` と `surface duplication` の区別が曖昧になる。

### 3. mirror 候補が「同じ family の residual slice」であること

他 family にも server-only / extension-only の差分はある。  
ただしそれが

- metadata source 未整備
- sidecar / runtime gating 未実装
- feature 入口そのものの未整理

のような前段課題なら、mirror policy の論点ではない。  
比較対象は「同じ family table に載る residual slice」である必要がある。

### 4. mirror 追加が helper / schema 増設を伴わないこと

他 family へ持ち出すなら、mirror 追加コストも workbook root family と同様に軽い必要がある。  
たとえば次が必要になるなら、共通 policy 適用より設計整理を先に行うべきである。

- 新しい generator
- sidecar schema 変更
- host bridge 追加
- family 専用 helper / local table 新設

### 5. server と extension の観測対象が同じ user-facing family を見ていること

workbook root family は、server unit test と extension E2E が同じ `workbook root` family を見ている。  
他 family では、server 側が resolver 単位、extension 側が host / sidecar / metadata integration 単位を見ていて、まだ同一 family と言い切れない場合がある。

## 現在の候補 family の評価

### `worksheet/chart control` 系

現状では、最も近い候補だが、まだ前提が足りない。

- `OLEObjects`
- `Shapes`
- `OLEObject.Object`
- `Shape.OLEFormat.Object`
- `Sheet1.ControlCodeName`

は user-facing には同じ control 支援に見えるが、docs と実装では

- root family
- sidecar lookup の要否
- shape name / code name
- broad root / workbook-qualified root / document module root

が別論点として積み上がっている。  
shared case table もまだ無いため、`route-specific gap` と `surface duplication` を同じ family table の residual slice として比較できない。

### `DialogSheet` control collection 系

`Buttons` / `CheckBoxes` / `OptionButtons` は selector 正規化の論点が中心で、server mirror 拡張の論点にはまだ乗っていない。

- literal-only selector 正規化
- collection owner と item owner の切り分け
- interop 補助ソース由来の surface 制限

が主であり、shared spec / scope 非対称 / residual slice の整理までは進んでいない。  
したがって workbook root family の mirror policy をそのまま流用する段階ではない。

### sidecar / workbook binding / metadata probe 系

これらは test family ではなく基盤 artifact / runtime contract の論点である。  
server と extension の差分があっても、それは mirror coverage の差ではなく input source / integration boundary の差であるため、共通 policy の適用対象にしない。

## 判断

### 他 family へは、shared spec 整理より先に持ち出さない

現段階では、`worksheet/chart control` 系を含む他 family に対して

- `route-specific gap`
- `surface duplication`

の物差しを直接適用しない。  
先に、その family に shared spec と vocabulary が必要かどうかを整理する。

### 最初の移植候補は「worksheet/chart control の shared spec 候補整理」

他 family で最初に見るべきなのは、個別 route の mirror 可否ではなく、次のどこまでを 1 family table と見なせるかである。

- `OLEObjects(...).Object`
- `Shapes(...).OLEFormat.Object`
- `Sheet1.ControlCodeName`
- workbook-qualified root / broad root / document module root のどこまでを同一 family とするか

ここが定まらない限り、mirror policy だけを持ち出しても判断単位が揺れる。

## 次に見るべき点

1. `worksheet/chart control` 系に、family canonical source を切り出す価値があるか  
2. 切り出すなら、root / selector / route / state を workbook root family と同じ粒度で並べられるか  
3. sidecar 有無や metadata source 差を、family residual slice ではなく input 前提として外出しできるか  
4. shared spec 化した場合に、server mirror 候補が「scope 更新 + 最小 anchor 追加」で閉じるか  

## 今やらないこと

- `worksheet/chart control` 系へ `route-specific gap` / `surface duplication` をそのまま適用する
- `DialogSheet` control collection 系へ同じ policy 名で docs を増やす
- shared spec が無い family に対して、mirror policy だけ先に正本化する
- 基盤 artifact 系の差分を mirror coverage の論点として扱う

## 関連文書

- 正本 policy: [workbook-root-family-server-mirror-policy.md](./workbook-root-family-server-mirror-policy.md)
- shared spec 境界: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
- 近い候補 family の docs: [worksheet-chart-control-entrypoint-feasibility.md](./worksheet-chart-control-entrypoint-feasibility.md), [worksheet-chart-shapes-root-feasibility.md](./worksheet-chart-shapes-root-feasibility.md), [worksheet-control-metadata-sidecar-artifact.md](./worksheet-control-metadata-sidecar-artifact.md)
