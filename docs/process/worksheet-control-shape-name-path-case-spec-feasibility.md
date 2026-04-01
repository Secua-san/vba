# worksheet control shapeName path の dedicated case spec 抽出方針

## 結論

- `worksheetControlShapeNamePath` の dedicated case spec は、fixture ごとの table ではなく family 単位の `test-support/worksheetControlShapeNamePathCaseTables.cjs` を v1 正本にする。
- ただし初回抽出は `completion` から始め、`hover` / `signature` / `semantic` は package-local adapter と anchor topology を見ながら後続で広げる。
- v1 の common field は `fixture` / `anchor` / `rootKind` / `routeKind` / `scopes` に固定し、negative entry だけ `reason` を追加する。`matched/closed` は従来どおり `rootKind` 側へ残し、別の `state` 軸は足さない。
- `OleObjectBuiltIn.bas` と `ShapesBuiltIn.bas` は route-local execution source のまま維持し、family 専用 mixed fixture は case spec 抽出後の drift が見えてから再評価する。

## 目的

[worksheet-control-shape-name-path-vocabulary-feasibility.md](./worksheet-control-shape-name-path-vocabulary-feasibility.md) で、`worksheetControlShapeNamePath` の vocabulary と canonical anchor source の考え方は固定した。  
今回の論点は、その次段として

- dedicated case spec を family / fixture / route のどこで切るか
- 初回にどの interaction slice から shared 化するか
- `test-support/` へ持ち込む field をどこまでに絞るか

を整理することにある。

## 入力として使う正本

- family vocabulary: [worksheet-control-shape-name-path-vocabulary-feasibility.md](./worksheet-control-shape-name-path-vocabulary-feasibility.md)
- shared case spec の共通 policy: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
- route-local fixture:
  - [OleObjectBuiltIn.bas](../../packages/extension/test/fixtures/OleObjectBuiltIn.bas)
  - [ShapesBuiltIn.bas](../../packages/extension/test/fixtures/ShapesBuiltIn.bas)

## 比較した切り方

### 案1. route ごとに別 table を持つ

例:

- `worksheetControlOleObjectCaseTables`
- `worksheetControlShapeOleFormatCaseTables`

この案の利点は fixture と 1 対 1 で結び付けやすいことだが、`worksheet owner + shapeName + sidecar 一致` という family の主語が route 単位に割れてしまう。  
`ole-object` と `shape-oleformat` を別 table にすると、後から同じ `rootKind` / `reason` を 2 箇所に持つことになり、family 単位で drift を見づらい。

### 案2. family 単位で 1 つの case spec を持ち、entry が `fixture` を指す

例:

- `test-support/worksheetControlShapeNamePathCaseTables.cjs`
- top-level key: `worksheetControlShapeNamePath`

この案なら vocabulary の主語を family に保ったまま、entry ごとに `fixture` で `OleObjectBuiltIn.bas` / `ShapesBuiltIn.bas` を切り替えられる。  
`routeKind` と `fixture` を両方持つことで、「family では同じ論点だが execution source は 2 本」という現状を素直に表現できる。

### 案3. 先に family 専用 mixed fixture を作ってから table 化する

この案は case spec の `fixture` field を省ける可能性があるが、現時点では

- route-local regression と family canonical anchor の二重管理が増える
- generic `OLEObject` / `Shape` / `OLEFormat` surface の既存回帰をどう分離するか追加判断が要る
- まだ case spec なしの段階で fixture だけ先に再編することになる

ため、初手としては重い。

## 観察結果

### 1. 切る単位は `fixture` ではなく family が自然

- `rootKind=document-module/workbook-qualified-static/workbook-qualified-matched/workbook-qualified-closed`
- `routeKind=ole-object/shape-oleformat`
- `reason=numeric-selector/dynamic-selector/code-name-selector/plain-shape/chartsheet-root/non-target-root`

という vocabulary は、すでに family 単位で固定されている。  
この段階で table を fixture ごとに分けると、family vocabulary の正本が分散する。

### 2. `fixture` は common field として持つ方が安い

`worksheetControlShapeNamePath` は最初から 2 本の route-local fixture を参照するので、`fixture` を entry の common field に含める方が境界が明瞭である。  
`routeKind` だけで fixture を暗黙決定すると、「将来 1 route を複数 fixture へ分けたとき」に schema の拡張が必要になる。

### 3. `direct` / `.Item("...")` は独立軸にしない

`Sheet1.OLEObjects("CheckBox1").Object` と `Sheet1.OLEObjects.Item("CheckBox1").Object` の違いは、resolver family の語彙というより anchor の差である。  
ここで `selectorKind=direct/item` のような軸を足すと、`rootKind` / `routeKind` / `reason` とは別に schema を増やすわりに、server / extension adapter の期待値がほとんど減らない。

したがって v1 では

- `anchor` に direct / item の両形をそのまま持つ
- server slice の違いは `scopes` 側で表す

方がレビューしやすい。

### 4. 初回抽出は `completion` が最も軽い

`hover` / `signature` / `semantic` まで一度に shared 化すると、初回から

- `identifier`
- `tokenKind`
- partial token anchor
- package-local payload assertion

の境界も同時に決める必要がある。

一方 `completion` は、

- `anchor`
- `rootKind`
- `routeKind`
- `reason`
- `scopes`

だけで family coverage の主語をかなり表現できる。  
また、`ole-object` / `shape-oleformat` の両 route と `document-module` / `static` / `matched` / `closed` を一通り踏めるため、case spec の切り方を検証する初回 slice として扱いやすい。

## 採用する抽出方針

### 1. dedicated case spec は family 単位で 1 file に置く

v1 正本:

- `test-support/worksheetControlShapeNamePathCaseTables.cjs`

想定 top-level:

```js
{
  worksheetControlShapeNamePath: {
    completion: {
      positive: [],
      negative: []
    }
  }
}
```

`hover` / `signature` / `semantic` は、初回抽出後に同じ file へ段階追加する。

### 2. v1 common field は 5 つに固定する

全 entry に持つ field:

- `fixture`
- `anchor`
- `rootKind`
- `routeKind`
- `scopes`

negative entry にだけ追加する field:

- `reason`

初回の `completion` では、これ以上の field を増やさない。

### 3. `state` / `selectorKind` / `fixtureVariant` は足さない

v1 で持たないもの:

- `state`
  - `matched/closed` は `rootKind` に残す
- `selectorKind`
  - direct / `.Item("...")` は `anchor` に埋め込む
- `fixtureVariant`
  - `fixture` path をそのまま正本にする
- `occurrenceIndex`
  - 初回 `completion` では duplicate anchor 問題が主論点ではないため先回りしない

これらは、shared 化した後に drift や duplicate anchor が実際に発生したときだけ再評価する。

### 4. mixed fixture はまだ導入しない

この段階では

- [OleObjectBuiltIn.bas](../../packages/extension/test/fixtures/OleObjectBuiltIn.bas)
- [ShapesBuiltIn.bas](../../packages/extension/test/fixtures/ShapesBuiltIn.bas)

を execution source として使い続ける。  
family canonical source は case spec 側に寄せ、fixture の再編は次の条件が見えてから判断する。

再評価トリガー:

- completion case spec 抽出後も anchor drift が review の主因になる
- route-local fixture の generic surface と family anchor が頻繁に衝突する
- `hover` / `signature` / `semantic` 追加時に `fixture` field だけでは読みづらくなる

### 5. `completion` 最小 PoC の完了条件を固定する

次タスクでいう「最小 PoC」は、次を満たしたときだけ完了とみなす。

- `routeKind=ole-object` と `routeKind=shape-oleformat` の両方が `completion` table に入っている
- positive は `document-module` / `workbook-qualified-static` / `workbook-qualified-matched` を、両 route を合わせて少なくとも 1 回ずつ表現している
- negative は `rootKind=workbook-qualified-closed` を各 route で少なくとも 1 回持ち、さらに `numeric-selector` / `dynamic-selector` / `code-name-selector` / `plain-shape` / `chartsheet-root` / `non-target-root` の各 `reason` を family 全体で少なくとも 1 回ずつ持つ
- `fixture` は [OleObjectBuiltIn.bas](../../packages/extension/test/fixtures/OleObjectBuiltIn.bas) と [ShapesBuiltIn.bas](../../packages/extension/test/fixtures/ShapesBuiltIn.bas) の両方を参照している
- `scopes` は `extension` と各 route に対応する server 側 scope を少なくとも 1 entry ずつ含み、adapter 配線の片側だけで終わらない

言い換えると、`completion.positive` 数件だけを切り出した状態や、1 route だけを shared 化した状態は「最小 PoC 完了」とは扱わない。

## schema イメージ

```js
{
  worksheetControlShapeNamePath: {
    completion: {
      positive: [
        {
          fixture: "packages/extension/test/fixtures/OleObjectBuiltIn.bas",
          anchor: 'Debug.Print Sheet1.OLEObjects("CheckBox1").Object.',
          rootKind: "document-module",
          routeKind: "ole-object",
          scopes: ["extension", "server-worksheet-control-ole"]
        }
      ],
      negative: [
        {
          fixture: "packages/extension/test/fixtures/ShapesBuiltIn.bas",
          anchor: 'Debug.Print ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value',
          rootKind: "workbook-qualified-static",
          routeKind: "shape-oleformat",
          reason: "code-name-selector",
          scopes: ["extension", "server-worksheet-control-shape"]
        }
      ]
    }
  }
}
```

補足:

- `fixture` は repo root 基準の相対 path に固定し、v1 では次の 2 値だけを許容する
  - `packages/extension/test/fixtures/OleObjectBuiltIn.bas`
  - `packages/extension/test/fixtures/ShapesBuiltIn.bas`
- `scopes` の naming は、初回 PoC で server 側 slice 名をどう切るかを見てから固定する

## 非採用

- route ごとに独立した case spec file を持つ
- family 専用 mixed fixture を先に作り、その fixture を唯一の canonical source にする
- `completion` / `hover` / `signature` / `semantic` を最初から全部一括抽出する
- `direct/item` を schema の独立軸にする
- `matched/closed` を `state` として分離し、`rootKind` を再分解する

## 次段の候補

1. `worksheetControlShapeNamePath.completion` の dedicated case spec を最小抽出する  
2. `server` / `extension` で使う scope 名の最小 vocabulary を PoC する  
3. completion 抽出後の drift を見て、mixed fixture 再評価が必要かを判断する  

## 関連文書

- family 候補の切り出し: [worksheet-control-shared-spec-family-candidate-feasibility.md](./worksheet-control-shared-spec-family-candidate-feasibility.md)
- vocabulary と canonical anchor source: [worksheet-control-shape-name-path-vocabulary-feasibility.md](./worksheet-control-shape-name-path-vocabulary-feasibility.md)
- shared case spec の共通 policy: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
