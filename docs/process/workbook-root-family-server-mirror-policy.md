# workbook root family の server mirror 拡張条件

## 結論

- workbook root family で `server` scope を増やすのは、`route-specific gap` を閉じるときに限る。
- 既に server 側で同じ root / selector / route family を踏めており、別 API surface で「同じく閉じる」を重ねるだけなら、`surface duplication` とみなし extension-only に残す。
- `shared spec に残すか` と `server でも mirror するか` は分けて判断する。canonical anchor source なら scope 非対称でも shared spec に残してよい。

## 目的

`worksheetBroadRoot` と `applicationWorkbookRoot` では、

- completion negative は server へ mirror したものがある
- hover / signature / semantic の一部は extension-only のまま残したものがある

という判断が並行している。  
このメモでは、family 個別メモで使ってきた物差しを、workbook root family 共通の判断基準として固定する。

## 用語

### route-specific gap

次のいずれかを満たし、server 側でまだ踏めていない経路差分を指す。

- 同じ root / selector / route prefix を server completion / interaction / semantic のどれでもまだ踏めていない
- 既存 server coverage では failure point が `DocumentService` のどの gating か切り分けにくい
- scope 追加と server fixture への最小 anchor 追加だけで、resolver 近傍の回帰を早く止められる

### surface duplication

次の状態を指す。

- root / selector / route gating 自体は server 側で既に別 surface から閉じている
- 追加しても「同じ anchor family が別 API でも閉じる」を重ねる比重が大きい
- user-facing residual slice を見る主戦場が extension E2E で、server に増やしても説明力が大きく増えない

## 共通の判断軸

### 1. mirror 対象は route 単位で見る

`static` / `matched` / `shadowed` の state だけでは決めない。  
次の組み合わせで見る。

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
- API surface
  - completion
  - hover
  - signature
  - semantic

server mirror を増やすのは、「同じ route family を server がまだ踏めていない」と言えるときに限る。

### 2. completion は route-specific gap を閉じやすい

completion negative は、

- built-in root gating
- selector gating
- route 解決の入口

を `DocumentService.getCompletionSymbols()` で直接見る。  
そのため、同じ route prefix が server に無いなら mirror する価値が高い。

既存の該当例:

- [application-workbook-root-completion-server-scope-feasibility.md](./application-workbook-root-completion-server-scope-feasibility.md)
  - `Application.ThisWorkbook.Worksheets(1)` や `.Item("Sheet1")` を含む completion negative 3 本を server へ広げた
- [worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md](./worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md)
  - broad root non-target では completion negative を server 側で shared 化し、root / selector gating の境界を server でも止めている

### 3. interaction / semantic は surface duplication になりやすい

hover / signature / semantic は、次の条件がそろうと extension-only に残す。

- 同じ reason の completion negative が server にある
- positive interaction / semantic anchor か closed-state coverage が server にある
- 追加しても failure point が resolver 近傍へ寄らず、「表示されない」を別 surface で重ねるだけになる

既存の該当例:

- [application-workbook-root-interaction-semantic-server-scope-feasibility.md](./application-workbook-root-interaction-semantic-server-scope-feasibility.md)
  - completion 実装後に残った extension-only `hover` / `signature` / `semantic` 7 本は `surface duplication` 寄りとして server へ広げなかった
- [worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md](./worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md)
  - broad root non-target `hover` / `signature` negative は、completion negative と positive closed-state で十分と判断し server mirror を増やさなかった

### 4. shared spec と server mirror は独立に扱う

canonical anchor source なら、`server` が使わなくても shared spec に残してよい。  
判断軸は次のとおり。

- shared spec に残すか
  - anchor / reason / state が family canonical source か
  - package 固有事情が adapter 層に隔離されているか
- server に mirror するか
  - route-specific gap を閉じるか
  - 追加後の server test が failure point を早く示せるか

この分離は次のメモと整合する。

- [worksheet-broad-root-extension-only-interaction-shared-spec-feasibility.md](./worksheet-broad-root-extension-only-interaction-shared-spec-feasibility.md)
- [application-workbook-root-extension-only-interaction-shared-spec-feasibility.md](./application-workbook-root-extension-only-interaction-shared-spec-feasibility.md)
- [application-workbook-root-extension-only-completion-shared-spec-feasibility.md](./application-workbook-root-extension-only-completion-shared-spec-feasibility.md)

### 5. mirror 追加のコストは「scope 更新 + 最小 anchor 追加」で見積もる

次を超える変更が必要なら、mirror の費用対効果を再確認する。

- shared case table の `scopes` 更新
- 既存 server fixture text への最小 anchor 追加
- 既存 helper の再利用

逆に、helper 新設や schema 拡張が必要なら、単なる coverage 追加ではなく設計変更の論点として切り分ける。

## 判断フロー

1. 追加したい entry が family canonical anchor source か確認する  
2. server 側に同じ root / selector / route family の coverage が既にあるか確認する  
3. 無ければ `route-specific gap` とみなし、completion 優先で server mirror を検討する  
4. 既にあるなら、追加分が failure point を早めるか確認する  
5. 早まらないなら `surface duplication` とみなし、extension-only に残す  
6. どちらの場合も、shared spec から外すかは別軸で判断する

## 再評価トリガー

- workbook root family で、新しい route family や selector kind を追加したとき
- server 側で hover / signature / semantic の実不具合が起き、completion negative だけでは failure point が追えなかったとき
- `reviewer` または CodeRabbit から、同じ family で server mirror の判断揺れを繰り返し指摘されたとき
- helper / schema を増やさずには mirror できない residual slice が増え、この policy だけでは判断しづらくなったとき

## 今やらないこと

- workbook root family 以外へこの物差しを即座に横展開する
- `shared spec に残すか` と `server に mirror するか` を 1 つの真偽値へ潰す
- hover / signature / semantic を一律で server へ mirror する
- reverse に、scope 非対称 entry を一律で local table へ戻す

## 関連文書

- 共通 spec 方針: [workbook-root-family-case-table-policy.md](./workbook-root-family-case-table-policy.md)
- broad root family 側の判断: [worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md](./worksheet-broad-root-server-nontarget-interaction-shared-scope-feasibility.md), [worksheet-broad-root-extension-only-interaction-shared-spec-feasibility.md](./worksheet-broad-root-extension-only-interaction-shared-spec-feasibility.md)
- application workbook root 側の判断: [application-workbook-root-completion-server-scope-feasibility.md](./application-workbook-root-completion-server-scope-feasibility.md), [application-workbook-root-interaction-semantic-server-scope-feasibility.md](./application-workbook-root-interaction-semantic-server-scope-feasibility.md)
