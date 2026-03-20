# ADR ガイド

`docs/adr/` の入口。設計変更前に、関係する ADR だけを読む。

## ADR 一覧

| ADR | テーマ | 開くタイミング |
| --- | --- | --- |
| [0001 Parser Strategy](0001-parser-strategy.md) | 手書きパーサと解析パイプライン | lexer / parser / AST / symbol / inference を変えるとき |
| [0002 vbac Integration Safety](0002-vbac-integration-safety.md) | `vbac` / XLAM 連携の安全方針 | extract / combine / XLAM 連携を触るとき |
| [0003 MCP Call Retry Policy](0003-mcp-call-retry-policy.md) | 外部 MCP 呼び出しの retry / rate-limit | MCP クライアントや取得スクリプトを触るとき |
| [0004 DialogSheet Document Module Policy](0004-dialogsheet-document-module-policy.md) | DialogSheet document module の扱い | DialogSheet root、owner、公的ソース不足の論点を触るとき |
| [0005 Explicit Sheet-Name Root Policy](0005-explicit-sheet-name-root-policy.md) | `Worksheets("Sheet1")` 系 root と sidecar join key | explicit sheet-name root、`sheetName` / `sheetCodeName` の使い分けを触るとき |
| [0006 Workbook Binding Policy](0006-workbook-binding-policy.md) | broad root 再評価の前提となる workbook identity binding | `ActiveWorkbook` / unqualified `Worksheets`、binding manifest、host identity 受け渡しを触るとき |

## 追加ルール

- 複数パッケージにまたがって長く効く判断は ADR に残す。
- 実装メモや一時的な調査結果は `docs/process/` または `TASKS.md` に置き、ADR へ肥大化させない。
- 既存 ADR を覆す場合は、差分理由を明記した新しい ADR を追加する。
