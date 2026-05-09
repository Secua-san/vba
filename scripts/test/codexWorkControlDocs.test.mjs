import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");

function readRepoFile(relativePath) {
  return readFileSync(path.join(REPO_ROOT, relativePath), "utf8");
}

test("Codex safety guard documents implementation and test selection checkpoints", () => {
  const safetyGuard = readRepoFile("docs/codex-safety-guard.md");

  for (const requiredText of [
    "対象ファイル",
    "変更理由",
    "影響範囲",
    "最小変更案",
    ".codex/skills/minimal-change",
    ".codex/skills/no-speculation",
    ".codex/skills/test-budget",
    "npm run test --workspace vba-extension",
    "明示指示時だけ実行する",
    "簡易自己レビュー",
    "reviewer",
    "ユーザーが PR 前 full gate を明示した場合",
    "npm run lint",
    "npm test",
    "npm run test:host",
    "通常 PR"
  ]) {
    assert.match(safetyGuard, new RegExp(requiredText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  }

  assert.match(
    safetyGuard,
    /PR 前は `reviewer` の自己レビューを行う。ユーザーが PR 前 full gate を明示した場合は `npm run lint`、`npm test`、`npm run test:host` を通す/
  );
});

test("automation policy separates commit gates from PR gates and heavy tests", () => {
  const automationPolicy = readRepoFile("docs/process/automation-policy.md");

  for (const requiredText of [
    "通常コミット前",
    "PR 作成前（ユーザーが full gate を明示した場合）",
    "npm run lint",
    "npm test",
    "npm run test:host",
    "重いテストの分類",
    "npm run test --workspace vba-extension",
    "E2E / VS Code host",
    "`AGENTS.md` の明示指示",
    "PR 前 full gate"
  ]) {
    assert.match(automationPolicy, new RegExp(requiredText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  }

  assert.match(
    automationPolicy,
    /これらは `AGENTS\.md` の明示指示がある場合、またはユーザーが PR 前 full gate を明示した場合だけ実行する/
  );
});

test("AGENTS keeps the canonical heavy-test command names", () => {
  const agents = readRepoFile("AGENTS.md");

  assert.match(agents, /npm run test --workspace vba-extension/);
  assert.match(agents, /npm run test:host/);
  assert.match(agents, /npm test/);
});
