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
    "npm run test:host",
    "npm test",
    "簡易自己レビュー",
    "reviewer",
    "通常 PR"
  ]) {
    assert.match(safetyGuard, new RegExp(requiredText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  }
});

test("automation policy separates commit gates from PR gates and heavy tests", () => {
  const automationPolicy = readRepoFile("docs/process/automation-policy.md");

  for (const requiredText of [
    "通常コミット前",
    "PR 作成前",
    "npm run lint",
    "npm test",
    "npm run test:host",
    "重いテストの分類",
    "npm run test --workspace vba-extension",
    "E2E / VS Code host",
    "PR 作成前ゲート"
  ]) {
    assert.match(automationPolicy, new RegExp(requiredText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  }
});

test("AGENTS keeps the canonical heavy-test command names", () => {
  const agents = readRepoFile("AGENTS.md");

  assert.match(agents, /npm run test --workspace vba-extension/);
  assert.match(agents, /npm run test:host/);
  assert.match(agents, /npm test/);
});
