---
name: test-budget
description: Work guard skill for selecting the smallest relevant validation set in this VBA VS Code extension repository. Use before running tests, builds, lint, or extension host checks, especially to avoid full repository tests, avoid E2E unless explicitly approved, and minimize test time based on changed files.
---

# Test Budget

## Purpose
Act as a work guard skill that runs only the validation needed for the changed files. Avoid full test suites and extension host tests unless the user explicitly asks for them.

## Workflow
1. List changed files with `git diff --name-only` or the planned target files.
2. Map each changed file to the smallest relevant validation command.
3. Prefer single-package or single-file checks over root commands.
4. Do not run full repository tests unless the user explicitly requests them.
5. Do not run E2E / VS Code extension host tests unless the user explicitly requests them.
6. If no test is appropriate for docs-only or rule-only changes, run `git diff --check -- <files>` and report tests as not run with the reason.
7. Report executed commands, skipped commands, and why skips were correct.

## Test Command Map
- Root scripts:
  - `scripts/` changes: `npm run test:scripts`
  - Narrow script test: `node --test scripts/test/<file>.test.mjs`
- Core package:
  - `packages/core/` code or test changes: `npm run test --workspace @vba/core`
  - Build-only when tests are unrelated but types matter: `npm run build --workspace @vba/core`
- Server package:
  - `packages/server/` code or test changes: `npm run test --workspace @vba/server`
  - Narrow server test: `node --test packages/server/test/<file>.test.js`
  - Build-only when tests are unrelated but bundle output matters: `npm run build --workspace @vba/server`
- Extension package:
  - `packages/extension/` code changes: start with `npm run build --workspace vba-extension`
  - `npm run test --workspace vba-extension` requires explicit user approval because it builds dependencies and runs VS Code tests.
  - `npm run test:host` requires explicit user approval because it starts the VS Code extension host.
- Documentation or rule-only changes:
  - `git diff --check -- <changed-doc-files>`

## Forbidden Commands Without Explicit User Approval
- `npm test`
- `npm run test`
- `npm run test --workspace vba-extension`
- `npm run test:host`
- Any command that launches VS Code or the extension host
- Any broad command chosen only "just in case"

## Budget Rules
- Add tests only for the changed behavior and only the minimum cases needed.
- Do not add "just in case" tests.
- Do not update unrelated fixtures or snapshots.
- If one focused test proves the changed behavior, stop there.
- If the focused test reveals wider breakage, report it before expanding validation.
