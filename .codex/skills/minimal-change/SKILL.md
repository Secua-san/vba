---
name: minimal-change
description: Work guard skill for keeping implementation changes minimal in this VBA VS Code extension repository. Use before and during any code or documentation edit when Codex must limit scope, avoid unrelated changes, avoid refactoring, and keep diffs small for packages/core, packages/server, packages/extension, scripts, docs, or root project files.
---

# Minimal Change

## Purpose
Act as a work guard skill that makes the smallest change satisfying the approved task in this repository. Protect the existing parser, language server, extension, scripts, and docs from unrelated churn.

## Workflow
1. Identify the approved task and restate the smallest observable outcome.
2. List the target files before editing.
3. Define the non-goals:
   - no unrelated fixes
   - no naming-only changes
   - no formatting-only changes
   - no broad refactor
   - no new abstraction unless explicitly approved
4. Inspect only the files needed to understand the target behavior.
5. Edit the fewest files and lines possible.
6. Preserve existing APIs, exported names, AST shapes, diagnostics behavior, fixtures, and test structure unless the task explicitly requires changing them.
7. After editing, review `git diff -- <target files>` and remove accidental churn.
8. Report changed files, the reason each file changed, and any remaining risk.

## Repository-Specific Guardrails
- Prefer existing patterns in `packages/core`, `packages/server`, `packages/extension`, and `scripts`.
- Do not mix product code changes with process/docs cleanup unless the user asked for both.
- Do not update `PLAN.md`, `TASKS.md`, or `TASKLOG.md` unless the task explicitly includes task tracking or the implementation directly requires it.
- Do not touch generated or bulky outputs such as `dist/`, `node_modules/`, `.vscode-test/`, or extension packages unless explicitly requested.
- Keep one logical change per task. If the fix expands into a second concern, stop and propose a follow-up.

## Stop Conditions
- Required behavior is unclear.
- The smallest change would require a new abstraction or broad refactor.
- The change affects more packages than originally scoped.
- A file has unrelated user changes that make a minimal edit risky.

When stopped, ask for confirmation or leave a TODO instead of widening the implementation.
