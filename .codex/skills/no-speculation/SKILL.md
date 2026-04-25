---
name: no-speculation
description: Work guard skill for preventing guessed implementation in this VBA VS Code extension repository. Use whenever requirements, VBA behavior, parser semantics, diagnostics behavior, test expectations, file ownership, or acceptance criteria are unclear, and before filling gaps with assumptions or placeholder behavior.
---

# No Speculation

## Purpose
Act as a work guard skill that avoids guessed behavior and placeholder implementation. Codex must either verify the repository facts, ask the user, or record a TODO instead of inventing behavior.

## Workflow
1. Separate known facts from unknowns.
2. Verify facts from repository sources before editing:
   - `AGENTS.md` for project rules
   - `PLAN.md` and `TASKS.md` for current roadmap context
   - relevant files under `packages/core`, `packages/server`, `packages/extension`, `scripts`, or `docs`
   - existing tests and fixtures near the target behavior
3. For each unknown, choose one:
   - ask the user if the answer controls implementation behavior
   - leave a precise TODO if the user allowed deferred work
   - narrow the task to behavior already proven by existing code
4. Do not create temporary, fake, or placeholder logic to make tests pass.
5. Do not infer Excel/VBA semantics from memory when a local requirement, ADR, fixture, or existing parser behavior can be checked.
6. Before finalizing, list any assumptions that remain.

## Repository-Specific Unknowns That Require Care
- VBA grammar and parser behavior in `packages/core`
- AST node shape, range/text compatibility, diagnostics, references, rename, semantic tokens, and formatter behavior
- Excel workbook, worksheet, control metadata, or `vbac.wsf` integration behavior
- VS Code extension host behavior in `packages/extension`
- External MCP retry/rate-limit behavior

## Forbidden Patterns
- "Likely", "probably", or "should be fine" implementation without verification.
- Adding a broad fallback that hides unclear behavior.
- Creating an abstraction to cover cases that were not requested or verified.
- Updating tests to match guessed behavior.
- Treating a docs-only rule as product behavior unless the task explicitly asks for rule documentation.

## Output
When this skill is used, include:
- confirmed facts used for the change
- unknowns that were resolved
- remaining unknowns, if any
- whether any TODO was left instead of implementation
