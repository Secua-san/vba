# ADR 0003: MCP Call Retry Policy

## Status

Proposed

## Context

External MCP-style server calls can hit `429 Too Many Requests`, and ad-hoc retry code makes logging, backoff, and duplicate suppression inconsistent across integrations.

## Decision

Introduce a shared MCP request helper that centralizes:

- `429` detection with `Retry-After` priority
- exponential backoff with jitter when `Retry-After` is absent
- explicit max retry failure
- minimum spacing between calls
- in-flight duplicate suppression by request key
- structured logs including MCP name, retry count, wait time, and final failure reason

## Consequences

- New MCP integrations should reuse the shared helper instead of implementing per-call retry logic
- Call sites must provide a stable MCP name and, when necessary, a stable request key for duplicate suppression
- Retry and rate-limit behavior can be tested once and reused across scripts and future integrations
