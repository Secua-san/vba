# ADR 0002: vbac Integration Safety

## Status

Proposed

## Context

vbac integration requires safe command execution and workspace file handling.

## Decision

Wrap vbac calls in extension commands with explicit input/output paths, validation, and clear error messages.

## Consequences

- Reduced risk of unsafe overwrite
- Better diagnostics for command failures
