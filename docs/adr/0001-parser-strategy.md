# ADR 0001: Parser Strategy

## Status

Proposed

## Context

VBA parsing must prioritize resilience, incremental value, and diagnostics under broken code.

## Decision

Start with a hand-written parser focused on declaration/procedure boundaries and error recovery.

## Consequences

- Fast iteration for VBA-specific edge cases
- Later migration path to formal grammar remains possible
