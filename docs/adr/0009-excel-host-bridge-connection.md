# ADR 0009: Excel Host Bridge Connection

## Status

Accepted

## Context

- ADR [0007](0007-active-workbook-identity-provider-contract.md) fixes the `ActiveWorkbookIdentitySnapshot` schema and the extension -> server notification contract.
- ADR [0008](0008-frx-and-real-host-bridge-design-gate.md) keeps real Excel host bridge implementation out of product code until the connection mode is selected.
- The first release remains local VSIX distribution for editing support, so host integration must stay explicit, reviewable, and optional.
- The server already consumes validated snapshots and must not own Excel lifecycle or host communication.

## Decision

- The v1 real Excel host bridge connection mode is an extension-owned manual `cscript.exe` helper.
- The future user-facing command name is `vba.refreshActiveWorkbookIdentity`.
- The extension owns helper process launch, timeout, cancellation, output parsing, logging, and notification to the language server.
- The server remains a snapshot consumer only. It never connects to Excel, starts helper processes, or reads host state directly.
- The helper may only produce the ADR 0007 snapshot payload and must not execute workbook macros or mutate workbooks.
- No startup placeholder snapshot is sent. Until a helper run returns a valid snapshot, the server keeps runtime workbook-dependent resolvers closed.
- No automatic polling, long-running helper, or background host bridge is part of v1.
- `.frx` binary parsing is not part of this bridge and remains under the ADR 0008 gate.

## Consequences

- The first implementation PR can add one manual refresh command without changing resolver contracts.
- Host unavailability, Protected View, unsaved workbooks, and add-ins continue to flow through the existing ADR 0007 snapshot states.
- The extension can surface helper failures in the existing output/log path while the server keeps conservative behavior.
- Future automatic refresh or alternate transports require a new ADR or an explicit update to this ADR.
