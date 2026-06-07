# VBA Extension

Language support for Excel VBA in VS Code.

## First release scope

This release is intended for local VSIX installation and day-to-day editing of exported Excel VBA source files.

Supported source files:

- `.bas`
- `.cls`
- `.frm`

Primary editing features:

- VBA language detection with language id `vba`
- TextMate syntax highlighting
- VBA snippets
- diagnostics for syntax, declarations, duplicate definitions, unused/write-only locals, unreachable code, type mismatch, ByRef risks, missing `PtrSafe`, and unsafe `Long` usage in `Declare PtrSafe`
- completion for local/workspace symbols, built-in Excel/VBA reference data, known members, and known `CreateObject` ProgIDs
- hover, signature help, definition, references, local rename, document symbols, workspace symbols, semantic tokens, and document formatting
- quick fixes for `Option Explicit` and `PtrSafe`

## Install from a local VSIX

From the repository root:

```sh
npm install
npm run package
code --install-extension dist/vba-extension.vsix
```

You can also install the generated `dist/vba-extension.vsix` through VS Code's "Install from VSIX..." command.

After installation, open a folder that contains exported VBA source files. VS Code should recognize `.bas`, `.cls`, and `.frm` files as Excel VBA.

## Smoke check

Use `packages/extension/test/fixtures` or a real exported VBA source folder.

- Open a `.bas`, `.cls`, or `.frm` file and confirm syntax highlighting is active.
- Type a snippet prefix such as `sub`, `function`, or `if` and confirm snippet suggestions appear.
- In a `Declare` statement without `PtrSafe`, confirm the diagnostic and `PtrSafe` quick fix appear.
- Type `Application.` or `WorksheetFunction.` and confirm completion, hover, and signature help are available.
- Use Go to Definition, Find References, Rename Symbol on a local variable, document symbols, workspace symbols, and Format Document.
- Confirm the commands `VBA: Refresh Active Workbook Identity`, `VBA: Extract Source with vbac`, and `VBA: Combine Source with vbac` are listed in the Command Palette. Real Excel refresh and workbook extract/combine are not required for this release smoke check.

## Settings

- `vba.analysis.debounceMs`: debounce delay in milliseconds before the language server re-analyzes a VBA document. Default: `300`.
- `vba.analysis.logPerformance`: opt-in language server console logging for analysis timing, line count, character count, diagnostic count, document version, and trigger. Default: `false`; source text and absolute file paths are not logged.

## Known limits

- This extension targets Excel VBA editing only. It does not execute VBA or replace the VBE runtime.
- The first release is local VSIX distribution. VS Marketplace and GitHub Releases publishing are out of scope.
- Automatic Excel host bridge integration and `.frx` binary object parsing are out of scope.
- vbac commands are available, but production workbook extract/combine validation is tracked separately from the editing-focused release gate.

## Active workbook identity

- `VBA: Refresh Active Workbook Identity` (`vba.refreshActiveWorkbookIdentity`) manually runs a Windows `cscript.exe` helper, reads the active Excel workbook identity, validates the snapshot, and sends it to the language server.
- The command sends workbook identity only after the helper returns a valid, fresh payload. If a non-cancelled refresh fails, it sends an `unavailable` snapshot to close any previously cached active workbook binding. It does not start a background bridge or send startup placeholder snapshots.
- The helper reads workbook identity properties only; it does not execute workbook macros or mutate workbooks.
- If Excel is in Protected View, the helper returns a `protected-view` snapshot and may include source name/path metadata; workbook binding remains disabled.
- To smoke-check the same helper outside VS Code, run `npm run smoke:active-workbook-identity` from the repository root.
- To require an opened workbook match, run `npm run smoke:active-workbook-identity -- --expect-state available --expect-full-name C:\path\Book1.xlsm`.
- To require a Protected View source match, run `npm run smoke:active-workbook-identity -- --expect-protected-source-name Book1.xlsm --expect-protected-source-path C:\Downloads`.

## vbac commands

- `VBA: Extract Source with vbac` (`vba.extract`) selects an Excel VBA workbook and a vbac source root, then writes source under `<source root>/<workbook file name>`. If that folder already exists, it is backed up before replacement.
- `VBA: Combine Source with vbac` (`vba.combine`) selects an Excel VBA workbook and a vbac source root containing `<source root>/<workbook file name>`, confirms overwrite, backs up the workbook, runs vbac on a temporary copy, verifies by re-extracting the combined workbook, then replaces the selected workbook.
- Both commands require Windows `cscript.exe`, write logs under `.vscode-vba/logs`, and write backups under `.vscode-vba/backups`.

## Third-party

- `ariawase` ([vbaidiot/ariawase](https://github.com/vbaidiot/ariawase)) is MIT-licensed.
- License text: <https://github.com/vbaidiot/ariawase/blob/master/LICENSE.txt>
