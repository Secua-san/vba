# vscode-vba-extension

Monorepo for a VS Code extension focused on Excel VBA editing support.

## Documentation

- [Documentation Guide](./docs/README.md)
- [Task Summary](./TASKS.md)
- [Task Log](./TASKLOG.md)
- [Process Guide](./docs/process/README.md)
- [Requirements Overview](./docs/requirements/000-overview.md)

## Packages

- `packages/core`: lexer/parser/AST/symbol/types/diagnostics
- `packages/server`: Language Server (LSP)
- `packages/extension`: VS Code client extension

## Development Host

Run `npm run dev:host` to build the workspace and open an Extension Development Host against `packages/extension/test/fixtures`.

If you open the repository in VS Code, press `F5` and use `Run VBA Extension`. Use `Run VBA Extension Tests` when you want to execute the extension test suite from the debugger.

## Third-party Licenses

Third-party license notices are listed in [THIRD_PARTY_LICENSES.md](./THIRD_PARTY_LICENSES.md).
