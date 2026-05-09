# VBA Extension

Language support for Excel VBA in VS Code.

## vbac commands

- `VBA: Extract Source with vbac` (`vba.extract`) selects an Excel VBA workbook and a vbac source root, then writes source under `<source root>/<workbook file name>`. If that folder already exists, it is backed up before replacement.
- `VBA: Combine Source with vbac` (`vba.combine`) selects an Excel VBA workbook and a vbac source root containing `<source root>/<workbook file name>`, confirms overwrite, backs up the workbook, runs vbac on a temporary copy, verifies by re-extracting the combined workbook, then replaces the selected workbook.
- Both commands require Windows `cscript.exe`, write logs under `.vscode-vba/logs`, and write backups under `.vscode-vba/backups`.

## Third-party

- `ariawase` ([vbaidiot/ariawase](https://github.com/vbaidiot/ariawase)) is MIT-licensed.
- License text: <https://github.com/vbaidiot/ariawase/blob/master/LICENSE.txt>
