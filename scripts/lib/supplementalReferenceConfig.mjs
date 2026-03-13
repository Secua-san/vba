export const dialogSheetCommonCallableMemberNames = new Set([
  "Activate",
  "Evaluate",
  "ExportAsFixedFormat",
  "Move",
  "PrintOut",
  "SaveAs",
  "Select",
  "Unprotect",
]);

export const supplementalInteropOwners = [
  {
    kind: "object",
    learnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.dialogsheet?view=excel-pia",
    memberAllowList: dialogSheetCommonCallableMemberNames,
    name: "DialogSheet",
    sectionName: "Methods",
    title: "DialogSheet object",
  },
];

export const supplementalOwnerClones = [
  {
    kind: "collection",
    learnUrl: "https://learn.microsoft.com/office/vba/excel/concepts/workbooks-and-worksheets/refer-to-sheets-by-name",
    memberTypeOverrides: new Map([["Item", "DialogSheet"]]),
    name: "DialogSheets",
    sourceOwnerName: "Sheets",
    title: "DialogSheets collection",
  },
];
