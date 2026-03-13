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

export const supplementalOwnerMembers = [
  {
    members: [
      {
        learnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.applicationclass.dialogsheets?view=excel-pia",
        name: "DialogSheets",
        // interop page は `As Sheets` を返すが、built-in root 解決では DialogSheet item owner へ接続したい。
        typeName: "DialogSheets",
      },
    ],
    ownerName: "Application",
    sectionName: "Properties",
  },
  {
    members: [
      {
        learnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.workbookclass.dialogsheets?view=excel-pia",
        name: "DialogSheets",
        // Workbook 側も同じ理由で DialogSheets collection owner へ正規化する。
        typeName: "DialogSheets",
      },
    ],
    ownerName: "Workbook",
    sectionName: "Properties",
  },
];
