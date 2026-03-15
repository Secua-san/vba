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

export const dialogSheetPropertyMemberNames = new Set(["DialogFrame"]);
export const dialogSheetControlCollectionMemberNames = new Set(["Buttons", "CheckBoxes", "OptionButtons"]);
export const dialogSheetMethodMemberNames = new Set([
  ...dialogSheetCommonCallableMemberNames,
  ...dialogSheetControlCollectionMemberNames,
]);

export const dialogFramePropertyMemberNames = new Set([
  "Caption",
  "Name",
  "OnAction",
  "Text",
]);

export const dialogFrameMethodMemberNames = new Set(["Select"]);

const dialogSheetControlItemCommonPropertyMemberNames = ["Caption", "Name", "OnAction", "Text"];
const dialogSheetControlCollectionPropertyMemberNames = new Set(["Count"]);
const dialogSheetControlCollectionMethodMemberNames = new Set(["Item"]);

export const dialogSheetControlCollectionOwnerConfigs = [
  {
    collectionLearnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.buttons?view=excel-pia",
    collectionName: "Buttons",
    itemLearnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.button?view=excel-pia",
    itemMethodMemberNames: new Set(["Select"]),
    itemName: "Button",
    itemPropertyMemberNames: new Set(dialogSheetControlItemCommonPropertyMemberNames),
  },
  {
    collectionLearnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.checkboxes?view=excel-pia",
    collectionName: "CheckBoxes",
    itemLearnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.checkbox?view=excel-pia",
    itemMethodMemberNames: new Set(["Select"]),
    itemName: "CheckBox",
    itemPropertyMemberNames: new Set([...dialogSheetControlItemCommonPropertyMemberNames, "Value"]),
  },
  {
    collectionLearnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.optionbuttons?view=excel-pia",
    collectionName: "OptionButtons",
    itemLearnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.optionbutton?view=excel-pia",
    itemMethodMemberNames: new Set(["Select"]),
    itemName: "OptionButton",
    itemPropertyMemberNames: new Set([...dialogSheetControlItemCommonPropertyMemberNames, "Value"]),
  },
];

export const supplementalInteropOwners = [
  {
    kind: "object",
    learnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.dialogsheet?view=excel-pia",
    name: "DialogSheet",
    sections: [
      {
        memberAllowList: dialogSheetPropertyMemberNames,
        sectionName: "Properties",
      },
      {
        memberAllowList: dialogSheetMethodMemberNames,
        memberTypeOverrides: new Map(
          [...dialogSheetControlCollectionMemberNames].map((memberName) => [memberName, memberName]),
        ),
        sectionName: "Methods",
      },
    ],
    title: "DialogSheet object",
  },
  {
    kind: "object",
    learnUrl: "https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel.dialogframe?view=excel-pia",
    name: "DialogFrame",
    sections: [
      {
        memberAllowList: dialogFramePropertyMemberNames,
        sectionName: "Properties",
      },
      {
        memberAllowList: dialogFrameMethodMemberNames,
        sectionName: "Methods",
      },
    ],
    title: "DialogFrame object",
  },
  ...dialogSheetControlCollectionOwnerConfigs.flatMap((config) => [
    {
      kind: "collection",
      learnUrl: config.collectionLearnUrl,
      name: config.collectionName,
      sections: [
        {
          memberAllowList: dialogSheetControlCollectionPropertyMemberNames,
          sectionName: "Properties",
        },
        {
          memberAllowList: dialogSheetControlCollectionMethodMemberNames,
          // Item 自体の typeName は collection owner に残し、literal selector のときだけ
          // built-in root 解決側で item owner へ降ろす。raw メタデータと運用正規化を分離する。
          memberTypeOverrides: new Map([["Item", config.collectionName]]),
          sectionName: "Methods",
        },
      ],
      title: `${config.collectionName} collection`,
    },
    {
      kind: "object",
      learnUrl: config.itemLearnUrl,
      name: config.itemName,
      sections: [
        {
          memberAllowList: config.itemPropertyMemberNames,
          sectionName: "Properties",
        },
        {
          memberAllowList: config.itemMethodMemberNames,
          sectionName: "Methods",
        },
      ],
      title: `${config.itemName} object`,
    },
  ]),
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

export const supplementalOwnerMemberOverrides = [
  {
    members: [
      {
        name: "Item",
        // OLEObjects.Item(1) / Item(i + 1) を collection marker 解決へ流すため、
        // raw doc の `Object` ではなく collection owner 名を正本にする。
        typeName: "OLEObjects",
      },
    ],
    ownerName: "OLEObjects",
    sectionName: "Methods",
  },
  {
    members: [
      {
        name: "Item",
        // Shapes.Item(1) / Item(i + 1) も indexed access の種類に応じて
        // Shape owner へ降ろしたいため、raw doc の型欠落は collection owner 名で補う。
        typeName: "Shapes",
      },
    ],
    ownerName: "Shapes",
    sectionName: "Methods",
  },
];
