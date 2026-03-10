export const signatureMemberAllowList = new Map([
  ["Application", new Set(["Calculate", "CalculateFull", "CalculateFullRebuild", "CalculateUntilAsyncQueriesDone"])],
  ["Range", new Set(["Address", "AddressLocal"])],
  [
    "WorksheetFunction",
    new Set([
      "And",
      "Average",
      "Choose",
      "Count",
      "CountA",
      "CountBlank",
      "EDate",
      "EoMonth",
      "Find",
      "HLookup",
      "Index",
      "Lookup",
      "Match",
      "Max",
      "Median",
      "Min",
      "Or",
      "Power",
      "Round",
      "Search",
      "Sum",
      "Text",
      "Transpose",
      "VLookup",
      "Xor",
    ]),
  ],
]);

export const signatureMissingMemberWatchList = new Map([
  ["WorksheetFunction", new Set(["XLookup", "XMATCH"])],
]);
