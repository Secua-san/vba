import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { VBA_KEYWORDS } from "../lexer/keywords";
import { normalizeIdentifier } from "../types/helpers";

export type BuiltinCompletionKind = "constant" | "function" | "keyword" | "type" | "variable";
export type BuiltinMemberKind = "event" | "member" | "method" | "property";
export type BuiltinSemanticModifier = "readonly";
export type BuiltinSemanticType = "enumMember" | "function" | "keyword" | "type" | "variable";

export interface BuiltinSignatureParameter {
  dataType?: string;
  description?: string;
  isRequired?: boolean;
  label: string;
  name: string;
}

export interface BuiltinCallableSignature {
  label: string;
  ownerName?: string;
  parameters: BuiltinSignatureParameter[];
  returnType?: string;
}

export interface BuiltinReferenceItem {
  completionKind: BuiltinCompletionKind;
  detail: string;
  documentation?: string;
  learnUrl?: string;
  modifiers: BuiltinSemanticModifier[];
  name: string;
  normalizedName: string;
  semanticType: BuiltinSemanticType;
  signature?: BuiltinCallableSignature;
  summary?: string;
  typeName?: string;
}

export interface BuiltinMemberReferenceItem extends BuiltinReferenceItem {
  memberKind: BuiltinMemberKind;
  ownerName: string;
  ownerNormalizedName: string;
}

interface DerivedReferenceData {
  builtinIdentifiers: Set<string>;
  completionItems: BuiltinReferenceItem[];
  byNormalizedName: Map<string, BuiltinReferenceItem>;
  memberItemsByOwnerAndName: Map<string, BuiltinMemberReferenceItem>;
  memberItemsByOwnerNormalizedName: Map<string, BuiltinMemberReferenceItem[]>;
  reservedIdentifiers: Set<string>;
}

type RawReferenceData = Record<string, unknown>;
type RawReferenceEntry = Record<string, unknown>;

const REFERENCE_FILE_NAME = "mslearn-vba-reference.json";
const INDEXED_COLLECTION_OWNER_TYPES = new Map<string, string>([
  ["dialogsheets", "DialogSheet"],
  ["worksheets", "Worksheet"],
]);
const BASE_BUILTIN_COMPLETIONS: Array<
  Omit<BuiltinReferenceItem, "detail" | "documentation" | "modifiers" | "normalizedName"> & {
    detail: string;
    documentation?: string;
    modifiers?: BuiltinSemanticModifier[];
    priority: number;
  }
> = [
  { completionKind: "variable", detail: "Excel built-in object", name: "ActiveCell", priority: 10, semanticType: "variable", typeName: "Range" },
  { completionKind: "variable", detail: "Excel built-in object", name: "ActiveChart", priority: 10, semanticType: "variable" },
  { completionKind: "variable", detail: "Excel built-in object", name: "ActivePrinter", priority: 10, semanticType: "variable" },
  { completionKind: "variable", detail: "Excel built-in object", name: "ActiveSheet", priority: 10, semanticType: "variable" },
  { completionKind: "variable", detail: "Excel built-in object", name: "ActiveWorkbook", priority: 10, semanticType: "variable", typeName: "Workbook" },
  { completionKind: "function", detail: "VBA function", name: "Abs", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Array", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Asc", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "CBool", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "CByte", priority: 10, semanticType: "function" },
  { completionKind: "variable", detail: "Excel built-in object", name: "Cells", priority: 10, semanticType: "variable", typeName: "Range" },
  { completionKind: "function", detail: "VBA function", name: "CDate", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "CDbl", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "CInt", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "CLng", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "CLngPtr", priority: 10, semanticType: "function" },
  { completionKind: "type", detail: "VBA object", name: "Collection", priority: 10, semanticType: "type" },
  { completionKind: "function", detail: "VBA function", name: "Command$", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "CreateObject", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "CStr", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "CVar", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Date", priority: 10, semanticType: "function" },
  { completionKind: "variable", detail: "VBA object", name: "Debug", priority: 10, semanticType: "variable" },
  { completionKind: "function", detail: "VBA function", name: "Dir", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "DoEvents", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Environ", priority: 10, semanticType: "function" },
  { completionKind: "variable", detail: "VBA object", name: "Err", priority: 10, semanticType: "variable" },
  { completionKind: "function", detail: "VBA function", name: "Format", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "InStr", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Int", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "IsArray", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "IsDate", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "IsEmpty", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "IsError", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "IsNull", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "IsNumeric", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "IsObject", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "LBound", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Left", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Len", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Mid", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "MsgBox", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Now", priority: 10, semanticType: "function" },
  { completionKind: "variable", detail: "Excel built-in object", name: "Range", priority: 10, semanticType: "variable" },
  { completionKind: "function", detail: "VBA function", name: "Replace", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Right", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Round", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Split", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Time", priority: 10, semanticType: "function" },
  { completionKind: "variable", detail: "Excel built-in object", name: "ThisWorkbook", priority: 10, semanticType: "variable", typeName: "Workbook" },
  { completionKind: "function", detail: "VBA function", name: "Trim", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "TypeName", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "UBound", priority: 10, semanticType: "function" },
  { completionKind: "function", detail: "VBA function", name: "Val", priority: 10, semanticType: "function" },
  { completionKind: "variable", detail: "VBA built-in namespace", name: "VBA", priority: 10, semanticType: "variable" },
  { completionKind: "variable", detail: "Excel built-in object", name: "WorksheetFunction", priority: 10, semanticType: "variable" },
  { completionKind: "variable", detail: "Excel built-in object", name: "Worksheets", priority: 10, semanticType: "variable" }
];

const derivedReferenceData = createDerivedReferenceData();

export const BUILTIN_IDENTIFIERS = new Set([
  ...BASE_BUILTIN_COMPLETIONS.map((entry) => normalizeIdentifier(entry.name)),
  ...derivedReferenceData.builtinIdentifiers
]);

export const RESERVED_IDENTIFIERS = new Set([
  ...VBA_KEYWORDS,
  ...derivedReferenceData.reservedIdentifiers,
  ...BUILTIN_IDENTIFIERS
]);

export const BUILTIN_REFERENCE_ITEMS = derivedReferenceData.completionItems;

export function getBuiltinCompletionItems(prefix?: string): BuiltinReferenceItem[] {
  if (!prefix) {
    return [];
  }

  const normalizedPrefix = normalizeIdentifier(prefix);

  return BUILTIN_REFERENCE_ITEMS.filter((item) => item.normalizedName.startsWith(normalizedPrefix));
}

export function getBuiltinReferenceItem(name: string): BuiltinReferenceItem | undefined {
  return derivedReferenceData.byNormalizedName.get(normalizeIdentifier(name));
}

export function getBuiltinMemberCompletionItems(ownerName: string, prefix?: string): BuiltinMemberReferenceItem[] {
  const items = derivedReferenceData.memberItemsByOwnerNormalizedName.get(normalizeIdentifier(ownerName)) ?? [];

  if (!prefix) {
    return [...items];
  }

  const normalizedPrefix = normalizeIdentifier(prefix);
  return items.filter((item) => item.normalizedName.startsWith(normalizedPrefix));
}

export function getBuiltinMemberReferenceItem(ownerName: string, memberName: string): BuiltinMemberReferenceItem | undefined {
  return derivedReferenceData.memberItemsByOwnerAndName.get(createOwnerMemberKey(ownerName, memberName));
}

export function getBuiltinMemberSignature(ownerName: string, memberName: string): BuiltinCallableSignature | undefined {
  return getBuiltinMemberReferenceItem(ownerName, memberName)?.signature;
}

export function resolveBuiltinMemberOwner(pathSegments: string[]): string | undefined {
  if (pathSegments.length === 0) {
    return undefined;
  }

  const [rootSegment, ...memberSegments] = pathSegments;
  const rootReference = getBuiltinReferenceItem(stripIndexedAccessMarker(rootSegment));

  if (!rootReference) {
    return undefined;
  }

  let currentOwnerName = rootReference.typeName ?? rootReference.name;

  if (hasIndexedAccessMarker(rootSegment)) {
    currentOwnerName = resolveIndexedCollectionOwnerTypeName(currentOwnerName) ?? currentOwnerName;
  }

  return resolveBuiltinMemberOwnerFromRootType(currentOwnerName, memberSegments);
}

export function resolveBuiltinMemberOwnerFromRootType(rootOwnerName: string, memberSegments: string[]): string | undefined {
  let currentOwnerName = rootOwnerName;

  for (const memberSegment of memberSegments) {
    const memberReference = getBuiltinMemberReferenceItem(currentOwnerName, stripIndexedAccessMarker(memberSegment));

    if (!memberReference?.typeName) {
      return undefined;
    }

    currentOwnerName = memberReference.typeName;

    if (hasIndexedAccessMarker(memberSegment)) {
      currentOwnerName = resolveIndexedCollectionOwnerTypeName(currentOwnerName) ?? currentOwnerName;
    }
  }

  return currentOwnerName;
}

export function isReservedOrBuiltinIdentifier(name: string): boolean {
  return RESERVED_IDENTIFIERS.has(normalizeIdentifier(name));
}

function createDerivedReferenceData(): DerivedReferenceData {
  const byNormalizedName = new Map<string, BuiltinReferenceItem>();
  const memberItemsByOwnerAndName = new Map<string, BuiltinMemberReferenceItem>();
  const memberItemsByOwnerNormalizedName = new Map<string, BuiltinMemberReferenceItem[]>();
  const memberPriorities = new Map<string, number>();
  const priorities = new Map<string, number>();
  const builtinIdentifiers = new Set<string>();
  const reservedIdentifiers = new Set<string>();
  const rawReferenceData = loadReferenceData();

  const addEntry = (entry: BuiltinReferenceItem, priority: number): void => {
    reservedIdentifiers.add(entry.normalizedName);

    if (entry.completionKind !== "keyword") {
      builtinIdentifiers.add(entry.normalizedName);
    }

    const currentPriority = priorities.get(entry.normalizedName) ?? -1;

    if (currentPriority > priority) {
      return;
    }

    priorities.set(entry.normalizedName, priority);
    byNormalizedName.set(entry.normalizedName, entry);
  };

  const addMemberEntry = (entry: BuiltinMemberReferenceItem, priority: number): void => {
    const entryKey = createOwnerMemberKey(entry.ownerName, entry.name);
    const currentPriority = memberPriorities.get(entryKey) ?? -1;

    if (currentPriority > priority) {
      return;
    }

    memberPriorities.set(entryKey, priority);
    memberItemsByOwnerAndName.set(entryKey, entry);
  };

  for (const keyword of VBA_KEYWORDS) {
    addEntry(createReferenceItem(keyword, "keyword", "VBA keyword"), 20);
  }

  for (const item of readEntryArray(rawReferenceData, "languageReference", "keywords")) {
    addEntry(
      createReferenceItem(readString(item, "name"), "keyword", "VBA keyword", readDocumentation("VBA keyword", resolveKeywordUrl(item))),
      30
    );
  }

  for (const item of readEntryArray(rawReferenceData, "languageReference", "statements")) {
    addEntry(
      createReferenceItem(readString(item, "name"), "keyword", "VBA statement", readDocumentation("VBA statement", readString(item, "learnUrl"))),
      40
    );
  }

  for (const [groupName, items] of readGroupedEntries(rawReferenceData, "languageReference", "functions")) {
    for (const item of items) {
      addEntry(
        createReferenceItem(
          readString(item, "name"),
          "function",
          `VBA function (${groupName})`,
          readDocumentation(`VBA function (${groupName})`, readString(item, "learnUrl"))
        ),
        50
      );
    }
  }

  for (const item of readEntryArray(rawReferenceData, "languageReference", "objects")) {
    addEntry(
      createReferenceItem(readString(item, "name"), "type", "VBA object", readDocumentation("VBA object", readString(item, "learnUrl"))),
      60
    );
  }

  for (const item of readEntryArray(rawReferenceData, "excel", "constantsEnumeration")) {
    addEntry(
      createReferenceItem(
        readString(item, "name"),
        "constant",
        "Excel constant",
        readDocumentation("Excel constant", readString(item, "learnUrl")),
        {
          modifiers: ["readonly"],
          typeName: "Long"
        }
      ),
      70
    );
  }

  for (const item of readEntryArray(rawReferenceData, "excel", "objectModel", "enumerations")) {
    addEntry(
      createReferenceItem(readString(item, "name"), "type", "Excel enumeration", readDocumentation("Excel enumeration", readString(item, "learnUrl"))),
      80
    );
  }

  for (const item of readEntryArray(rawReferenceData, "excel", "objectModel", "items")) {
    addEntry(
      createReferenceItem(readString(item, "name"), "type", "Excel object", readDocumentation("Excel object", readString(item, "learnUrl"))),
      90
    );
  }

  for (const item of readEntryArray(rawReferenceData, "libraryReference", "reference", "enumerations")) {
    addEntry(
      createReferenceItem(readString(item, "name"), "type", "Office enumeration", readDocumentation("Office enumeration", readString(item, "learnUrl"))),
      100
    );
  }

  for (const item of readEntryArray(rawReferenceData, "libraryReference", "reference", "items")) {
    addEntry(
      createReferenceItem(readString(item, "name"), "type", "Office object", readDocumentation("Office object", readString(item, "learnUrl"))),
      110
    );
  }

  for (const fallback of BASE_BUILTIN_COMPLETIONS) {
    addEntry(
      createReferenceItem(
        fallback.name,
        fallback.completionKind,
        fallback.detail,
        fallback.documentation,
        {
          modifiers: fallback.modifiers ?? [],
          typeName: fallback.typeName
        }
      ),
      fallback.priority
    );
  }

  addObjectMemberEntries(
    readEntryArray(rawReferenceData, "excel", "objectModel", "items"),
    "Excel",
    byNormalizedName,
    addMemberEntry,
    120
  );
  addObjectMemberEntries(
    readEntryArray(rawReferenceData, "libraryReference", "reference", "items"),
    "Office",
    byNormalizedName,
    addMemberEntry,
    130
  );

  for (const memberItem of memberItemsByOwnerAndName.values()) {
    const ownerItems = memberItemsByOwnerNormalizedName.get(memberItem.ownerNormalizedName);

    if (ownerItems) {
      ownerItems.push(memberItem);
    } else {
      memberItemsByOwnerNormalizedName.set(memberItem.ownerNormalizedName, [memberItem]);
    }
  }

  for (const ownerItems of memberItemsByOwnerNormalizedName.values()) {
    ownerItems.sort((left, right) => left.name.localeCompare(right.name));
  }

  return {
    builtinIdentifiers,
    byNormalizedName,
    completionItems: [...byNormalizedName.values()].sort((left, right) => left.name.localeCompare(right.name)),
    memberItemsByOwnerAndName,
    memberItemsByOwnerNormalizedName,
    reservedIdentifiers
  };
}

function createReferenceItem(
  name: string | undefined,
  completionKind: BuiltinCompletionKind,
  detail: string,
  documentation?: string,
  options: {
    learnUrl?: string;
    modifiers?: BuiltinSemanticModifier[];
    signature?: BuiltinCallableSignature;
    summary?: string;
    typeName?: string;
  } = {}
): BuiltinReferenceItem {
  const safeName = name ?? "";

  return {
    completionKind,
    detail,
    documentation,
    learnUrl: options.learnUrl,
    modifiers: options.modifiers ?? [],
    name: safeName,
    normalizedName: normalizeIdentifier(safeName),
    semanticType: mapSemanticType(completionKind),
    signature: options.signature,
    summary: options.summary,
    typeName: options.typeName
  };
}

function createMemberReferenceItem(
  ownerName: string,
  memberSource: RawReferenceEntry,
  sectionTitle: string | undefined,
  sourceLabel: string,
  knownItemsByNormalizedName: ReadonlyMap<string, BuiltinReferenceItem>
): BuiltinMemberReferenceItem {
  const name = readString(memberSource, "name");
  const safeName = name ?? "";
  const normalizedSectionTitle = normalizeIdentifier(sectionTitle ?? "");
  const completionKind =
    normalizedSectionTitle === "methods" || normalizedSectionTitle === "events" ? "function" : "variable";
  const memberKind =
    normalizedSectionTitle === "methods"
      ? "method"
      : normalizedSectionTitle === "properties"
        ? "property"
        : normalizedSectionTitle === "events"
          ? "event"
          : "member";
  const inferredTypeName = readString(memberSource, "typeName") ?? inferBuiltinMemberTypeName(safeName, knownItemsByNormalizedName);
  const learnUrl = readString(memberSource, "learnUrl");
  const summary = readString(memberSource, "summary");
  const signature = readBuiltinSignature(memberSource);

  return {
    ...createReferenceItem(
      safeName,
      completionKind,
      `${sourceLabel} ${ownerName} ${memberKind}`,
      readDocumentation(`${sourceLabel} ${ownerName} ${memberKind}`, learnUrl, summary),
      {
        learnUrl,
        signature,
        summary,
        typeName: inferredTypeName
      }
    ),
    memberKind,
    ownerName,
    ownerNormalizedName: normalizeIdentifier(ownerName)
  };
}

function mapSemanticType(completionKind: BuiltinCompletionKind): BuiltinSemanticType {
  switch (completionKind) {
    case "constant":
      return "variable";
    case "function":
      return "function";
    case "keyword":
      return "keyword";
    case "type":
      return "type";
    case "variable":
    default:
      return "variable";
  }
}

function loadReferenceData(): RawReferenceData {
  const referenceFilePath = resolveReferenceFilePath();

  if (!referenceFilePath) {
    return {};
  }

  try {
    return JSON.parse(readFileSync(referenceFilePath, "utf8")) as RawReferenceData;
  } catch {
    return {};
  }
}

function addObjectMemberEntries(
  items: RawReferenceEntry[],
  sourceLabel: string,
  knownItemsByNormalizedName: ReadonlyMap<string, BuiltinReferenceItem>,
  addMemberEntry: (entry: BuiltinMemberReferenceItem, priority: number) => void,
  priority: number
): void {
  for (const item of items) {
    const ownerName = readString(item, "name");

    if (!ownerName) {
      continue;
    }

    for (const section of readNestedEntryArray(item, "sections")) {
      const sectionTitle = readString(section, "title");

      for (const member of readNestedEntryArray(section, "members")) {
        addMemberEntry(
          createMemberReferenceItem(
            ownerName,
            member,
            sectionTitle,
            sourceLabel,
            knownItemsByNormalizedName
          ),
          priority
        );
      }
    }
  }
}

function resolveReferenceFilePath(): string | undefined {
  const candidatePaths = [
    process.env.VBA_REFERENCE_DATA_PATH,
    path.resolve(process.cwd(), "resources", "reference", REFERENCE_FILE_NAME),
    path.resolve(__dirname, "..", "..", "resources", "reference", REFERENCE_FILE_NAME),
    path.resolve(__dirname, "..", "..", "..", "resources", "reference", REFERENCE_FILE_NAME),
    path.resolve(__dirname, "..", "..", "..", "..", "resources", "reference", REFERENCE_FILE_NAME)
  ].filter((candidatePath): candidatePath is string => Boolean(candidatePath));

  return candidatePaths.find((candidatePath) => existsSync(candidatePath));
}

function readEntryArray(source: RawReferenceData, ...pathSegments: string[]): RawReferenceEntry[] {
  const value = readNestedValue(source, ...pathSegments);
  return Array.isArray(value) ? value.filter(isRecord) : [];
}

function readNestedEntryArray(source: RawReferenceEntry, key: string): RawReferenceEntry[] {
  const value = source[key];
  return Array.isArray(value) ? value.filter(isRecord) : [];
}

function readGroupedEntries(source: RawReferenceData, ...pathSegments: string[]): Array<[string, RawReferenceEntry[]]> {
  const value = readNestedValue(source, ...pathSegments);

  if (!isRecord(value)) {
    return [];
  }

  return Object.entries(value)
    .filter(([, items]) => Array.isArray(items))
    .map(([groupName, items]) => [groupName, (items as unknown[]).filter(isRecord)]);
}

function readNestedValue(source: RawReferenceData, ...pathSegments: string[]): unknown {
  return pathSegments.reduce<unknown>((currentValue, pathSegment) => {
    if (!isRecord(currentValue)) {
      return undefined;
    }

    return currentValue[pathSegment];
  }, source);
}

function readString(source: RawReferenceEntry, key: string): string | undefined {
  const value = source[key];
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

function resolveKeywordUrl(source: RawReferenceEntry): string | undefined {
  const contexts = source.contexts;

  if (Array.isArray(contexts)) {
    const firstContext = contexts.find(isRecord);
    return firstContext ? readString(firstContext, "learnUrl") : undefined;
  }

  return readString(source, "learnUrl");
}

function readDocumentation(detail: string, learnUrl?: string, summary?: string): string | undefined {
  const lines = [detail];

  if (summary) {
    lines.push(summary);
  }

  if (learnUrl) {
    lines.push(learnUrl);
  }

  return lines.length > 1 ? lines.join("\n") : learnUrl ? `${detail}\n${learnUrl}` : undefined;
}

function inferBuiltinMemberTypeName(
  memberName: string,
  knownItemsByNormalizedName: ReadonlyMap<string, BuiltinReferenceItem>
): string | undefined {
  const knownItem = knownItemsByNormalizedName.get(normalizeIdentifier(memberName));

  if (!knownItem) {
    return undefined;
  }

  return knownItem.typeName ?? (knownItem.completionKind === "type" ? knownItem.name : undefined);
}

function readBuiltinSignature(source: RawReferenceEntry): BuiltinCallableSignature | undefined {
  const value = source.signature;

  if (!isRecord(value)) {
    return undefined;
  }

  const parameters = Array.isArray(value.parameters)
    ? value.parameters.filter(isRecord).map((parameter) => ({
        dataType: readString(parameter, "dataType"),
        description: readString(parameter, "description"),
        isRequired: typeof parameter.isRequired === "boolean" ? parameter.isRequired : undefined,
        label: readString(parameter, "label") ?? "",
        name: readString(parameter, "name") ?? ""
      }))
    : [];

  const label = readString(value, "label");

  if (!label) {
    return undefined;
  }

  return {
    label,
    ownerName: readString(value, "ownerName"),
    parameters,
    returnType: readString(value, "returnType")
  };
}

function createOwnerMemberKey(ownerName: string, memberName: string): string {
  return `${normalizeIdentifier(ownerName)}:${normalizeIdentifier(memberName)}`;
}

function hasIndexedAccessMarker(pathSegment: string): boolean {
  return pathSegment.endsWith("()");
}

export function stripIndexedAccessMarker(pathSegment: string): string {
  return hasIndexedAccessMarker(pathSegment) ? pathSegment.slice(0, -2) : pathSegment;
}

function resolveIndexedCollectionOwnerTypeName(ownerName: string): string | undefined {
  return INDEXED_COLLECTION_OWNER_TYPES.get(normalizeIdentifier(ownerName));
}

function isRecord(value: unknown): value is RawReferenceEntry {
  return typeof value === "object" && value !== null;
}
