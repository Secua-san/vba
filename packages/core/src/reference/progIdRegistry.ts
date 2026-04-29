import { normalizeIdentifier } from "../types/helpers";

const CREATE_OBJECT_PROGID_TYPES = new Map<string, string>([
  ["scripting.dictionary", "ScriptingDictionary"],
  ["wscript.shell", "WshShell"]
]);

const KNOWN_PROGID_OWNER_TYPE_NAMES = new Set(
  [...CREATE_OBJECT_PROGID_TYPES.values()].map((typeName) => normalizeIdentifier(typeName))
);

export function resolveCreateObjectProgIdType(progId: string): string | undefined {
  return CREATE_OBJECT_PROGID_TYPES.get(progId.toLowerCase());
}

export function isKnownProgIdOwnerTypeName(typeName: string | undefined): boolean {
  return KNOWN_PROGID_OWNER_TYPE_NAMES.has(normalizeIdentifier(typeName ?? ""));
}
