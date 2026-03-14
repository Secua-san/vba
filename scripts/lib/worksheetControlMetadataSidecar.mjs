import path from "node:path";

export const worksheetControlMetadataSidecarArtifact = "worksheet-control-metadata-sidecar";
export const worksheetControlMetadataSidecarFileName = "worksheet-control-metadata.json";
export const worksheetControlMetadataSidecarDirectoryName = ".vba";
export const worksheetControlMetadataSidecarVersion = 1;

const controlTypeByProgId = new Map([
  ["forms.checkbox.1", "CheckBox"],
  ["forms.combobox.1", "ComboBox"],
  ["forms.commandbutton.1", "CommandButton"],
  ["forms.image.1", "Image"],
  ["forms.label.1", "Label"],
  ["forms.listbox.1", "ListBox"],
  ["forms.optionbutton.1", "OptionButton"],
  ["forms.scrollbar.1", "ScrollBar"],
  ["forms.spinbutton.1", "SpinButton"],
  ["forms.textbox.1", "TextBox"],
  ["forms.togglebutton.1", "ToggleButton"],
]);

const controlTypeByClassId = new Map([
  ["{8bd21d40-ec42-11ce-9e0d-00aa006002f3}", "CheckBox"],
]);

export function buildWorksheetControlMetadataSidecarPath(bundleRoot) {
  return path.join(
    bundleRoot,
    worksheetControlMetadataSidecarDirectoryName,
    worksheetControlMetadataSidecarFileName,
  );
}

export function convertWorksheetControlMetadataProbeToSidecar(probeMetadata) {
  const workbookName = requireNonEmptyString(probeMetadata?.workbook, "probe.workbook");
  const worksheets = requireArray(probeMetadata?.worksheets, "probe.worksheets");

  return {
    artifact: worksheetControlMetadataSidecarArtifact,
    owners: worksheets.map((worksheet, index) => convertWorksheetOwner(worksheet, index)),
    version: worksheetControlMetadataSidecarVersion,
    workbook: {
      name: workbookName,
      sourceKind: "openxml-package",
    },
  };
}

function convertWorksheetOwner(worksheet, index) {
  const ownerPath = `probe.worksheets[${index}]`;
  const controls = requireArray(worksheet?.controls, `${ownerPath}.controls`);

  return {
    controls: controls.map((control, controlIndex) => convertWorksheetControl(control, ownerPath, controlIndex)),
    ownerKind: "worksheet",
    sheetCodeName: requireNonEmptyString(worksheet?.sheetCodeName, `${ownerPath}.sheetCodeName`),
    sheetName: requireNonEmptyString(worksheet?.sheetName, `${ownerPath}.sheetName`),
    status: "supported",
  };
}

function convertWorksheetControl(control, ownerPath, controlIndex) {
  const controlPath = `${ownerPath}.controls[${controlIndex}]`;
  const progId = normalizeOptionalString(control?.progId, `${controlPath}.progId`);
  const classId = normalizeOptionalString(control?.classId, `${controlPath}.classId`);
  const controlType = resolveControlType(progId, classId);

  if (!controlType) {
    throw new Error(`${controlPath} の controlType を解決できません`);
  }

  return {
    classId,
    codeName: requireNonEmptyString(control?.codeName, `${controlPath}.codeName`),
    controlType,
    progId,
    shapeId: requireUnsignedInteger(control?.shapeId, `${controlPath}.shapeId`),
    shapeName: requireNonEmptyString(control?.shapeName, `${controlPath}.shapeName`),
  };
}

function normalizeOptionalString(value, pathLabel) {
  if (value === undefined || value === null) {
    return null;
  }

  const normalizedValue = String(value).trim();

  if (normalizedValue.length === 0) {
    throw new Error(`${pathLabel} は空文字列を許可しません`);
  }

  return normalizedValue;
}

function requireArray(value, pathLabel) {
  if (!Array.isArray(value)) {
    throw new Error(`${pathLabel} は配列である必要があります`);
  }

  return value;
}

function requireNonEmptyString(value, pathLabel) {
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`${pathLabel} は空でない文字列である必要があります`);
  }

  return value.trim();
}

function requireUnsignedInteger(value, pathLabel) {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new Error(`${pathLabel} は 0 以上の整数である必要があります`);
  }

  return value;
}

function resolveControlType(progId, classId) {
  const normalizedProgId = progId ? progId.toLowerCase() : null;

  if (normalizedProgId && controlTypeByProgId.has(normalizedProgId)) {
    return controlTypeByProgId.get(normalizedProgId);
  }

  const normalizedClassId = classId ? classId.toLowerCase() : null;

  if (normalizedClassId && controlTypeByClassId.has(normalizedClassId)) {
    return controlTypeByClassId.get(normalizedClassId);
  }

  return null;
}
