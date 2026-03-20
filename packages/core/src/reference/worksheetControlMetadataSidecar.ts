import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

export const WORKSHEET_CONTROL_METADATA_SIDECAR_ARTIFACT = "worksheet-control-metadata-sidecar";
export const WORKSHEET_CONTROL_METADATA_SIDECAR_DIRECTORY_NAME = ".vba";
export const WORKSHEET_CONTROL_METADATA_SIDECAR_FILE_NAME = "worksheet-control-metadata.json";
export const WORKSHEET_CONTROL_METADATA_SIDECAR_VERSION = 1;

export interface WorksheetControlMetadataSidecarControl {
  classId?: string;
  codeName: string;
  controlType: string;
  progId?: string;
  shapeId: number;
  shapeName: string;
}

export interface WorksheetControlMetadataSidecarWorkbook {
  name: string;
  sourceKind: "openxml-package";
}

export interface WorksheetControlMetadataSupportedOwner {
  controls: WorksheetControlMetadataSidecarControl[];
  ownerKind: string;
  sheetCodeName: string;
  sheetName: string;
  status: "supported";
}

export interface WorksheetControlMetadataUnsupportedOwner {
  ownerKind: string;
  reason: string;
  sheetCodeName: string;
  sheetName: string;
  status: "unsupported";
}

export type WorksheetControlMetadataSidecarOwner =
  | WorksheetControlMetadataSupportedOwner
  | WorksheetControlMetadataUnsupportedOwner;

export interface WorksheetControlMetadataSidecar {
  artifact: typeof WORKSHEET_CONTROL_METADATA_SIDECAR_ARTIFACT;
  owners: WorksheetControlMetadataSidecarOwner[];
  version: typeof WORKSHEET_CONTROL_METADATA_SIDECAR_VERSION;
  workbook: WorksheetControlMetadataSidecarWorkbook;
}

export interface WorksheetControlMetadataSidecarLookupOptions {
  workspaceRoots?: readonly string[];
}

export interface WorksheetControlMetadataSidecarLocation {
  bundleRoot: string;
  sidecarPath: string;
}

export interface WorksheetControlMetadataValidationIssue {
  code:
    | "invalid-artifact"
    | "invalid-control"
    | "invalid-json"
    | "invalid-owner"
    | "invalid-top-level"
    | "invalid-version"
    | "missing-required-field";
  message: string;
  path: string;
}

export interface WorksheetControlMetadataSidecarParseResult {
  issues: WorksheetControlMetadataValidationIssue[];
  sidecar?: WorksheetControlMetadataSidecar;
}

export function buildWorksheetControlMetadataSidecarPath(bundleRoot: string): string {
  return path.join(bundleRoot, WORKSHEET_CONTROL_METADATA_SIDECAR_DIRECTORY_NAME, WORKSHEET_CONTROL_METADATA_SIDECAR_FILE_NAME);
}

export function findNearestWorksheetControlMetadataSidecar(
  filePath: string,
  options?: WorksheetControlMetadataSidecarLookupOptions
): WorksheetControlMetadataSidecarLocation | undefined {
  const resolvedFilePath = path.resolve(filePath);
  const boundaryPath = findLookupBoundaryPath(resolvedFilePath, options?.workspaceRoots);

  if (!boundaryPath) {
    return undefined;
  }

  let currentDirectory = path.dirname(resolvedFilePath);

  while (true) {
    const candidatePath = buildWorksheetControlMetadataSidecarPath(currentDirectory);

    if (existsSync(candidatePath)) {
      return {
        bundleRoot: currentDirectory,
        sidecarPath: candidatePath
      };
    }

    if (isSamePath(currentDirectory, boundaryPath)) {
      return undefined;
    }

    const parentDirectory = path.dirname(currentDirectory);

    if (isSamePath(parentDirectory, currentDirectory)) {
      return undefined;
    }

    currentDirectory = parentDirectory;
  }
}

export function getSupportedWorksheetControlMetadataOwners(
  sidecar: WorksheetControlMetadataSidecar
): WorksheetControlMetadataSupportedOwner[] {
  return sidecar.owners.filter((owner): owner is WorksheetControlMetadataSupportedOwner => owner.status === "supported");
}

export function parseWorksheetControlMetadataSidecar(rawText: string): WorksheetControlMetadataSidecarParseResult {
  let parsedValue: unknown;

  try {
    parsedValue = JSON.parse(rawText);
  } catch (error) {
    return {
      issues: [
        {
          code: "invalid-json",
          message: `JSON として解釈できません: ${String(error)}`,
          path: "$"
        }
      ]
    };
  }

  if (!isRecord(parsedValue)) {
    return {
      issues: [
        {
          code: "invalid-top-level",
          message: "top-level は object である必要があります",
          path: "$"
        }
      ]
    };
  }

  const issues: WorksheetControlMetadataValidationIssue[] = [];
  const version = parsedValue.version;

  if (version !== WORKSHEET_CONTROL_METADATA_SIDECAR_VERSION) {
    issues.push({
      code: "invalid-version",
      message: `version は ${WORKSHEET_CONTROL_METADATA_SIDECAR_VERSION} である必要があります`,
      path: "$.version"
    });
  }

  if (parsedValue.artifact !== WORKSHEET_CONTROL_METADATA_SIDECAR_ARTIFACT) {
    issues.push({
      code: "invalid-artifact",
      message: `artifact は ${WORKSHEET_CONTROL_METADATA_SIDECAR_ARTIFACT} である必要があります`,
      path: "$.artifact"
    });
  }

  const workbook = parseWorkbook(parsedValue.workbook, issues);
  const owners = parseOwners(parsedValue.owners, issues);

  if (issues.some((issue) => issue.code === "invalid-top-level" || issue.code === "invalid-version" || issue.code === "invalid-artifact")) {
    return { issues };
  }

  if (!workbook) {
    return { issues };
  }

  return {
    issues,
    sidecar: {
      artifact: WORKSHEET_CONTROL_METADATA_SIDECAR_ARTIFACT,
      owners,
      version: WORKSHEET_CONTROL_METADATA_SIDECAR_VERSION,
      workbook
    }
  };
}

function findLookupBoundaryPath(filePath: string, workspaceRoots?: readonly string[]): string | undefined {
  const resolvedWorkspaceRoots = (workspaceRoots ?? [])
    .map(resolveWorkspaceRootPath)
    .filter((value): value is string => Boolean(value))
    .filter((candidatePath) => isPathInsideBoundary(filePath, candidatePath))
    .sort((left, right) => right.length - left.length);

  return resolvedWorkspaceRoots[0];
}

function getNonEmptyString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
}

function isPathInsideBoundary(targetPath: string, boundaryPath: string): boolean {
  const relativePath = path.relative(boundaryPath, targetPath);
  return relativePath === "" || (!relativePath.startsWith("..") && !path.isAbsolute(relativePath));
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isSamePath(left: string, right: string): boolean {
  if (looksLikeWindowsPath(left) || looksLikeWindowsPath(right)) {
    return path.win32.resolve(left).toLowerCase() === path.win32.resolve(right).toLowerCase();
  }

  return path.resolve(left) === path.resolve(right);
}

function looksLikeWindowsPath(value: string): boolean {
  return /^[A-Za-z]:[\\/]/.test(value) || /^\\\\[^\\]+\\[^\\]+/.test(value);
}

function parseControl(
  value: unknown,
  ownerIndex: number,
  controlIndex: number,
  issues: WorksheetControlMetadataValidationIssue[]
): WorksheetControlMetadataSidecarControl | undefined {
  const controlPath = `$.owners[${ownerIndex}].controls[${controlIndex}]`;

  if (!isRecord(value)) {
    issues.push({
      code: "invalid-control",
      message: "control は object である必要があります",
      path: controlPath
    });
    return undefined;
  }

  const codeName = getNonEmptyString(value.codeName);
  const controlType = getNonEmptyString(value.controlType);
  const shapeName = getNonEmptyString(value.shapeName);
  const shapeId = value.shapeId;
  const isValidShapeId = typeof shapeId === "number" && Number.isSafeInteger(shapeId) && shapeId >= 0;

  if (!codeName) {
    issues.push({
      code: "missing-required-field",
      message: "control.codeName は空でない文字列である必要があります",
      path: `${controlPath}.codeName`
    });
  }

  if (!controlType) {
    issues.push({
      code: "missing-required-field",
      message: "control.controlType は空でない文字列である必要があります",
      path: `${controlPath}.controlType`
    });
  }

  if (!shapeName) {
    issues.push({
      code: "missing-required-field",
      message: "control.shapeName は空でない文字列である必要があります",
      path: `${controlPath}.shapeName`
    });
  }

  if (!isValidShapeId) {
    issues.push({
      code: "missing-required-field",
      message: "control.shapeId は 0 以上の整数である必要があります",
      path: `${controlPath}.shapeId`
    });
  }

  if (!codeName || !controlType || !shapeName || !isValidShapeId) {
    return undefined;
  }

  const normalizedShapeId = shapeId as number;

  const classId = getNonEmptyString(value.classId);
  const progId = getNonEmptyString(value.progId);

  return {
    ...(classId ? { classId } : {}),
    codeName,
    controlType,
    ...(progId ? { progId } : {}),
    shapeId: normalizedShapeId,
    shapeName
  };
}

function parseOwner(
  value: unknown,
  index: number,
  issues: WorksheetControlMetadataValidationIssue[]
): WorksheetControlMetadataSidecarOwner | undefined {
  const ownerPath = `$.owners[${index}]`;

  if (!isRecord(value)) {
    issues.push({
      code: "invalid-owner",
      message: "owner は object である必要があります",
      path: ownerPath
    });
    return undefined;
  }

  const ownerKind = getNonEmptyString(value.ownerKind);
  const sheetCodeName = getNonEmptyString(value.sheetCodeName);
  const sheetName = getNonEmptyString(value.sheetName);
  const status = value.status;

  if (!ownerKind) {
    issues.push({
      code: "missing-required-field",
      message: "owner.ownerKind は空でない文字列である必要があります",
      path: `${ownerPath}.ownerKind`
    });
  }

  if (!sheetCodeName) {
    issues.push({
      code: "missing-required-field",
      message: "owner.sheetCodeName は空でない文字列である必要があります",
      path: `${ownerPath}.sheetCodeName`
    });
  }

  if (!sheetName) {
    issues.push({
      code: "missing-required-field",
      message: "owner.sheetName は空でない文字列である必要があります",
      path: `${ownerPath}.sheetName`
    });
  }

  if (status !== "supported" && status !== "unsupported") {
    issues.push({
      code: "missing-required-field",
      message: "owner.status は supported または unsupported である必要があります",
      path: `${ownerPath}.status`
    });
    return undefined;
  }

  if (!ownerKind || !sheetCodeName || !sheetName) {
    return undefined;
  }

  if (status === "unsupported") {
    const reason = getNonEmptyString(value.reason);

    if (!reason) {
      issues.push({
        code: "missing-required-field",
        message: "unsupported owner には reason が必要です",
        path: `${ownerPath}.reason`
      });
      return undefined;
    }

    return {
      ownerKind,
      reason,
      sheetCodeName,
      sheetName,
      status
    };
  }

  if (!Array.isArray(value.controls)) {
    issues.push({
      code: "missing-required-field",
      message: "supported owner には controls 配列が必要です",
      path: `${ownerPath}.controls`
    });
    return undefined;
  }

  const controls = value.controls
    .map((control, controlIndex) => parseControl(control, index, controlIndex, issues))
    .filter((control): control is WorksheetControlMetadataSidecarControl => Boolean(control));

  return {
    controls,
    ownerKind,
    sheetCodeName,
    sheetName,
    status
  };
}

function parseOwners(value: unknown, issues: WorksheetControlMetadataValidationIssue[]): WorksheetControlMetadataSidecarOwner[] {
  if (!Array.isArray(value)) {
    issues.push({
      code: "invalid-top-level",
      message: "owners は配列である必要があります",
      path: "$.owners"
    });
    return [];
  }

  return value
    .map((owner, index) => parseOwner(owner, index, issues))
    .filter((owner): owner is WorksheetControlMetadataSidecarOwner => Boolean(owner));
}

function parseWorkbook(
  value: unknown,
  issues: WorksheetControlMetadataValidationIssue[]
): WorksheetControlMetadataSidecarWorkbook | undefined {
  if (!isRecord(value)) {
    issues.push({
      code: "invalid-top-level",
      message: "workbook は object である必要があります",
      path: "$.workbook"
    });
    return undefined;
  }

  const name = getNonEmptyString(value.name);
  const sourceKind = value.sourceKind;

  if (!name) {
    issues.push({
      code: "missing-required-field",
      message: "workbook.name は空でない文字列である必要があります",
      path: "$.workbook.name"
    });
  }

  if (sourceKind !== "openxml-package") {
    issues.push({
      code: "missing-required-field",
      message: "workbook.sourceKind は openxml-package である必要があります",
      path: "$.workbook.sourceKind"
    });
  }

  if (!name || sourceKind !== "openxml-package") {
    return undefined;
  }

  return {
    name,
    sourceKind
  };
}

function resolveWorkspaceRootPath(value: string): string | undefined {
  if (!value) {
    return undefined;
  }

  if (value.startsWith("file:")) {
    try {
      return fileURLToPath(value);
    } catch {
      return undefined;
    }
  }

  return path.resolve(value);
}
