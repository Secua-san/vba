import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

export const WORKBOOK_BINDING_MANIFEST_ARTIFACT = "workbook-binding-manifest";
export const WORKBOOK_BINDING_MANIFEST_BINDING_KIND = "active-workbook-fullname";
export const WORKBOOK_BINDING_MANIFEST_DIRECTORY_NAME = ".vba";
export const WORKBOOK_BINDING_MANIFEST_FILE_NAME = "workbook-binding.json";
export const WORKBOOK_BINDING_MANIFEST_VERSION = 1;

export interface WorkbookBindingManifestWorkbook {
  fullName: string;
  isAddIn: boolean;
  name: string;
  path: string;
  sourceKind: "openxml-package";
}

export interface WorkbookBindingManifest {
  artifact: typeof WORKBOOK_BINDING_MANIFEST_ARTIFACT;
  bindingKind: typeof WORKBOOK_BINDING_MANIFEST_BINDING_KIND;
  version: typeof WORKBOOK_BINDING_MANIFEST_VERSION;
  workbook: WorkbookBindingManifestWorkbook;
}

export interface WorkbookBindingManifestLookupOptions {
  workspaceRoots?: readonly string[];
}

export interface WorkbookBindingManifestLocation {
  bundleRoot: string;
  manifestPath: string;
}

export interface WorkbookBindingManifestValidationIssue {
  code:
    | "invalid-artifact"
    | "invalid-binding-kind"
    | "invalid-json"
    | "invalid-top-level"
    | "invalid-version"
    | "invalid-workbook"
    | "missing-required-field";
  message: string;
  path: string;
}

export interface WorkbookBindingManifestParseResult {
  issues: WorkbookBindingManifestValidationIssue[];
  manifest?: WorkbookBindingManifest;
}

export function buildWorkbookBindingManifestPath(bundleRoot: string): string {
  return path.join(bundleRoot, WORKBOOK_BINDING_MANIFEST_DIRECTORY_NAME, WORKBOOK_BINDING_MANIFEST_FILE_NAME);
}

export function findNearestWorkbookBindingManifest(
  filePath: string,
  options?: WorkbookBindingManifestLookupOptions
): WorkbookBindingManifestLocation | undefined {
  const resolvedFilePath = path.resolve(filePath);
  const boundaryPath = findLookupBoundaryPath(resolvedFilePath, options?.workspaceRoots);

  if (!boundaryPath) {
    return undefined;
  }

  let currentDirectory = path.dirname(resolvedFilePath);

  while (true) {
    const candidatePath = buildWorkbookBindingManifestPath(currentDirectory);

    if (existsSync(candidatePath)) {
      return {
        bundleRoot: currentDirectory,
        manifestPath: candidatePath
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

export function parseWorkbookBindingManifest(rawText: string): WorkbookBindingManifestParseResult {
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

  const issues: WorkbookBindingManifestValidationIssue[] = [];

  if (parsedValue.version !== WORKBOOK_BINDING_MANIFEST_VERSION) {
    issues.push({
      code: "invalid-version",
      message: `version は ${WORKBOOK_BINDING_MANIFEST_VERSION} である必要があります`,
      path: "$.version"
    });
  }

  if (parsedValue.artifact !== WORKBOOK_BINDING_MANIFEST_ARTIFACT) {
    issues.push({
      code: "invalid-artifact",
      message: `artifact は ${WORKBOOK_BINDING_MANIFEST_ARTIFACT} である必要があります`,
      path: "$.artifact"
    });
  }

  if (parsedValue.bindingKind !== WORKBOOK_BINDING_MANIFEST_BINDING_KIND) {
    issues.push({
      code: "invalid-binding-kind",
      message: `bindingKind は ${WORKBOOK_BINDING_MANIFEST_BINDING_KIND} である必要があります`,
      path: "$.bindingKind"
    });
  }

  const workbook = parseWorkbook(parsedValue.workbook, issues);

  if (issues.some((issue) => issue.code === "invalid-version" || issue.code === "invalid-artifact" || issue.code === "invalid-binding-kind")) {
    return { issues };
  }

  if (!workbook) {
    return { issues };
  }

  return {
    issues,
    manifest: {
      artifact: WORKBOOK_BINDING_MANIFEST_ARTIFACT,
      bindingKind: WORKBOOK_BINDING_MANIFEST_BINDING_KIND,
      version: WORKBOOK_BINDING_MANIFEST_VERSION,
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
  if (process.platform === "win32") {
    return path.resolve(left).toLowerCase() === path.resolve(right).toLowerCase();
  }

  return path.resolve(left) === path.resolve(right);
}

function parseWorkbook(
  value: unknown,
  issues: WorkbookBindingManifestValidationIssue[]
): WorkbookBindingManifestWorkbook | undefined {
  if (!isRecord(value)) {
    issues.push({
      code: "invalid-top-level",
      message: "workbook は object である必要があります",
      path: "$.workbook"
    });
    return undefined;
  }

  const fullName = getNonEmptyString(value.fullName);
  const name = getNonEmptyString(value.name);
  const workbookPath = getNonEmptyString(value.path);
  const isAddIn = value.isAddIn;
  const sourceKind = value.sourceKind;

  if (!fullName) {
    issues.push({
      code: "missing-required-field",
      message: "workbook.fullName は空でない文字列である必要があります",
      path: "$.workbook.fullName"
    });
  }

  if (!name) {
    issues.push({
      code: "missing-required-field",
      message: "workbook.name は空でない文字列である必要があります",
      path: "$.workbook.name"
    });
  }

  if (!workbookPath) {
    issues.push({
      code: "missing-required-field",
      message: "workbook.path は空でない文字列である必要があります",
      path: "$.workbook.path"
    });
  }

  if (typeof isAddIn !== "boolean") {
    issues.push({
      code: "missing-required-field",
      message: "workbook.isAddIn は boolean である必要があります",
      path: "$.workbook.isAddIn"
    });
  }

  if (sourceKind !== "openxml-package") {
    issues.push({
      code: "missing-required-field",
      message: "workbook.sourceKind は openxml-package である必要があります",
      path: "$.workbook.sourceKind"
    });
  }

  if (typeof isAddIn === "boolean" && isAddIn) {
    issues.push({
      code: "invalid-workbook",
      message: "workbook.isAddIn は false である必要があります",
      path: "$.workbook.isAddIn"
    });
  }

  if (!fullName || !name || !workbookPath || typeof isAddIn !== "boolean" || sourceKind !== "openxml-package" || isAddIn) {
    return undefined;
  }

  return {
    fullName,
    isAddIn,
    name,
    path: workbookPath,
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
