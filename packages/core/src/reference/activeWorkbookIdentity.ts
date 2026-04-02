export const ACTIVE_WORKBOOK_IDENTITY_NOTIFICATION_METHOD = "vba/activeWorkbookIdentity";
export const ACTIVE_WORKBOOK_IDENTITY_TEST_SET_REQUEST_METHOD = "vba/test/setActiveWorkbookIdentitySnapshot";
export const ACTIVE_WORKBOOK_IDENTITY_TEST_STATE_REQUEST_METHOD = "vba/test/getActiveWorkbookIdentitySnapshot";
export const ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND = "excel-active-workbook";
export const ACTIVE_WORKBOOK_IDENTITY_VERSION = 1;

export type ActiveWorkbookIdentityUnavailableReason =
  | "host-error"
  | "host-unreachable"
  | "no-active-workbook"
  | "non-workbook-window";
export type ActiveWorkbookIdentityUnsupportedReason = "addin" | "unsaved";

export interface ActiveWorkbookIdentityFields {
  fullName: string;
  isAddin: boolean;
  name: string;
  path: string;
}

export interface ActiveWorkbookIdentityAvailableSnapshot {
  identity: ActiveWorkbookIdentityFields;
  observedAt: string;
  providerKind: typeof ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND;
  state: "available";
  version: typeof ACTIVE_WORKBOOK_IDENTITY_VERSION;
}

export interface ActiveWorkbookIdentityUnavailableSnapshot {
  observedAt: string;
  providerKind: typeof ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND;
  reason: ActiveWorkbookIdentityUnavailableReason;
  state: "unavailable";
  version: typeof ACTIVE_WORKBOOK_IDENTITY_VERSION;
}

export interface ActiveWorkbookProtectedViewFields {
  sourceName?: string;
  sourcePath?: string;
}

export interface ActiveWorkbookIdentityProtectedViewSnapshot {
  observedAt: string;
  protectedView?: ActiveWorkbookProtectedViewFields;
  providerKind: typeof ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND;
  state: "protected-view";
  version: typeof ACTIVE_WORKBOOK_IDENTITY_VERSION;
}

export interface ActiveWorkbookIdentityUnsupportedSnapshot {
  identity: ActiveWorkbookIdentityFields;
  observedAt: string;
  providerKind: typeof ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND;
  reason: ActiveWorkbookIdentityUnsupportedReason;
  state: "unsupported";
  version: typeof ACTIVE_WORKBOOK_IDENTITY_VERSION;
}

export type ActiveWorkbookIdentitySnapshot =
  | ActiveWorkbookIdentityAvailableSnapshot
  | ActiveWorkbookIdentityUnavailableSnapshot
  | ActiveWorkbookIdentityProtectedViewSnapshot
  | ActiveWorkbookIdentityUnsupportedSnapshot;

export interface ActiveWorkbookIdentityValidationIssue {
  code:
    | "invalid-identity"
    | "invalid-observed-at"
    | "invalid-provider-kind"
    | "invalid-protected-view"
    | "invalid-reason"
    | "invalid-state"
    | "invalid-top-level"
    | "invalid-version"
    | "missing-required-field";
  message: string;
  path: string;
}

export interface ActiveWorkbookIdentityParseResult {
  issues: ActiveWorkbookIdentityValidationIssue[];
  snapshot?: ActiveWorkbookIdentitySnapshot;
}

interface IdentityValidationOptions {
  requireNonAddin?: boolean;
  requireNonEmptyPath?: boolean;
}

export function normalizeWorkbookFullNameForComparison(fullName: string): string {
  const slashNormalized = fullName.trim().replace(/\//g, "\\");
  const withoutTrailingSeparators = trimTrailingSeparators(slashNormalized);

  // Workbook.FullName は Excel/VBA 起源の Windows path として比較する。
  return withoutTrailingSeparators.toLowerCase();
}

export function parseActiveWorkbookIdentitySnapshot(value: unknown): ActiveWorkbookIdentityParseResult {
  if (!isRecord(value)) {
    return {
      issues: [
        {
          code: "invalid-top-level",
          message: "snapshot は object である必要があります",
          path: "$"
        }
      ]
    };
  }

  const issues: ActiveWorkbookIdentityValidationIssue[] = [];
  const version = value.version;
  const providerKind = value.providerKind;
  const state = value.state;
  const observedAt = parseObservedAt(value.observedAt, issues);

  if (version !== ACTIVE_WORKBOOK_IDENTITY_VERSION) {
    issues.push({
      code: "invalid-version",
      message: `version は ${ACTIVE_WORKBOOK_IDENTITY_VERSION} である必要があります`,
      path: "$.version"
    });
  }

  if (providerKind !== ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND) {
    issues.push({
      code: "invalid-provider-kind",
      message: `providerKind は ${ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND} である必要があります`,
      path: "$.providerKind"
    });
  }

  if (state !== "available" && state !== "unavailable" && state !== "protected-view" && state !== "unsupported") {
    issues.push({
      code: "invalid-state",
      message: "state は available / unavailable / protected-view / unsupported のいずれかである必要があります",
      path: "$.state"
    });
    return { issues };
  }

  let snapshot: ActiveWorkbookIdentitySnapshot | undefined;

  switch (state) {
    case "available": {
      const identity = parseIdentity(value.identity, issues, "$.identity", {
        requireNonAddin: true,
        requireNonEmptyPath: true
      });

      if (identity && observedAt && version === ACTIVE_WORKBOOK_IDENTITY_VERSION && providerKind === ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND) {
        snapshot = {
          identity,
          observedAt,
          providerKind: ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND,
          state,
          version: ACTIVE_WORKBOOK_IDENTITY_VERSION
        };
      }
      break;
    }
    case "unavailable": {
      const reason = value.reason;

      if (
        reason !== "host-error" &&
        reason !== "host-unreachable" &&
        reason !== "no-active-workbook" &&
        reason !== "non-workbook-window"
      ) {
        issues.push({
          code: "invalid-reason",
          message: "unavailable reason は host-error / host-unreachable / no-active-workbook / non-workbook-window のいずれかである必要があります",
          path: "$.reason"
        });
      }

      if (
        (reason === "host-error" ||
          reason === "host-unreachable" ||
          reason === "no-active-workbook" ||
          reason === "non-workbook-window") &&
        observedAt &&
        version === ACTIVE_WORKBOOK_IDENTITY_VERSION &&
        providerKind === ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND
      ) {
        snapshot = {
          observedAt,
          providerKind: ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND,
          reason,
          state,
          version: ACTIVE_WORKBOOK_IDENTITY_VERSION
        };
      }
      break;
    }
    case "protected-view": {
      const protectedView = parseProtectedView(value.protectedView, issues);

      if (
        !issues.some((issue) => issue.code === "invalid-protected-view") &&
        observedAt &&
        version === ACTIVE_WORKBOOK_IDENTITY_VERSION &&
        providerKind === ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND
      ) {
        snapshot = {
          ...(protectedView ? { protectedView } : {}),
          observedAt,
          providerKind: ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND,
          state,
          version: ACTIVE_WORKBOOK_IDENTITY_VERSION
        };
      }
      break;
    }
    case "unsupported": {
      const reason = value.reason;

      if (reason !== "addin" && reason !== "unsaved") {
        issues.push({
          code: "invalid-reason",
          message: "unsupported reason は addin または unsaved である必要があります",
          path: "$.reason"
        });
      }

      const identity = parseIdentity(value.identity, issues, "$.identity");

      if (identity && reason === "addin" && identity.isAddin !== true) {
        issues.push({
          code: "invalid-identity",
          message: "unsupported reason=addin の identity.isAddin は true である必要があります",
          path: "$.identity.isAddin"
        });
      }

      if (identity && reason === "unsaved" && identity.path.trim().length > 0) {
        issues.push({
          code: "invalid-identity",
          message: "unsupported reason=unsaved の identity.path は空文字列である必要があります",
          path: "$.identity.path"
        });
      }

      if (
        identity &&
        (reason === "addin" || reason === "unsaved") &&
        observedAt &&
        version === ACTIVE_WORKBOOK_IDENTITY_VERSION &&
        providerKind === ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND
      ) {
        snapshot = {
          identity,
          observedAt,
          providerKind: ACTIVE_WORKBOOK_IDENTITY_PROVIDER_KIND,
          reason,
          state,
          version: ACTIVE_WORKBOOK_IDENTITY_VERSION
        };
      }
      break;
    }
    default:
      break;
  }

  if (issues.length > 0) {
    return { issues };
  }

  return snapshot ? { issues, snapshot } : { issues };
}

function getNonEmptyString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function parseIdentity(
  value: unknown,
  issues: ActiveWorkbookIdentityValidationIssue[],
  basePath: string,
  options?: IdentityValidationOptions
): ActiveWorkbookIdentityFields | undefined {
  if (!isRecord(value)) {
    issues.push({
      code: "invalid-identity",
      message: "identity は object である必要があります",
      path: basePath
    });
    return undefined;
  }

  const fullName = getNonEmptyString(value.fullName);
  const name = getNonEmptyString(value.name);
  const workbookPath = typeof value.path === "string" ? value.path : undefined;
  const isAddin = value.isAddin;

  if (!fullName) {
    issues.push({
      code: "missing-required-field",
      message: "identity.fullName は空でない文字列である必要があります",
      path: `${basePath}.fullName`
    });
  }

  if (!name) {
    issues.push({
      code: "missing-required-field",
      message: "identity.name は空でない文字列である必要があります",
      path: `${basePath}.name`
    });
  }

  if (typeof workbookPath !== "string") {
    issues.push({
      code: "missing-required-field",
      message: "identity.path は文字列である必要があります",
      path: `${basePath}.path`
    });
  }

  if (typeof isAddin !== "boolean") {
    issues.push({
      code: "missing-required-field",
      message: "identity.isAddin は boolean である必要があります",
      path: `${basePath}.isAddin`
    });
  }

  if (typeof workbookPath === "string" && options?.requireNonEmptyPath && workbookPath.trim().length === 0) {
    issues.push({
      code: "invalid-identity",
      message: "identity.path は空文字列を許可しません",
      path: `${basePath}.path`
    });
  }

  if (typeof isAddin === "boolean" && options?.requireNonAddin && isAddin) {
    issues.push({
      code: "invalid-identity",
      message: "identity.isAddin は false である必要があります",
      path: `${basePath}.isAddin`
    });
  }

  if (!fullName || !name || typeof workbookPath !== "string" || typeof isAddin !== "boolean") {
    return undefined;
  }

  return {
    fullName,
    isAddin,
    name,
    path: workbookPath
  };
}

function parseObservedAt(
  value: unknown,
  issues: ActiveWorkbookIdentityValidationIssue[]
): string | undefined {
  if (typeof value !== "string" || value.trim().length === 0 || Number.isNaN(Date.parse(value))) {
    issues.push({
      code: "invalid-observed-at",
      message: "observedAt は ISO 8601 として解釈できる文字列である必要があります",
      path: "$.observedAt"
    });
    return undefined;
  }

  return value;
}

function parseProtectedView(
  value: unknown,
  issues: ActiveWorkbookIdentityValidationIssue[]
): ActiveWorkbookProtectedViewFields | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (!isRecord(value)) {
    issues.push({
      code: "invalid-protected-view",
      message: "protectedView は object である必要があります",
      path: "$.protectedView"
    });
    return undefined;
  }

  const sourceName = value.sourceName === undefined ? undefined : getNonEmptyString(value.sourceName);
  const sourcePath = value.sourcePath === undefined ? undefined : getNonEmptyString(value.sourcePath);

  if (value.sourceName !== undefined && !sourceName) {
    issues.push({
      code: "invalid-protected-view",
      message: "protectedView.sourceName は空でない文字列である必要があります",
      path: "$.protectedView.sourceName"
    });
  }

  if (value.sourcePath !== undefined && !sourcePath) {
    issues.push({
      code: "invalid-protected-view",
      message: "protectedView.sourcePath は空でない文字列である必要があります",
      path: "$.protectedView.sourcePath"
    });
  }

  if (!sourceName && !sourcePath) {
    return undefined;
  }

  return {
    ...(sourceName ? { sourceName } : {}),
    ...(sourcePath ? { sourcePath } : {})
  };
}

function trimTrailingSeparators(value: string): string {
  let normalizedValue = value;

  while (
    normalizedValue.length > 0 &&
    /[\\/]/.test(normalizedValue[normalizedValue.length - 1] ?? "") &&
    !looksLikeWindowsDriveRoot(normalizedValue) &&
    !looksLikeUncRoot(normalizedValue)
  ) {
    normalizedValue = normalizedValue.slice(0, -1);
  }

  return normalizedValue;
}

function looksLikeUncRoot(value: string): boolean {
  return /^\\\\[^\\]+\\[^\\]+\\?$/.test(value);
}

function looksLikeWindowsDriveRoot(value: string): boolean {
  return /^[A-Za-z]:\\?$/.test(value);
}
