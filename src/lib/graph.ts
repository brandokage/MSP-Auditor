/**
 * Live Microsoft 365 connection (read-only) via Microsoft Graph.
 *
 * The browser signs in with MSAL and calls Graph directly, so the user directory never
 * touches a server of ours. This file has no React and no MSAL imports: it takes a
 * `GraphGet` function, which makes it easy to test with fake responses.
 *
 * Output is the same `UserRow[]` that the CSV parser produces, so `auditUsers()` and
 * everything after it (Tabs 3 and 4) work unchanged.
 */
import { DEFAULT_LICENSE_COST, parseAccountType, type UserRow } from "./identity";

/** Read-only, delegated permissions. Keep this list in sync with LIVE-CONNECT.md. */
export const GRAPH_SCOPES = [
  "User.Read.All", // users, account status, licenses
  "AuditLog.Read.All", // last sign-in (needs Entra ID P1) and MFA registration report
  "RoleManagement.Read.Directory", // who holds the Global Administrator role
  "LicenseAssignment.Read.All", // license SKU names (subscribedSkus)
];

const GRAPH = "https://graph.microsoft.com/v1.0";
/** Built-in role template ID for "Global Administrator" (the same in every tenant). */
export const GLOBAL_ADMIN_TEMPLATE_ID = "62e90394-69f5-4237-9190-012177145e10";

export class GraphError extends Error {
  constructor(
    public status: number,
    message: string,
  ) {
    super(message);
  }
}

/** Fetches one Graph URL (absolute, or a path starting with "/") and returns parsed JSON. Throws GraphError. */
export type GraphGet = (urlOrPath: string) => Promise<unknown>;

export function makeGraphGet(token: string, fetchImpl: typeof fetch = fetch): GraphGet {
  return async (urlOrPath) => {
    const url = urlOrPath.startsWith("http") ? urlOrPath : `${GRAPH}${urlOrPath}`;
    const res = await fetchImpl(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      let detail = "";
      try {
        const body = (await res.json()) as { error?: { message?: string } };
        detail = body.error?.message ?? "";
      } catch {
        /* no JSON body */
      }
      throw new GraphError(res.status, detail || `Microsoft Graph returned ${res.status}`);
    }
    return res.json();
  };
}

interface Page<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

/** Follows @odata.nextLink until every page is read (capped, so a bad response can't loop forever). */
export async function getAll<T>(get: GraphGet, path: string, maxPages = 50): Promise<T[]> {
  const out: T[] = [];
  let next: string | undefined = path;
  for (let i = 0; next && i < maxPages; i++) {
    const page = (await get(next)) as Page<T>;
    out.push(...(page.value ?? []));
    next = page["@odata.nextLink"];
  }
  return out;
}

// ---------------------------------------------------------------------------
// Raw shapes (only the fields we read)
// ---------------------------------------------------------------------------

export interface GraphUser {
  id: string;
  displayName?: string | null;
  userPrincipalName?: string | null;
  jobTitle?: string | null;
  accountEnabled?: boolean | null;
  userType?: string | null;
  assignedLicenses?: { skuId: string }[];
  signInActivity?: { lastSignInDateTime?: string | null; lastNonInteractiveSignInDateTime?: string | null } | null;
}
export interface GraphRegistration {
  id: string;
  userPrincipalName?: string | null;
  isMfaRegistered?: boolean | null;
}
export interface GraphSku {
  skuId: string;
  skuPartNumber: string;
}

export interface TenantSnapshot {
  users: GraphUser[];
  /** False when sign-in data could not be read (for example, no Entra ID P1). */
  signInIncluded: boolean;
  /** null = the report could not be read. */
  registrations: GraphRegistration[] | null;
  /** null = roles could not be read. */
  globalAdminIds: Set<string> | null;
  /** null = SKUs could not be read. */
  skus: GraphSku[] | null;
  warnings: string[];
}

// ---------------------------------------------------------------------------
// License names and approximate prices
// ---------------------------------------------------------------------------

/** Approximate US list prices per user per month (annual billing). Graph does not return prices. */
const SKU_INFO: Record<string, { name: string; cost: number }> = {
  O365_BUSINESS_ESSENTIALS: { name: "Microsoft 365 Business Basic", cost: 6 },
  O365_BUSINESS_PREMIUM: { name: "Microsoft 365 Business Standard", cost: 12.5 },
  SPB: { name: "Microsoft 365 Business Premium", cost: 22 },
  SPE_E3: { name: "Microsoft 365 E3", cost: 36 },
  SPE_E5: { name: "Microsoft 365 E5", cost: 57 },
  SPE_F1: { name: "Microsoft 365 F3", cost: 8 },
  STANDARDPACK: { name: "Office 365 E1", cost: 10 },
  ENTERPRISEPACK: { name: "Office 365 E3", cost: 23 },
  ENTERPRISEPREMIUM: { name: "Office 365 E5", cost: 38 },
  EXCHANGESTANDARD: { name: "Exchange Online (Plan 1)", cost: 4 },
  EXCHANGEENTERPRISE: { name: "Exchange Online (Plan 2)", cost: 8 },
  AAD_PREMIUM: { name: "Microsoft Entra ID P1", cost: 6 },
  AAD_PREMIUM_P2: { name: "Microsoft Entra ID P2", cost: 9 },
};
/** Free or trial-only SKUs: assigned, but not wasted money. */
const FREE_SKUS = new Set(["FLOW_FREE", "POWER_BI_STANDARD", "TEAMS_EXPLORATORY", "WINDOWS_STORE", "STREAM", "DEVELOPERPACK_E5"]);

// ---------------------------------------------------------------------------
// Pure mapping: Graph data -> UserRow[]
// ---------------------------------------------------------------------------

function latest(...dates: (string | null | undefined)[]): Date | null {
  const times = dates.filter((d): d is string => !!d).map((d) => new Date(d).getTime()).filter((t) => !isNaN(t));
  return times.length ? new Date(Math.max(...times)) : null;
}

export function mapGraphToUsers(snap: TenantSnapshot): { users: UserRow[]; warnings: string[] } {
  const warnings = [...snap.warnings];
  const skuById = new Map((snap.skus ?? []).map((s) => [s.skuId, s.skuPartNumber]));
  const mfaById = new Map((snap.registrations ?? []).map((r) => [r.id, r.isMfaRegistered ?? null]));
  const unpriced = new Set<string>();

  const users = snap.users.map((u): UserRow => {
    const upn = u.userPrincipalName ?? "";
    const displayName = u.displayName || upn;

    const parts = (u.assignedLicenses ?? []).map((l) => skuById.get(l.skuId) ?? l.skuId);
    const names: string[] = [];
    let cost = 0;
    let isPaid = false;
    for (const part of parts) {
      const info = SKU_INFO[part];
      if (FREE_SKUS.has(part)) {
        names.push(part);
      } else if (info) {
        names.push(info.name);
        cost += info.cost;
        isPaid = true;
      } else {
        names.push(part);
        cost += DEFAULT_LICENSE_COST;
        isPaid = true;
        unpriced.add(part);
      }
    }

    // Graph returns accountEnabled=true/false. Treat "unknown" as enabled so nobody is silently skipped.
    const signIn = snap.signInIncluded ? latest(u.signInActivity?.lastSignInDateTime, u.signInActivity?.lastNonInteractiveSignInDateTime) : null;

    return {
      displayName,
      upn,
      title: u.jobTitle ?? "",
      enabled: u.accountEnabled !== false,
      mfa: snap.registrations ? (mfaById.get(u.id) ?? null) : null,
      lastSignIn: signIn,
      signInUnreadable: false,
      role: snap.globalAdminIds?.has(u.id) ? "Global Admin" : "User",
      isGlobalAdmin: !!snap.globalAdminIds?.has(u.id),
      license: names.length ? names.join(" + ") : "Unlicensed",
      isPaid,
      monthlyCost: Math.round(cost * 100) / 100,
      accountType: parseAccountType(u.userType === "Guest" ? "Guest" : undefined, upn, displayName),
    };
  });

  if (unpriced.size) {
    warnings.push(`No price on file for ${[...unpriced].join(", ")}, so $${DEFAULT_LICENSE_COST}/mo was assumed for each.`);
  }
  if (users.some((u) => u.isPaid)) {
    warnings.push("License prices are US list prices. Nonprofit discounts or your contract may lower them.");
  }
  return { users, warnings };
}

// ---------------------------------------------------------------------------
// Fetching: several Graph calls, each allowed to fail on its own (with a plain warning)
// ---------------------------------------------------------------------------

const USER_FIELDS = "id,displayName,userPrincipalName,jobTitle,accountEnabled,userType,assignedLicenses";

function describeFailure(e: unknown): string {
  if (e instanceof GraphError) {
    if (e.status === 403) return "Microsoft refused (403). The signed-in account may lack the required admin role, or admin consent is missing.";
    return `Microsoft returned ${e.status}: ${e.message}`;
  }
  return e instanceof Error ? e.message : "unknown error";
}

export async function fetchTenantSnapshot(get: GraphGet): Promise<TenantSnapshot> {
  const warnings: string[] = [];

  // 1. Users with sign-in activity. Without Entra ID P1 this fails; fall back to users only.
  let users: GraphUser[];
  let signInIncluded = true;
  try {
    // Graph limits $top to 500 when signInActivity is selected.
    users = await getAll<GraphUser>(get, `/users?$select=${USER_FIELDS},signInActivity&$top=500`);
  } catch (e) {
    if (e instanceof GraphError && (e.status === 400 || e.status === 403)) {
      signInIncluded = false;
      warnings.push("Last sign-in could not be read (it needs Microsoft Entra ID P1 and the AuditLog.Read.All permission), so stale accounts were not checked.");
      users = await getAll<GraphUser>(get, `/users?$select=${USER_FIELDS}&$top=999`);
    } else {
      throw new GraphError(e instanceof GraphError ? e.status : 0, `Could not read the user list. ${describeFailure(e)}`);
    }
  }

  // 2. MFA registration report.
  let registrations: GraphRegistration[] | null = null;
  try {
    registrations = await getAll<GraphRegistration>(get, "/reports/authenticationMethods/userRegistrationDetails?$top=999");
  } catch (e) {
    warnings.push(`MFA registration could not be read, so MFA coverage was not measured. ${describeFailure(e)}`);
  }

  // 3. Global Administrators (direct, active assignments to users).
  let globalAdminIds: Set<string> | null = null;
  try {
    const assignments = await getAll<{ principalId: string; principal?: { "@odata.type"?: string } }>(
      get,
      `/roleManagement/directory/roleAssignments?$filter=roleDefinitionId eq '${GLOBAL_ADMIN_TEMPLATE_ID}'&$expand=principal`,
    );
    globalAdminIds = new Set(assignments.filter((a) => !a.principal?.["@odata.type"] || a.principal["@odata.type"] === "#microsoft.graph.user").map((a) => a.principalId));
    warnings.push("Global Admin count covers active, direct assignments. Admins who get the role through a group or on request (PIM) are not included.");
  } catch (e) {
    warnings.push(`Admin roles could not be read, so the Global Admin count is missing. ${describeFailure(e)}`);
  }

  // 4. License names.
  let skus: GraphSku[] | null = null;
  try {
    skus = await getAll<GraphSku>(get, "/subscribedSkus");
  } catch (e) {
    warnings.push(`License names could not be read, so license waste uses an assumed price. ${describeFailure(e)}`);
  }

  return { users, signInIncluded, registrations, globalAdminIds, skus, warnings };
}

/** Convenience: fetch + map. */
export async function loadLiveUsers(get: GraphGet) {
  return mapGraphToUsers(await fetchTenantSnapshot(get));
}
