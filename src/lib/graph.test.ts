import { describe, expect, it } from "vitest";
import { GLOBAL_ADMIN_TEMPLATE_ID, GraphError, fetchTenantSnapshot, getAll, mapGraphToUsers, type GraphGet, type GraphUser } from "./graph";
import { auditUsers } from "./identity";

const NOW = new Date("2026-11-05T12:00:00Z");
const iso = (daysAgo: number) => new Date(NOW.getTime() - daysAgo * 86_400_000).toISOString();

const SKUS = [
  { skuId: "sku-premium", skuPartNumber: "SPB" },
  { skuId: "sku-basic", skuPartNumber: "O365_BUSINESS_ESSENTIALS" },
  { skuId: "sku-free", skuPartNumber: "FLOW_FREE" },
  { skuId: "sku-weird", skuPartNumber: "SOME_NEW_SKU" },
];

const USERS: GraphUser[] = [
  { id: "u1", displayName: "Ada Admin", userPrincipalName: "ada@org.org", accountEnabled: true, userType: "Member", assignedLicenses: [{ skuId: "sku-premium" }], signInActivity: { lastSignInDateTime: iso(1) } },
  { id: "u2", displayName: "Ben Staff", userPrincipalName: "ben@org.org", accountEnabled: true, userType: "Member", assignedLicenses: [{ skuId: "sku-basic" }, { skuId: "sku-free" }], signInActivity: { lastSignInDateTime: iso(400), lastNonInteractiveSignInDateTime: iso(3) } },
  { id: "u3", displayName: "Cara Gone", userPrincipalName: "cara@org.org", accountEnabled: true, userType: "Member", assignedLicenses: [{ skuId: "sku-premium" }], signInActivity: { lastSignInDateTime: iso(200) } },
  { id: "u4", displayName: "Dan Disabled", userPrincipalName: "dan@org.org", accountEnabled: false, userType: "Member", assignedLicenses: [{ skuId: "sku-premium" }], signInActivity: null },
  { id: "u5", displayName: "Gus Guest", userPrincipalName: "gus_gmail.com#EXT#@org.onmicrosoft.com", accountEnabled: true, userType: "Guest", assignedLicenses: [] },
  { id: "u6", displayName: "Weird Licensed", userPrincipalName: "weird@org.org", accountEnabled: true, userType: "Member", assignedLicenses: [{ skuId: "sku-weird" }], signInActivity: { lastSignInDateTime: iso(2) } },
];
const REGS = [
  { id: "u1", isMfaRegistered: true },
  { id: "u2", isMfaRegistered: false },
  { id: "u3", isMfaRegistered: true },
  { id: "u4", isMfaRegistered: false },
  { id: "u6", isMfaRegistered: true },
];

/** A fake Graph. `fail` maps a URL fragment to a status code to simulate a refused call. */
function fakeGraph(fail: Record<string, number> = {}): GraphGet {
  return async (url) => {
    for (const [frag, status] of Object.entries(fail)) if (url.includes(frag)) throw new GraphError(status, "refused");
    if (url.includes("/users?")) {
      // Two pages, to prove paging works.
      if (url.includes("page2")) return { value: USERS.slice(3) };
      return { value: USERS.slice(0, 3), "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?page2" };
    }
    if (url.includes("userRegistrationDetails")) return { value: REGS };
    if (url.includes("roleAssignments")) {
      expect(url).toContain(GLOBAL_ADMIN_TEMPLATE_ID);
      return { value: [{ principalId: "u1", principal: { "@odata.type": "#microsoft.graph.user" } }, { principalId: "grp1", principal: { "@odata.type": "#microsoft.graph.group" } }] };
    }
    if (url.includes("subscribedSkus")) return { value: SKUS };
    throw new Error(`unexpected URL ${url}`);
  };
}

describe("getAll", () => {
  it("follows nextLink across pages", async () => {
    const all = await getAll<GraphUser>(fakeGraph(), "/users?$select=id");
    expect(all.map((u) => u.id)).toEqual(["u1", "u2", "u3", "u4", "u5", "u6"]);
  });
});

describe("live Microsoft 365 -> audit", () => {
  it("maps a full tenant and the audit flags the right accounts", async () => {
    const snap = await fetchTenantSnapshot(fakeGraph());
    const { users, warnings } = mapGraphToUsers(snap);
    const byUpn = Object.fromEntries(users.map((u) => [u.upn, u]));

    expect(byUpn["ada@org.org"]).toMatchObject({ isGlobalAdmin: true, role: "Global Admin", mfa: true, license: "Microsoft 365 Business Premium", monthlyCost: 22 });
    // Uses the newer of interactive and non-interactive sign-in.
    expect(byUpn["ben@org.org"].lastSignIn?.toISOString()).toBe(iso(3));
    // A free SKU is listed but not billed.
    expect(byUpn["ben@org.org"]).toMatchObject({ license: "Microsoft 365 Business Basic + FLOW_FREE", monthlyCost: 6, isPaid: true });
    expect(byUpn["dan@org.org"]).toMatchObject({ enabled: false, isPaid: true });
    expect(byUpn["gus_gmail.com#EXT#@org.onmicrosoft.com"]).toMatchObject({ accountType: "Guest", license: "Unlicensed", isPaid: false, mfa: null });
    // Group principals are not counted as admins.
    expect(users.filter((u) => u.isGlobalAdmin)).toHaveLength(1);
    // Unknown SKU: assumed price, and the user is told.
    expect(byUpn["weird@org.org"].monthlyCost).toBe(22);
    expect(warnings.join(" ")).toMatch(/SOME_NEW_SKU/);

    const audit = auditUsers(users, "Microsoft 365 (live)", warnings, NOW);
    expect(audit.missingMfa.map((u) => u.upn)).toEqual(["ben@org.org"]);
    // Cara (200 days) and the never-signed-in guest are stale. Ben's non-interactive sign-in keeps him fresh.
    expect(audit.stale.map((u) => u.upn).sort()).toEqual(["cara@org.org", "gus_gmail.com#EXT#@org.onmicrosoft.com"]);
    expect(audit.wasted.map((u) => u.upn).sort()).toEqual(["cara@org.org", "dan@org.org"]);
    expect(audit.monthlyWaste).toBe(44);
    expect(audit.globalAdmins).toHaveLength(1);
  });

  it("falls back and warns when sign-in data needs Entra ID P1", async () => {
    const snap = await fetchTenantSnapshot(fakeGraph({ "signInActivity": 403 }));
    expect(snap.signInIncluded).toBe(false);
    const { users, warnings } = mapGraphToUsers(snap);
    expect(users.every((u) => u.lastSignIn === null)).toBe(true);
    expect(warnings.join(" ")).toMatch(/Entra ID P1/);
    // No sign-in data anywhere -> nobody is called stale by mistake.
    expect(auditUsers(users, "x", warnings, NOW).stale).toHaveLength(0);
  });

  it("keeps going when MFA, roles or SKUs cannot be read, and says so", async () => {
    const snap = await fetchTenantSnapshot(fakeGraph({ userRegistrationDetails: 403, roleAssignments: 403, subscribedSkus: 403 }));
    const { users, warnings } = mapGraphToUsers(snap);
    expect(users.every((u) => u.mfa === null)).toBe(true);
    expect(users.some((u) => u.isGlobalAdmin)).toBe(false);
    expect(warnings.filter((w) => /could not be read/.test(w))).toHaveLength(3);
  });

  it("throws a clear error when the user list itself cannot be read", async () => {
    await expect(fetchTenantSnapshot(fakeGraph({ "/users?": 401 }))).rejects.toThrow(/Could not read the user list/);
  });
});
