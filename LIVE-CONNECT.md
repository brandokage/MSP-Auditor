# Live Microsoft 365 connection

The **Connect Microsoft 365** button on Step 2 reads a tenant directly (read-only) instead of using a CSV.
It appears only when a Microsoft app client ID is set at build time. Without one, the app works as before.

## How it works
1. The browser opens a Microsoft sign-in popup (MSAL Browser v5, PKCE, no client secret). Microsoft sends the result to the `/redirect` page, which passes it back to the main window.
2. The browser calls Microsoft Graph directly. No server of ours is involved.
3. `src/lib/graph.ts` turns the answers into the same rows the CSV parser makes.
4. `auditUsers()` runs, so Steps 3 and 4 work unchanged.

Graph calls (all read-only):

| What | Graph call |
|---|---|
| Users, status, licenses, last sign-in | `GET /users?$select=...,signInActivity` |
| MFA registered | `GET /reports/authenticationMethods/userRegistrationDetails` |
| Global Admins | `GET /roleManagement/directory/roleAssignments?$filter=roleDefinitionId eq '62e90394-...'` |
| License names | `GET /subscribedSkus` |

The MFA report does not list disabled accounts. That is fine: the audit only checks MFA on enabled accounts.

If one call is refused, the audit still runs and a warning says what is missing.

## One-time setup (about 15 minutes)
1. Go to <https://entra.microsoft.com> → **Entra ID** → **App registrations** → **New registration**.
2. Name: `MSP Auditor`. Supported account types: **Accounts in any organizational directory (Multitenant)**.
3. Redirect URI: platform **Single-page application (SPA)**. Add one per place you run the app. It must match exactly:
   - `http://localhost:3000/redirect`
   - Vercel: `https://YOUR-HOST/redirect` (no slash at the end)
   - GitHub Pages: `https://YOUR-USER.github.io/MSP-Auditor/redirect/` (slash at the end)
4. Tap **Register**. Copy the **Application (client) ID**. It is not a secret.
5. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**. Add:
   - `User.Read.All`
   - `AuditLog.Read.All`
   - `RoleManagement.Read.Directory`
   - `LicenseAssignment.Read.All`
6. Tap **Grant admin consent for (your tenant)**.
7. Put the ID in `.env.local` (copy `.env.example`):
   `NEXT_PUBLIC_MSAL_CLIENT_ID=<the client ID>`
   - Vercel: add the same variable in Project Settings → Environment Variables, then redeploy.
   - GitHub Pages: add a repository variable `MSAL_CLIENT_ID` (Settings → Variables). The workflow passes it to the build.
8. `npm run dev`, open Step 2, tap **Connect Microsoft 365**.

## Who can sign in
The person who signs in needs an admin role that can read these reports. **Global Reader** is the best choice: it is read-only and covers every call. Reports Reader or Security Reader also work for the MFA report.
For a tenant other than the one that owns the app registration, an admin of that tenant must approve the permissions at first sign-in.

## Known limits
- **Last sign-in needs Microsoft Entra ID P1** (included in Business Premium). Without it, stale accounts are not checked, and the audit says so.
- Sign-in data can be up to about 24 hours behind. A brand-new account with no sign-in shows as "never signed in".
- Global Admin count covers active, direct role assignments to users. Group-based or PIM-eligible admins are not counted.
- License prices are US list prices from a small table in `graph.ts`. Graph does not return prices.
- Point-in-time scan, not continuous monitoring.

## Security headers (Vercel)
`vercel.json` allows `login.microsoftonline.com` and `graph.microsoft.com` in `connect-src`, and leaves the
`Cross-Origin-Opener-Policy` header off the `/redirect` page only. Microsoft's docs say the redirect bridge page must not have COOP.
GitHub Pages sends no custom headers, so nothing is needed there.

## Tested vs. not tested
- Tested: the mapping, paging, and failure handling, with fake Graph responses (`src/lib/graph.test.ts`).
- **Not tested against live Microsoft** when this was written. Check the first real run with Graph Explorer (<https://developer.microsoft.com/graph/graph-explorer>) if a number looks wrong, and confirm the permission names above in the Graph docs.
