/**
 * Module 2 — Identity, Access & License Waste Auditor.
 * Runs entirely in the browser: the CSV never leaves the user's machine.
 */
import { letterGrade } from "./grading";

export const STALE_DAYS = 90;
export const DEFAULT_LICENSE_COST = 22;

export type AccountType = "Member" | "Guest" | "Shared/Service";

export interface UserRow {
  displayName: string;
  upn: string;
  title: string;
  enabled: boolean;
  mfa: boolean | null;
  lastSignIn: Date | null;
  /** The sign-in cell had a value we couldn't read as a date (e.g. 14/03/2025 after an Excel re-save). */
  signInUnreadable: boolean;
  role: string;
  isGlobalAdmin: boolean;
  license: string;
  isPaid: boolean;
  monthlyCost: number;
  accountType: AccountType;
}

export type FlagCode = "ADMIN_NO_MFA" | "NO_MFA" | "STALE" | "LICENSE_WASTE" | "GLOBAL_ADMIN";

export interface AuditedUser extends UserRow {
  daysSinceSignIn: number | null;
  stale: boolean;
  flags: FlagCode[];
  severity: "critical" | "warning" | "ok";
}

export interface IdentityAudit {
  source: string;
  totalAccounts: number;
  enabledAccounts: number;
  mfaCoveragePct: number;
  missingMfa: AuditedUser[];
  stale: AuditedUser[];
  globalAdmins: AuditedUser[];
  adminsWithoutMfa: AuditedUser[];
  adminSprawl: "ok" | "warning" | "critical";
  wasted: AuditedUser[];
  monthlyWaste: number;
  annualWaste: number;
  score: number;
  grade: string;
  users: AuditedUser[];
  warnings: string[];
}

// ---------------------------------------------------------------------------
// CSV parsing (RFC 4180: quoted fields, escaped quotes, CRLF, BOM)
// ---------------------------------------------------------------------------

/**
 * Excel writes semicolon- or tab-separated "CSV" in many locales. Pick whichever of
 * comma, semicolon or tab appears most often (outside quotes) in the first line.
 */
export function detectDelimiter(firstLine: string): "," | ";" | "\t" {
  const counts = { ",": 0, ";": 0, "\t": 0 };
  let inQuotes = false;
  for (const c of firstLine) {
    if (c === '"') inQuotes = !inQuotes;
    else if (!inQuotes && c in counts) counts[c as keyof typeof counts]++;
  }
  return counts[";"] > counts[","] && counts[";"] >= counts["\t"] ? ";" : counts["\t"] > counts[","] ? "\t" : ",";
}

export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  let s = text.replace(/^\uFEFF/, "");
  // Excel's optional delimiter hint line, e.g. "sep=;"
  const hint = /^sep=(.)\r?\n/i.exec(s);
  if (hint) s = s.slice(hint[0].length);
  const delimiter = hint ? hint[1] : detectDelimiter(s.slice(0, s.search(/\r|\n|$/)));
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (inQuotes) {
      if (c === '"') {
        if (s[i + 1] === '"') {
          field += '"';
          i++;
        } else inQuotes = false;
      } else field += c;
    } else if (c === '"') inQuotes = true;
    else if (c === delimiter) {
      row.push(field);
      field = "";
    } else if (c === "\n" || c === "\r") {
      if (c === "\r" && s[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += c;
  }
  if (field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows.filter((r) => r.some((f) => f.trim() !== ""));
}

export function toCsv(rows: (string | number)[][]): string {
  return rows
    .map((r) => r.map((v) => (/[",\n]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v))).join(","))
    .join("\n");
}

// ---------------------------------------------------------------------------
// Header normalization — Microsoft 365 / Entra ID and Google Workspace exports
// ---------------------------------------------------------------------------

type Field = "displayName" | "upn" | "title" | "enabled" | "blocked" | "mfa" | "lastSignIn" | "role" | "license" | "monthlyCost" | "accountType";

const HEADER_ALIASES: Record<Field, string[]> = {
  displayName: ["displayname", "name", "fullname", "username", "user", "userdisplayname"],
  title: ["jobtitle", "title", "position"],
  upn: ["userprincipalname", "email", "emailaddress", "primaryemail", "upn", "mail", "emailaddressrequired"],
  enabled: ["accountenabled", "status", "enabled", "userstatus", "accountstatus", "statusreadonly"],
  blocked: ["blockcredential", "signinblocked", "blocksignin", "suspended"],
  // "ismfaregistered" is Microsoft Entra's Authentication methods → User registration details report.
  mfa: ["mfastatus", "mfa", "2svenrolled", "2svenforced", "isenrolledin2sv", "mfaenabled", "strongauthentication", "2svenrolledreadonly", "multifactorauthstatus", "permfastate", "ismfaregistered"],
  lastSignIn: ["lastsignindatetime", "lastsignin", "lastlogin", "lastlogintime", "lastsignedin", "lastsignindate", "lastinteractivesignin", "lastsignintimereadonly", "lastactivity"],
  role: ["role", "adminroles", "roles", "adminrole", "directoryroles", "isadmin", "admin"],
  license: ["licenses", "assignedlicense", "assignedlicenses", "license", "licensesku", "skus", "assignedproducts"],
  monthlyCost: ["monthlycost", "costpermonth", "licensecost", "monthlylicensecost", "cost"],
  accountType: ["accounttype", "usertype", "type", "recipienttype"],
};

function normalizeHeader(h: string): string {
  return h.toLowerCase().replace(/\[.*?\]|\(.*?\)/g, "").replace(/[^a-z0-9]/g, "");
}

export function mapHeaders(headers: string[]): { map: Partial<Record<Field, number>>; unmapped: string[] } {
  const map: Partial<Record<Field, number>> = {};
  const unmapped: string[] = [];
  headers.forEach((h, idx) => {
    const n = normalizeHeader(h);
    const field = (Object.keys(HEADER_ALIASES) as Field[]).find((f) => map[f] === undefined && HEADER_ALIASES[f].includes(n));
    if (field) map[field] = idx;
    else unmapped.push(h);
  });
  return { map, unmapped };
}

const truthy = (v: string) => /^(true|yes|y|1|enabled|enforced|enrolled|active|on|registered|capable)$/i.test(v.trim());
const falsy = (v: string) => /^(false|no|n|0|disabled|not enrolled|notenrolled|none|off|suspended|blocked|inactive|deleted|archived|not registered)$/i.test(v.trim());

function parseEnabled(v: string | undefined, blocked: string | undefined): boolean {
  if (blocked !== undefined && blocked.trim() !== "" && truthy(blocked)) return false;
  if (v === undefined || v.trim() === "") return true;
  if (falsy(v)) return false;
  return true;
}

function parseMfa(v: string | undefined): boolean | null {
  if (v === undefined || v.trim() === "") return null;
  if (truthy(v)) return true;
  if (falsy(v)) return false;
  return null;
}

/** null = never signed in; "unreadable" = there is a value, but not one we can trust as a date. */
function parseDate(v: string | undefined): Date | null | "unreadable" {
  if (!v) return null;
  const t = v.trim();
  if (!t || /^(never|never logged in|n\/a|none|-|null)$/i.test(t)) return null;
  // Day-first dates (14/03/2025) and two-digit years are ambiguous; JavaScript would
  // misread or reject them, so refuse them rather than guess.
  if (/^\d{1,2}[/.-]\d{1,2}[/.-]\d{2}(\D|$)/.test(t)) return "unreadable";
  const dm = /^(\d{1,2})[/.](\d{1,2})[/.]\d{4}/.exec(t);
  if (dm && Number(dm[1]) > 12) return "unreadable";
  // Google exports like "2025/03/14 09:12:00 UTC"
  const d = new Date(t.replace(/^(\d{4})\/(\d{2})\/(\d{2})/, "$1-$2-$3").replace(/ UTC$/, "Z"));
  return isNaN(d.getTime()) ? "unreadable" : d;
}

export function parseAccountType(v: string | undefined, upn: string, name: string): AccountType {
  const s = `${v ?? ""}`.toLowerCase();
  if (/guest/.test(s) || /#ext#/i.test(upn)) return "Guest";
  if (/shared|service|room|resource|equipment/.test(s)) return "Shared/Service";
  if (!s && /^(info|admin|office|donations|noreply|no-reply|scanner|frontdesk)@/i.test(upn)) return "Shared/Service";
  if (!s && /shared mailbox|service account/i.test(name)) return "Shared/Service";
  return "Member";
}

export function isGlobalAdminRole(role: string): boolean {
  return /global\s*admin|super\s*admin|company\s*administrator|global\s*administrator|^true$|^yes$/i.test(role.trim());
}

export function isPaidLicense(license: string): boolean {
  const l = license.trim();
  return !!l && !/^(unlicensed|none|n\/a|-|free|no license|0)$/i.test(l);
}

export function parseUsers(text: string): { users: UserRow[]; warnings: string[] } {
  return parseUserRows(parseCsv(text));
}

export function parseUserRows(rows: string[][]): { users: UserRow[]; warnings: string[] } {
  const warnings: string[] = [];
  if (rows.length < 2) throw new Error("The CSV has no data rows. Check that you exported the user list with a header row.");
  const { map, unmapped } = mapHeaders(rows[0]);
  if (map.upn === undefined && map.displayName === undefined) {
    throw new Error(`Couldn't find a name or email column. Detected headers: ${rows[0].join(", ")}`);
  }
  if (map.mfa === undefined) warnings.push("No MFA column detected — MFA coverage cannot be measured.");
  if (map.lastSignIn === undefined) warnings.push("No last sign-in column detected — stale accounts cannot be measured.");
  if (map.license === undefined) warnings.push("No license column detected — license waste cannot be measured.");
  if (map.monthlyCost === undefined && map.license !== undefined) warnings.push(`No monthlyCost column — paid licenses assumed at $${DEFAULT_LICENSE_COST}/mo.`);
  if (unmapped.length) warnings.push(`Ignored columns: ${unmapped.join(", ")}`);
  if (map.role !== undefined && ["isadmin", "admin"].includes(normalizeHeader(rows[0][map.role]))) {
    warnings.push(`The "${rows[0][map.role].trim()}" column marks anyone with any admin role, so the Global Admin count may be too high. Check roles in the admin center.`);
  }

  const get = (r: string[], f: Field) => (map[f] !== undefined ? r[map[f]!] : undefined);
  const identified = rows.slice(1).filter((r) => (get(r, "upn") ?? "").trim() || (get(r, "displayName") ?? "").trim());
  const skipped = rows.length - 1 - identified.length;
  if (skipped) warnings.push(`Skipped ${skipped} row${skipped === 1 ? "" : "s"} with no name or email.`);
  const users = identified.map((r) => {
    const upn = (get(r, "upn") ?? "").trim();
    const displayName = (get(r, "displayName") ?? "").trim() || upn;
    const signIn = parseDate(get(r, "lastSignIn"));
    const role = (get(r, "role") ?? "").trim() || "User";
    const license = (get(r, "license") ?? "").trim();
    const isPaid = isPaidLicense(license);
    const costRaw = get(r, "monthlyCost");
    const parsedCost = costRaw ? Number(costRaw.replace(/[$,\s]/g, "")) : NaN;
    return {
      displayName,
      upn,
      title: (get(r, "title") ?? "").trim(),
      enabled: parseEnabled(get(r, "enabled"), get(r, "blocked")),
      mfa: parseMfa(get(r, "mfa")),
      lastSignIn: signIn === "unreadable" ? null : signIn,
      signInUnreadable: signIn === "unreadable",
      role,
      isGlobalAdmin: isGlobalAdminRole(role),
      license: license || "Unlicensed",
      isPaid,
      monthlyCost: isPaid ? (isNaN(parsedCost) ? DEFAULT_LICENSE_COST : parsedCost) : 0,
      accountType: parseAccountType(get(r, "accountType"), upn, displayName),
    } satisfies UserRow;
  });
  const unreadable = users.filter((u) => u.signInUnreadable).length;
  if (unreadable) {
    warnings.push(
      `${unreadable} sign-in date${unreadable === 1 ? "" : "s"} couldn't be read (for example, day-first dates like 14/03/2025), so ${unreadable === 1 ? "that account was" : "those accounts were"} not checked for staleness. Re-export without opening the file in Excel, or use YYYY-MM-DD dates.`,
    );
  }
  return { users, warnings };
}

// ---------------------------------------------------------------------------
// Combining exports — e.g. the M365 admin center user list + Entra's MFA report
// ---------------------------------------------------------------------------

const FIELDS = Object.keys(HEADER_ALIASES) as Field[];
type MergedRecord = Partial<Record<Field, string>>;

/**
 * Joins several exports into one table on the email column. The first file sets the
 * account order; later files fill in the columns it lacks or left blank. The first
 * non-blank value wins.
 */
export function mergeCsvFiles(files: { name: string; text: string }[]): { rows: string[][]; warnings: string[] } {
  const warnings: string[] = [];
  const header: MergedRecord = {};
  const records: MergedRecord[] = [];
  const byEmail = new Map<string, MergedRecord>();
  const ignored = new Set<string>();
  files.forEach(({ name, text }, i) => {
    const rows = parseCsv(text);
    if (rows.length < 2) throw new Error(`${name} has no data rows. Check that you exported it with a header row.`);
    const { map, unmapped } = mapHeaders(rows[0]);
    if (map.upn === undefined) {
      throw new Error(`${name} has no email column, so it can't be matched with the other files. Detected headers: ${rows[0].join(", ")}`);
    }
    unmapped.forEach((h) => h.trim() && ignored.add(h.trim()));
    for (const f of FIELDS) if (map[f] !== undefined && header[f] === undefined) header[f] = rows[0][map[f]!];
    let added = 0;
    let noEmail = 0;
    for (const r of rows.slice(1)) {
      const key = (r[map.upn] ?? "").trim().toLowerCase();
      if (!key) {
        noEmail++;
        continue;
      }
      let rec = byEmail.get(key);
      if (!rec) {
        rec = {};
        byEmail.set(key, rec);
        records.push(rec);
        if (i > 0) added++;
      }
      for (const f of FIELDS) {
        const v = map[f] !== undefined ? (r[map[f]!] ?? "") : "";
        if (v.trim() && !(rec[f] ?? "").trim()) rec[f] = v;
      }
    }
    if (noEmail) warnings.push(`Skipped ${noEmail} row${noEmail === 1 ? "" : "s"} in ${name} with no email, since ${noEmail === 1 ? "it" : "they"} can't be matched.`);
    if (added) warnings.push(`${added} account${added === 1 ? " appears" : "s appear"} only in ${name}, so some of ${added === 1 ? "its" : "their"} columns are blank.`);
  });
  if (ignored.size) warnings.push(`Ignored columns: ${[...ignored].join(", ")}`);
  const fields = FIELDS.filter((f) => header[f] !== undefined);
  return { rows: [fields.map((f) => header[f]!), ...records.map((rec) => fields.map((f) => rec[f] ?? ""))], warnings };
}

// ---------------------------------------------------------------------------
// Audit rules
// ---------------------------------------------------------------------------

export function auditUsers(users: UserRow[], source: string, warnings: string[] = [], now = new Date()): IdentityAudit {
  const hasSignInData = users.some((u) => u.lastSignIn);
  const audited: AuditedUser[] = users.map((u) => {
    const days = u.lastSignIn ? Math.floor((now.getTime() - u.lastSignIn.getTime()) / 86_400_000) : null;
    const stale = u.enabled && hasSignInData && !u.signInUnreadable && u.accountType !== "Shared/Service" && (days === null || days > STALE_DAYS);
    const flags: FlagCode[] = [];
    if (u.enabled && u.mfa === false) flags.push(u.isGlobalAdmin ? "ADMIN_NO_MFA" : "NO_MFA");
    if (stale) flags.push("STALE");
    if (u.isPaid && (!u.enabled || stale)) flags.push("LICENSE_WASTE");
    if (u.isGlobalAdmin && u.enabled) flags.push("GLOBAL_ADMIN");
    const severity: AuditedUser["severity"] = flags.some((f) => f === "ADMIN_NO_MFA" || f === "NO_MFA" || (f === "STALE" && u.isGlobalAdmin))
      ? "critical"
      : flags.some((f) => f !== "GLOBAL_ADMIN")
        ? "warning"
        : "ok";
    return { ...u, daysSinceSignIn: days, stale, flags, severity };
  });

  const enabled = audited.filter((u) => u.enabled);
  const mfaKnown = enabled.filter((u) => u.mfa !== null && u.accountType !== "Shared/Service");
  const mfaOn = mfaKnown.filter((u) => u.mfa);
  const mfaCoveragePct = mfaKnown.length ? Math.round((mfaOn.length / mfaKnown.length) * 100) : 100;
  const missingMfa = enabled.filter((u) => u.mfa === false);
  const stale = audited.filter((u) => u.stale);
  const globalAdmins = enabled.filter((u) => u.isGlobalAdmin);
  const adminsWithoutMfa = globalAdmins.filter((u) => u.mfa === false);
  const memberCount = Math.max(1, enabled.filter((u) => u.accountType === "Member").length);
  const adminSprawl: IdentityAudit["adminSprawl"] = adminsWithoutMfa.length
    ? "critical"
    : globalAdmins.length > 3 || globalAdmins.length / memberCount > 0.1
      ? "warning"
      : "ok";
  const wasted = audited.filter((u) => u.flags.includes("LICENSE_WASTE"));
  const monthlyWaste = Math.round(wasted.reduce((s, u) => s + u.monthlyCost, 0) * 100) / 100;

  // Score: MFA 40 · stale 20 · admin hygiene 20 · license hygiene 20
  const mfaPts = 40 * (mfaCoveragePct / 100);
  const stalePts = Math.max(0, 20 - stale.length * 4);
  const adminPts = Math.max(0, 20 - (globalAdmins.length > 3 || globalAdmins.length / memberCount > 0.1 ? 6 : 0) - adminsWithoutMfa.length * 10);
  const licensePts = Math.max(0, 20 - wasted.length * 4);
  const score = Math.round(mfaPts + stalePts + adminPts + licensePts);

  return {
    source,
    totalAccounts: audited.length,
    enabledAccounts: enabled.length,
    mfaCoveragePct,
    missingMfa,
    stale,
    globalAdmins,
    adminsWithoutMfa,
    adminSprawl,
    wasted,
    monthlyWaste,
    annualWaste: Math.round(monthlyWaste * 12 * 100) / 100,
    score,
    grade: letterGrade(score),
    users: audited,
    warnings,
  };
}

export function auditCsv(text: string, source: string): IdentityAudit {
  const { users, warnings } = parseUsers(text);
  return auditUsers(users, source, warnings);
}

export function auditCsvFiles(files: { name: string; text: string }[]): IdentityAudit {
  if (files.length === 1) return auditCsv(files[0].text, files[0].name);
  const merged = mergeCsvFiles(files);
  const { users, warnings } = parseUserRows(merged.rows);
  return auditUsers(users, files.map((f) => f.name).join(" + "), [...warnings, ...merged.warnings]);
}
