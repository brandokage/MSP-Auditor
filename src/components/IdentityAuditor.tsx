"use client";

import { Cloud, Crown, Download, FileSpreadsheet, FileUp, KeyRound, PiggyBank, Search, ShieldOff, UserX, Users } from "lucide-react";
import { useMemo, useRef, useState } from "react";
import { downloadText } from "@/lib/download";
import { formatUsd } from "@/lib/grading";
import { makeGraphGet, loadLiveUsers } from "@/lib/graph";
import { STALE_DAYS, auditCsvFiles, auditUsers, type AuditedUser, type FlagCode, type IdentityAudit } from "@/lib/identity";
import { describeSignInError, isLiveConnectConfigured, signInForGraph } from "@/lib/msalClient";
import { buildBlankTemplateCsv, buildSampleCsv } from "@/lib/sampleData";
import { Button, Card, EmptyState, GradeBadge, Kpi, SectionTitle, cx } from "./ui";

type Filter = "flagged" | "all" | "mfa" | "stale" | "admin" | "waste";

const FLAG_STYLE: Record<FlagCode, { label: string; cls: string }> = {
  ADMIN_NO_MFA: { label: "Admin · No MFA", cls: "bg-brick-600 text-white ring-brick-600" },
  NO_MFA: { label: "No MFA", cls: "bg-brick-50 text-brick-700 ring-brick-200" },
  STALE: { label: `Stale (${STALE_DAYS}+ days)`, cls: "bg-ochre-50 text-ochre-800 ring-ochre-200" },
  LICENSE_WASTE: { label: "License waste", cls: "bg-violet-50 text-violet-700 ring-violet-200" },
  GLOBAL_ADMIN: { label: "Global Admin", cls: "bg-stone-100 text-stone-700 ring-stone-300" },
};

// What each column is used for. Everything else in an export is ignored, so people can delete it first.
const COLUMN_GUIDE: { column: string; usedFor: string; need: "Required" | "Recommended" | "Optional" }[] = [
  { column: "Email or name", usedFor: "Telling accounts apart and naming them in the ticket", need: "Required" },
  { column: "MFA status", usedFor: "MFA coverage", need: "Recommended" },
  { column: "Last sign-in", usedFor: "Finding stale accounts", need: "Recommended" },
  { column: "Role / admin roles", usedFor: "Counting Global Admins", need: "Recommended" },
  { column: "Enabled / blocked", usedFor: "Skipping accounts that are already off", need: "Optional" },
  { column: "Licenses", usedFor: "License waste", need: "Optional" },
  { column: "Monthly cost", usedFor: "Dollar savings (otherwise $22/mo is assumed)", need: "Optional" },
  { column: "Account type", usedFor: "Leaving guests and shared mailboxes out of staff counts", need: "Optional" },
];

export function IdentityAuditor({
  audit,
  onAudit,
  onLoadSample,
}: {
  audit: IdentityAudit | null;
  onAudit: (a: IdentityAudit) => void;
  onLoadSample: () => void;
}) {
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [filter, setFilter] = useState<Filter>("flagged");
  const [query, setQuery] = useState("");
  const fileInput = useRef<HTMLInputElement>(null);
  const [connecting, setConnecting] = useState(false);
  const liveAvailable = isLiveConnectConfigured();

  /** Live mode: sign in to Microsoft 365 in a popup, read the tenant (read-only), then run the same audit as a CSV. */
  async function handleConnect() {
    setError(null);
    setConnecting(true);
    try {
      const { token, account } = await signInForGraph();
      const { users, warnings } = await loadLiveUsers(makeGraphGet(token));
      onAudit(auditUsers(users, `Microsoft 365 (live) · ${account}`, warnings));
      setFilter("flagged");
    } catch (e) {
      setError(describeSignInError(e) ?? (e instanceof Error ? e.message : "Could not connect to Microsoft 365."));
    } finally {
      setConnecting(false);
    }
  }

  async function handleFiles(list: FileList | null | undefined) {
    const files = Array.from(list ?? []);
    if (!files.length) return;
    setError(null);
    if (files.some((f) => !/\.csv$/i.test(f.name) && f.type !== "text/csv")) {
      setError("Please choose .csv files exported from Microsoft 365 or Google Workspace.");
      return;
    }
    try {
      onAudit(auditCsvFiles(await Promise.all(files.map(async (f) => ({ name: f.name, text: await f.text() })))));
      setFilter("flagged");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not read that CSV.");
    }
  }

  const rows = useMemo(() => {
    if (!audit) return [];
    const q = query.trim().toLowerCase();
    const match: Record<Filter, (u: AuditedUser) => boolean> = {
      flagged: (u) => u.severity !== "ok",
      all: () => true,
      mfa: (u) => u.flags.includes("NO_MFA") || u.flags.includes("ADMIN_NO_MFA"),
      stale: (u) => u.stale,
      admin: (u) => u.isGlobalAdmin && u.enabled,
      waste: (u) => u.flags.includes("LICENSE_WASTE"),
    };
    const rank = { critical: 0, warning: 1, ok: 2 };
    return audit.users
      .filter(match[filter])
      .filter((u) => !q || `${u.displayName} ${u.upn} ${u.title} ${u.role}`.toLowerCase().includes(q))
      .sort(
        (a, b) =>
          rank[a.severity] - rank[b.severity] ||
          Number(b.flags.includes("ADMIN_NO_MFA")) - Number(a.flags.includes("ADMIN_NO_MFA")) ||
          b.flags.length - a.flags.length ||
          a.displayName.localeCompare(b.displayName),
      );
  }, [audit, filter, query]);

  const filters: { key: Filter; label: string; count: number }[] = audit
    ? [
        { key: "flagged", label: "Flagged", count: audit.users.filter((u) => u.severity !== "ok").length },
        { key: "mfa", label: "Missing MFA", count: audit.missingMfa.length },
        { key: "stale", label: "Stale", count: audit.stale.length },
        { key: "admin", label: "Global Admins", count: audit.globalAdmins.length },
        { key: "waste", label: "License Waste", count: audit.wasted.length },
        { key: "all", label: "All", count: audit.totalAccounts },
      ]
    : [];

  return (
    <div className="space-y-6">
      <SectionTitle
        step={2}
        title="Users & licenses"
        subtitle="Add a user export from Microsoft 365 (Entra ID) or Google Workspace. You'll see who has no MFA, which ex-staff accounts are still on, how many people hold admin rights, and what unused licenses cost you."
        right={
          <div className="flex flex-wrap gap-2">
            <Button variant="secondary" onClick={() => downloadText("sample-hope-valley-users.csv", buildSampleCsv())}>
              <Download className="h-4 w-4" /> Sample CSV
            </Button>
            <Button variant="ghost" onClick={() => downloadText("blank-user-import-template.csv", buildBlankTemplateCsv())}>
              <FileSpreadsheet className="h-4 w-4" /> Blank template
            </Button>
          </div>
        }
      />

      <div className="grid gap-6 md:grid-cols-2">
        <section>
          <h3 className="font-serif text-xl font-semibold text-ink-900">Where your file goes</h3>
          <p className="mt-2 text-sm leading-relaxed text-stone-700">
            Nowhere. The file is read by code running in this browser tab and kept in memory. Nothing is uploaded, saved or logged, and refreshing the page erases it.
          </p>
          <p className="mt-3 text-sm font-medium text-ink-900">You don&apos;t have to take my word for it:</p>
          <ul className="mt-1 list-disc space-y-1 pl-5 text-sm leading-relaxed text-stone-700">
            <li>
              Open your browser&apos;s developer tools, go to the <b className="font-medium">Network</b> tab, then add the file. No request is sent.
            </li>
            <li>Or load this page, turn off Wi-Fi, then add the file. The audit still runs.</li>
          </ul>
        </section>
        <section>
          <h3 className="font-serif text-xl font-semibold text-ink-900">What it actually needs</h3>
          <table className="mt-2 w-full text-left text-sm">
            <tbody className="divide-y divide-stone-200">
              {COLUMN_GUIDE.map((c) => (
                <tr key={c.column}>
                  <td className="py-1.5 pr-3 font-medium text-ink-900">{c.column}</td>
                  <td className="py-1.5 pr-3 text-stone-600">{c.usedFor}</td>
                  <td className={cx("py-1.5 text-right text-xs whitespace-nowrap", c.need === "Required" ? "font-semibold text-brick-700" : "text-stone-500")}>{c.need}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="mt-3 text-sm leading-relaxed text-stone-700">
            Delete every other column (phone, address, department, manager) before you add the file. You can also swap real names and emails for placeholders like{" "}
            <code className="font-mono text-[13px]">user01@yourorg.org</code> and match them up yourself afterward. If you do, keep the account type column, or leave shared mailboxes like{" "}
            <code className="font-mono text-[13px]">info@</code> as they are, so they aren&apos;t counted as people.
          </p>
        </section>
      </div>

      {liveAvailable && (
        <div className="flex flex-col gap-3 rounded-md border border-moss-300 bg-moss-50/50 p-5 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <h3 className="flex items-center gap-2 font-serif text-lg font-semibold text-ink-900">
              <Cloud className="h-5 w-5 text-moss-700" /> Connect Microsoft 365 (live, read-only)
            </h3>
            <p className="mt-1 max-w-2xl text-sm leading-relaxed text-stone-700">
              Sign in as an admin and the audit reads the tenant directly: users, MFA registration, last sign-in, Global Admins and licenses. It can only read. Data goes from Microsoft to this browser tab and nowhere else, and closing the tab erases it.
            </p>
          </div>
          <Button onClick={handleConnect} disabled={connecting}>
            <Cloud className="h-4 w-4" /> {connecting ? "Connecting…" : "Connect Microsoft 365"}
          </Button>
        </div>
      )}

      <div
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragging(false);
          handleFiles(e.dataTransfer.files);
        }}
        onClick={() => fileInput.current?.click()}
        onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && fileInput.current?.click()}
        role="button"
        tabIndex={0}
        className={cx(
          "flex cursor-pointer flex-col items-center justify-center gap-2 rounded-md border-2 border-dashed px-6 py-8 text-center transition",
          dragging ? "border-moss-600 bg-moss-50" : "border-stone-400 bg-paper hover:border-moss-500 hover:bg-moss-50/40",
        )}
      >
        <FileUp className={cx("h-7 w-7", dragging ? "text-moss-700" : "text-stone-400")} />
        <div className="text-sm font-medium text-ink-900">Drop a user export (.csv) here, or click to choose one or more</div>
        <div className="text-xs text-stone-500">Stays in this tab. Microsoft 365 and Google Workspace column names are recognized automatically. Several files are matched up by email.</div>
        <input ref={fileInput} type="file" accept=".csv,text/csv" multiple className="hidden" onChange={(e) => handleFiles(e.target.files)} />
      </div>
      <p className="-mt-3 text-sm text-stone-600">
        Just looking?{" "}
        <button onClick={onLoadSample} className="font-medium text-moss-700 underline decoration-moss-300 underline-offset-4 hover:text-moss-900">
          Use the sample nonprofit
        </button>{" "}
        (22 made-up accounts).
      </p>

      {error && <div className="rounded-md border border-brick-200 bg-brick-50 p-4 text-sm text-brick-800">{error}</div>}

      {!audit ? (
        <EmptyState icon={<Users className="h-6 w-6" />} title="No user directory loaded yet">
          <b>Google Workspace:</b> Admin console → Directory → Users → Download users. This includes 2-Step Verification and last sign-in.
          <br />
          <b>Microsoft 365:</b> the standard user download doesn&apos;t include MFA or last sign-in. Add it together with Entra admin center → Protection → Authentication methods → User registration details (for MFA) and they&apos;re matched up by email. Or ask your MSP for one export with MFA status, last sign-in, roles, and licenses (the blank template shows the columns).
        </EmptyState>
      ) : (
        <div className="animate-fade-up space-y-6" key={audit.source + audit.totalAccounts}>
          <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-[repeat(4,1fr)_auto]">
            <Kpi
              label="MFA coverage"
              value={`${audit.mfaCoveragePct}%`}
              sub={`${audit.missingMfa.length} enabled account${audit.missingMfa.length === 1 ? "" : "s"} without MFA`}
              tone={audit.mfaCoveragePct === 100 ? "good" : audit.mfaCoveragePct >= 90 ? "ok" : "bad"}
              icon={<KeyRound className="h-4 w-4" />}
            />
            <Kpi
              label="Stale accounts"
              value={audit.stale.length}
              sub={`Enabled, with no sign-in for more than ${STALE_DAYS} days`}
              tone={audit.stale.length === 0 ? "good" : audit.stale.length <= 1 ? "ok" : "bad"}
              icon={<UserX className="h-4 w-4" />}
            />
            <Kpi
              label="Global Admins"
              value={audit.globalAdmins.length}
              sub={audit.adminsWithoutMfa.length ? `${audit.adminsWithoutMfa.length} without MFA (critical)` : audit.adminSprawl === "warning" ? "More than recommended (3 or fewer)" : "Within the recommended range"}
              tone={audit.adminSprawl === "critical" ? "bad" : audit.adminSprawl === "warning" ? "ok" : "good"}
              icon={<Crown className="h-4 w-4" />}
            />
            <Kpi
              label="Yearly license savings"
              value={formatUsd(audit.annualWaste)}
              sub={`${formatUsd(audit.monthlyWaste)} a month across ${audit.wasted.length} inactive account${audit.wasted.length === 1 ? "" : "s"}`}
              tone={audit.annualWaste === 0 ? "good" : "ok"}
              icon={<PiggyBank className="h-4 w-4" />}
            />
            <Card className="flex items-center justify-center p-5 sm:col-span-2 xl:col-span-1">
              <div className="text-center">
                <div className="mb-2 text-xs font-semibold tracking-wide text-stone-500 uppercase">Identity grade</div>
                <GradeBadge grade={audit.grade} size="lg" score={audit.score} />
              </div>
            </Card>
          </div>

          {audit.adminsWithoutMfa.length > 0 && (
            <div className="flex items-start gap-3 rounded-md border border-brick-300 bg-brick-50 p-4 text-sm text-brick-900">
              <ShieldOff className="mt-0.5 h-5 w-5 shrink-0 text-brick-600" />
              <div>
                <b>Critical:</b> {audit.adminsWithoutMfa.map((u) => `${u.displayName} (${u.upn})`).join(", ")} {audit.adminsWithoutMfa.length > 1 ? "hold" : "holds"} Global Admin rights with MFA disabled. One stolen password gives an attacker full control of every mailbox and file in the organization.
              </div>
            </div>
          )}

          <Card className="overflow-hidden">
            <div className="flex flex-wrap items-center justify-between gap-3 border-b border-stone-100 px-5 py-3">
              <div className="flex flex-wrap gap-1.5">
                {filters.map((f) => (
                  <button
                    key={f.key}
                    onClick={() => setFilter(f.key)}
                    className={cx(
                      "rounded px-3 py-1.5 text-xs font-semibold transition",
                      filter === f.key ? "bg-ink-900 text-white" : "bg-stone-100 text-stone-600 hover:bg-stone-200",
                    )}
                  >
                    {f.label} <span className="ml-1 tabular-nums opacity-70">{f.count}</span>
                  </button>
                ))}
              </div>
              <label className="relative">
                <span className="sr-only">Search accounts</span>
                <Search className="pointer-events-none absolute top-1/2 left-2.5 h-3.5 w-3.5 -translate-y-1/2 text-stone-400" />
                <input
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                  placeholder="Search name, email, role…"
                  className="w-56 rounded border border-stone-200 py-1.5 pr-3 pl-8 text-xs outline-none focus:border-moss-500"
                />
              </label>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full min-w-[860px] text-left text-sm">
                <thead className="bg-stone-50 text-[11px] font-semibold tracking-wide text-stone-500 uppercase">
                  <tr>
                    <th className="px-5 py-2.5">Account</th>
                    <th className="px-3 py-2.5">Status</th>
                    <th className="px-3 py-2.5">MFA</th>
                    <th className="px-3 py-2.5">Last Sign-In</th>
                    <th className="px-3 py-2.5">Role</th>
                    <th className="px-3 py-2.5">License</th>
                    <th className="px-5 py-2.5">Findings</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-stone-100">
                  {rows.map((u) => (
                    <tr key={u.upn + u.displayName} className={cx(u.severity === "critical" ? "bg-brick-50/60" : u.severity === "warning" ? "bg-ochre-50/50" : "")}>
                      <td className="px-5 py-2.5">
                        <div className="font-semibold text-stone-900">{u.displayName}</div>
                        <div className="text-xs text-stone-500">
                          {u.upn}
                          {u.title && <> · {u.title}</>}
                        </div>
                      </td>
                      <td className="px-3 py-2.5">
                        <span className={cx("text-xs font-semibold", u.enabled ? "text-stone-700" : "text-stone-400")}>{u.enabled ? "Enabled" : "Disabled"}</span>
                        {u.accountType !== "Member" && <div className="text-[11px] text-stone-400">{u.accountType}</div>}
                      </td>
                      <td className="px-3 py-2.5">
                        <span className={cx("text-xs font-semibold", u.mfa === false ? "text-brick-600" : u.mfa ? "text-moss-700" : "text-stone-400")}>
                          {u.mfa === null ? "Unknown" : u.mfa ? "On" : "Off"}
                        </span>
                      </td>
                      <td className={cx("px-3 py-2.5 text-xs tabular-nums", u.stale ? "font-semibold text-ochre-700" : "text-stone-600")}>
                        {u.lastSignIn ? `${u.lastSignIn.toLocaleDateString()} (${u.daysSinceSignIn}d)` : u.signInUnreadable ? "Unreadable date" : "Never"}
                      </td>
                      <td className={cx("px-3 py-2.5 text-xs", u.isGlobalAdmin ? "font-semibold text-stone-900" : "text-stone-600")}>{u.role}</td>
                      <td className="px-3 py-2.5 text-xs text-stone-600">
                        {u.license}
                        {u.isPaid && <span className="ml-1 text-stone-400 tabular-nums">({formatUsd(u.monthlyCost)}/mo)</span>}
                      </td>
                      <td className="px-5 py-2.5">
                        <div className="flex flex-wrap gap-1">
                          {u.flags.map((f) => (
                            <span key={f} className={cx("rounded-md px-1.5 py-0.5 text-[10.5px] font-semibold whitespace-nowrap ring-1 ring-inset", FLAG_STYLE[f].cls)}>
                              {FLAG_STYLE[f].label}
                            </span>
                          ))}
                          {u.flags.length === 0 && <span className="text-xs text-moss-600">✓ Clean</span>}
                        </div>
                      </td>
                    </tr>
                  ))}
                  {rows.length === 0 && (
                    <tr>
                      <td colSpan={7} className="px-5 py-8 text-center text-sm text-stone-500">
                        No accounts match this filter.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <div className="flex flex-wrap items-center justify-between gap-2 border-t border-stone-100 bg-stone-50/60 px-5 py-2.5 text-xs text-stone-500">
              <span>
                Source: <b className="text-stone-700">{audit.source}</b> · {audit.totalAccounts} accounts · {audit.enabledAccounts} enabled
              </span>
            </div>
          </Card>

          {audit.warnings.length > 0 && (
            <ul className="space-y-1 rounded-md bg-stone-100 px-4 py-3 text-xs text-stone-600">
              {audit.warnings.map((w) => (
                <li key={w}>• {w}</li>
              ))}
            </ul>
          )}
        </div>
      )}
    </div>
  );
}
