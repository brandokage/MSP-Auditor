/**
 * Browser-only Microsoft sign-in for the live connection (MSAL Browser v5).
 * The client ID is not a secret. Set NEXT_PUBLIC_MSAL_CLIENT_ID at build time (see LIVE-CONNECT.md).
 * Tokens are kept in sessionStorage, so closing the tab signs the user out.
 */
import { PublicClientApplication } from "@azure/msal-browser";
import { GRAPH_SCOPES } from "./graph";

export const MSAL_CLIENT_ID = process.env.NEXT_PUBLIC_MSAL_CLIENT_ID ?? "";
const BASE_PATH = process.env.NEXT_PUBLIC_BASE_PATH ?? "";
// The GitHub Pages export uses trailing slashes (/redirect/); Vercel does not (/redirect).
const TRAILING_SLASH = process.env.NEXT_PUBLIC_TRAILING_SLASH === "1" ? "/" : "";

export const isLiveConnectConfigured = () => MSAL_CLIENT_ID.length > 0;

/** The exact redirect URI to register in Entra for the current site. */
export const redirectUri = () => `${window.location.origin}${BASE_PATH}/redirect${TRAILING_SLASH}`;

let app: PublicClientApplication | null = null;

async function getApp(): Promise<PublicClientApplication> {
  if (!app) {
    app = new PublicClientApplication({
      auth: {
        clientId: MSAL_CLIENT_ID,
        // Work and school accounts from any organization. Personal Microsoft accounts have no tenant to audit.
        authority: "https://login.microsoftonline.com/organizations",
        redirectUri: redirectUri(),
      },
      cache: { cacheLocation: "sessionStorage" },
    });
    await app.initialize();
  }
  return app;
}

/** Opens the Microsoft sign-in popup and returns a Graph access token for the read-only scopes. */
export async function signInForGraph(): Promise<{ token: string; account: string }> {
  const msal = await getApp();
  const result = await msal.loginPopup({ scopes: GRAPH_SCOPES, prompt: "select_account" });
  return { token: result.accessToken, account: result.account?.username ?? "" };
}

/** MSAL v5 error messages are only links, so turn the common error codes into plain words. */
export function describeSignInError(e: unknown): string | null {
  const code = (e as { errorCode?: string } | null)?.errorCode;
  if (!code) return null;
  const text = `${code} ${(e as { subError?: string }).subError ?? ""} ${(e as { errorMessage?: string }).errorMessage ?? ""}`;
  if (code === "user_cancelled") return "Sign-in was cancelled.";
  if (code === "popup_window_error" || code === "empty_window_error") return "The sign-in window was blocked. Allow pop-ups for this site, then try again.";
  if (code === "interaction_in_progress") return "A sign-in window is already open. Finish or close it, then try again.";
  if (code === "timed_out" || code === "monitor_window_timeout") return "Sign-in timed out. Check that this site's /redirect page is registered as a redirect URI in Entra, then try again.";
  if (/AADSTS50011/.test(text)) return `The redirect URI is not registered. Add ${redirectUri()} to the app registration (platform: Single-page application).`;
  if (/AADSTS65001|AADSTS90094|consent/i.test(text)) return "An admin of this Microsoft 365 tenant must approve the app's read-only permissions first.";
  return `Microsoft sign-in failed (${code}). See https://aka.ms/msal.js.errors#${code}`;
}
