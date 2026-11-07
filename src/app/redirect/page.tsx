"use client";

/**
 * MSAL v5 redirect bridge. Microsoft's sign-in page sends the result here, and this page
 * passes it back to the main window (BroadcastChannel). It must NOT be served with a
 * Cross-Origin-Opener-Policy header (see vercel.json).
 */
import { useEffect } from "react";
import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";

export default function Redirect() {
  useEffect(() => {
    broadcastResponseToMainFrame().catch((error: unknown) => {
      console.error("Could not pass the sign-in result back to the app:", error);
    });
  }, []);
  return <p style={{ fontFamily: "system-ui", padding: 24, color: "#555" }}>Signing in… this window closes by itself.</p>;
}
