import type { NextConfig } from "next";

/**
 * Two deployment modes:
 *  - Default (Vercel / `next start`): full app including the /api/scan-domain serverless route.
 *  - GITHUB_PAGES=true (`npm run build:pages`): static export. The browser falls back to
 *    DNS-over-HTTPS for SPF/DMARC/DKIM because there is no API route on a static host.
 */
const isPages = process.env.GITHUB_PAGES === "true";
const basePath = process.env.NEXT_PUBLIC_BASE_PATH ?? "";

const nextConfig: NextConfig = {
  poweredByHeader: false,
  env: { NEXT_PUBLIC_TRAILING_SLASH: isPages ? "1" : "" },
  ...(isPages && {
    output: "export",
    basePath,
    trailingSlash: true,
    images: { unoptimized: true },
  }),
};

export default nextConfig;
