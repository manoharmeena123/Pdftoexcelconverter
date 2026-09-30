import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Keep PDF parsing libraries unbundled so pdf.js can load its Node worker.
  serverExternalPackages: ["pdf-parse", "pdfjs-dist"],
};

export default nextConfig;
