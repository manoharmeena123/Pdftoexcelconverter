// Copies the pdf.js worker from node_modules into public/ so it can be served
// as a static asset (works identically in `next dev` and on Vercel).
import { copyFileSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const scriptsDir = path.dirname(fileURLToPath(import.meta.url));
const source = path.join(
  scriptsDir,
  "..",
  "node_modules",
  "pdfjs-dist",
  "build",
  "pdf.worker.min.mjs"
);
const dest = path.join(scriptsDir, "..", "public", "pdf.worker.min.mjs");

try {
  if (!existsSync(source)) {
    console.warn("copy-pdf-worker: pdfjs-dist worker not found, skipping.");
    process.exit(0);
  }
  copyFileSync(source, dest);
  console.log("copy-pdf-worker: copied pdf.worker.min.mjs to public/");
} catch (err) {
  console.error("copy-pdf-worker: failed to copy worker file:", err.message);
  process.exit(0); // don't fail the install over this
}
