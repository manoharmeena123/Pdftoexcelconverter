// Generates a simple text-based test PDF for verification.
import { writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const lines = [
  "Employee Report 2026",
  "Name     Dept      Salary",
  "Alice    Sales     50000",
  "Bob      Engg      65000",
  "Charlie  HR        48000",
  "Total employees: 3",
];

const stream = lines
  .map((l, i) => `BT /F1 12 Tf 50 ${750 - i * 25} Td (${l.replace(/[()\\]/g, "")}) Tj ET`)
  .join("\n");

const objects = [
  "<< /Type /Catalog /Pages 2 0 R >>",
  "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
  "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
  `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
  "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
];

let pdf = "%PDF-1.4\n";
const offsets = [];
objects.forEach((obj, i) => {
  offsets.push(pdf.length);
  pdf += `${i + 1} 0 obj\n${obj}\nendobj\n`;
});
const xrefStart = pdf.length;
pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
offsets.forEach((o) => (pdf += String(o).padStart(10, "0") + " 00000 n \n"));
pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`;

const outPath = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "test-sample.pdf"
);
writeFileSync(outPath, pdf, "latin1");
console.log(`Wrote ${outPath} (${Buffer.byteLength(pdf, "latin1")} bytes)`);
