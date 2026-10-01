import * as XLSX from "xlsx";

export interface SheetData {
  name: string;
  rows: (string | number)[][];
}

export interface ExtractionResult {
  fileName: string;
  meta: { label: string; value: string }[];
  sheets: SheetData[];
  totalRows: number;
}

/** Split a text line into columns using tabs, 2+ spaces, or commas as delimiters. */
function splitIntoColumns(line: string): string[] {
  if (/\t/.test(line)) return line.split("\t").map((s) => s.trim());
  if (/ {2,}/.test(line)) return line.split(/ {2,}/).map((s) => s.trim());
  // Only treat commas as columns when the line looks like a CSV record
  const commas = (line.match(/,/g) || []).length;
  if (commas >= 1 && commas <= 15 && !/[.!?]$/.test(line)) {
    return line.split(",").map((s) => s.trim());
  }
  return [line.trim()];
}

/** Heuristic: detect tabular content from raw extracted text. */
function detectTableHeuristic(text: string): string[][] {
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0);
  if (lines.length === 0) return [];

  const parsed = lines.map(splitIntoColumns);
  const multiColCount = parsed.filter((cols) => cols.length >= 2).length;
  if (multiColCount < Math.max(3, lines.length * 0.2)) return [];

  const maxCols = Math.max(...parsed.map((cols) => cols.length));
  return parsed
    .filter((cols) => cols.length >= 2 || maxCols <= 2)
    .map((cols) => {
      const row = [...cols];
      while (row.length < maxCols) row.push("");
      return row;
    });
}

/** Convert numeric-looking cell strings to numbers so Excel treats them as numeric. */
function coerceCell(cell: string): string | number {
  const trimmed = (cell || "").trim();
  if (trimmed === "" || !/\d/.test(trimmed)) return trimmed;
  const num = Number(trimmed.replace(/,/g, ""));
  return Number.isFinite(num) ? num : trimmed;
}

interface PositionedItem {
  str: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Rebuild text lines from positioned PDF text items: group items sharing a
 * y-coordinate, sort by x, and insert a tab where the horizontal gap between
 * items is large (indicating a column boundary).
 */
function itemsToLines(items: PositionedItem[]): string[] {
  if (items.length === 0) return [];

  const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);

  // Group items into rows by vertical proximity (half the item height).
  const groups: PositionedItem[][] = [];
  for (const item of sorted) {
    const current = groups[groups.length - 1];
    const anchor = current?.[0];
    const tolerance = Math.max(item.height, anchor?.height ?? 0) * 0.5;
    if (anchor && Math.abs(anchor.y - item.y) <= tolerance) {
      current.push(item);
    } else {
      groups.push([item]);
    }
  }

  return groups.map((group) => {
    group.sort((a, b) => a.x - b.x);
    let line = "";
    let prevEnd = -Infinity;
    for (const item of group) {
      if (prevEnd !== -Infinity) {
        const gap = item.x - prevEnd;
        // A gap wider than ~0.6em means the PDF is visually separating columns.
        if (gap > item.height * 0.6) line += "\t";
        else if (gap > 0) line += " ";
      }
      line += item.str;
      prevEnd = item.x + item.width;
    }
    return line.trim();
  });
}

/** Lazily load pdf.js (browser build) and configure its static worker. */
async function loadPdfjs() {
  const pdfjs = await import("pdfjs-dist");
  if (!pdfjs.GlobalWorkerOptions.workerSrc) {
    pdfjs.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
  }
  return pdfjs;
}

export async function extractPdfData(
  data: Uint8Array,
  fileName: string
): Promise<ExtractionResult> {
  const pdfjs = await loadPdfjs();
  const loadingTask = pdfjs.getDocument({ data });
  const doc = await loadingTask.promise;
  try {
    // ---- Metadata ----
    const meta: { label: string; value: string }[] = [
      { label: "File Name", value: fileName },
    ];
    const pageCount = doc.numPages;
    meta.push({ label: "Pages", value: String(pageCount) });
    try {
      const metadata = await doc.getMetadata();
      const raw = (metadata?.info || {}) as Record<string, unknown>;
      const fields: [string, string][] = [
        ["Title", "Title"],
        ["Author", "Author"],
        ["Subject", "Subject"],
        ["Keywords", "Keywords"],
        ["Creator", "Creator"],
        ["Producer", "Producer"],
        ["CreationDate", "CreationDate"],
        ["ModDate", "ModDate"],
      ];
      for (const [label, key] of fields) {
        const value = raw[key] ?? raw[label.toLowerCase()];
        if (value && String(value).trim()) meta.push({ label, value: String(value) });
      }
    } catch {
      // metadata is optional; keep going
    }

    // ---- Text + tables per page ----
    const pageRows: (string | number)[][] = [["Page", "Text"]];
    const tableRows: (string | number)[][] = [];

    for (let num = 1; num <= doc.numPages; num++) {
      const page = await doc.getPage(num);
      const content = await page.getTextContent();
      const items = (content.items as Array<Record<string, unknown>>)
        .filter((it) => typeof it.str === "string" && (it.str as string).trim() !== "")
        .map((it) => {
          const transform = (it.transform as number[]) || [1, 0, 0, 1, 0, 0];
          return {
            str: it.str as string,
            x: transform[4],
            y: transform[5],
            width: (it.width as number) || 0,
            height: (it.height as number) || 10,
          };
        });

      const lines = itemsToLines(items).filter((l) => l.length > 0);
      const clean = lines.join("\n");
      if (clean) pageRows.push([num, clean]);

      // Position-aware text keeps column gaps as tabs, so the heuristic
      // reliably detects tabular content per page.
      const detected = detectTableHeuristic(clean);
      for (const row of detected) tableRows.push(row.map(coerceCell));
    }

    const sheets: SheetData[] = [
      {
        name: "PDF Details",
        rows: [["Field", "Value"], ...meta.map((m) => [m.label, m.value])],
      },
    ];
    if (tableRows.length > 0) sheets.push({ name: "Table Data", rows: tableRows });
    if (pageRows.length > 1) sheets.push({ name: "Page Text", rows: pageRows });

    const totalRows = sheets.reduce((acc, s) => acc + s.rows.length, 0);
    return { fileName, meta, sheets, totalRows: totalRows || pageCount };
  } finally {
    await loadingTask.destroy();
  }
}

export function buildExcelWorkbook(result: ExtractionResult): Uint8Array<ArrayBuffer> {
  const wb = XLSX.utils.book_new();
  for (const sheet of result.sheets) {
    const ws = XLSX.utils.aoa_to_sheet(sheet.rows);
    // Reasonable column widths
    const colCount = Math.max(...sheet.rows.map((r) => r.length), 1);
    ws["!cols"] = Array.from({ length: colCount }, (_, c) => {
      const maxLen = Math.max(
        ...sheet.rows.map((r) => String(r[c] ?? "").length),
        8
      );
      return { wch: Math.min(maxLen + 2, 80) };
    });
    XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31));
  }
  return new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" }) as ArrayBuffer);
}

export function getOutputFileName(fileName: string): string {
  return fileName.replace(/\.pdf$/i, "") + ".xlsx";
}
