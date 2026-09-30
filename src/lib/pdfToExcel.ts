import * as XLSX from "xlsx";
import { PDFParse } from "pdf-parse";

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

/** Heuristic: detect tabular content from raw extracted text when no vector tables exist. */
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

export async function extractPdfData(
  buffer: Buffer,
  fileName: string
): Promise<ExtractionResult> {
  const parser = new PDFParse({ data: new Uint8Array(buffer) });
  try {
    // ---- Metadata ----
    const meta: { label: string; value: string }[] = [
      { label: "File Name", value: fileName },
    ];
    let pageCount = 0;
    try {
      const info = await parser.getInfo();
      pageCount = info.total;
      meta.push({ label: "Pages", value: String(info.total) });
      const raw = (info.info || {}) as Record<string, unknown>;
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

    // ---- Text per page ----
    const textResult = await parser.getText();
    const pageRows: (string | number)[][] = [["Page", "Text"]];
    for (const page of textResult.pages) {
      const clean = (page.text || "").trim();
      if (clean) pageRows.push([page.num, clean]);
    }

    // ---- Tables: prefer pdf-parse vector-based extraction, fall back to heuristic ----
    let tableRows: (string | number)[][] = [];
    try {
      const tableResult = await parser.getTable();
      const tables = (tableResult.mergedTables || []).filter((t) => t.length > 0);
      if (tables.length > 0) {
        tableRows = tables.reduce(
          (acc, table) => [...acc, ...table.map((row) => row.map(coerceCell))],
          [] as (string | number)[][]
        );
      }
    } catch {
      // vector table detection can fail on unusual PDFs; heuristic fallback below
    }
    if (tableRows.length === 0) {
      tableRows = detectTableHeuristic(textResult.text).map((row) =>
        row.map(coerceCell)
      );
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
    await parser.destroy();
  }
}

export function buildExcelWorkbook(result: ExtractionResult): Buffer {
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
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" }) as Buffer;
}

export function getOutputFileName(fileName: string): string {
  return fileName.replace(/\.pdf$/i, "") + ".xlsx";
}
