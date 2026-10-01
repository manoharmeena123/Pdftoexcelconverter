"use client";

import { useCallback, useRef, useState } from "react";
import {
  extractPdfData,
  buildExcelWorkbook,
  getOutputFileName,
  type SheetData,
} from "@/lib/pdfToExcel";

interface ConversionResult {
  fileName: string;
  outputName: string;
  meta: { label: string; value: string }[];
  sheets: SheetData[];
  totalRows: number;
  excel: Uint8Array<ArrayBuffer>;
}

const MAX_PREVIEW_ROWS = 50;

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<ConversionResult | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = (f: File | null) => {
    setError(null);
    setResult(null);
    if (f && f.type !== "application/pdf" && !f.name.toLowerCase().endsWith(".pdf")) {
      setError("Please select a PDF file.");
      return;
    }
    setFile(f);
  };

  const convert = useCallback(async () => {
    if (!file) return;
    setLoading(true);
    setError(null);
    setResult(null);
    try {
      // Conversion runs entirely in the browser: no upload, no size limits,
      // and the file never leaves this device.
      const data = new Uint8Array(await file.arrayBuffer());
      const extraction = await extractPdfData(data, file.name);
      const excel = buildExcelWorkbook(extraction);
      setResult({
        fileName: file.name,
        outputName: getOutputFileName(file.name),
        meta: extraction.meta,
        sheets: extraction.sheets,
        totalRows: extraction.totalRows,
        excel,
      });
      setActiveSheet(0);
    } catch (err) {
      console.error("PDF conversion error:", err);
      setError(
        err instanceof Error && /password|encrypt/i.test(err.message)
          ? "The PDF is password-protected or encrypted. Please upload an unlocked PDF."
          : "Failed to process the PDF file. Please make sure it is a valid, non-encrypted PDF."
      );
    } finally {
      setLoading(false);
    }
  }, [file]);

  const downloadExcel = () => {
    if (!result) return;
    const blob = new Blob([result.excel], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = result.outputName;
    a.click();
    URL.revokeObjectURL(url);
  };

  const sheet = result?.sheets[activeSheet];

  return (
    <main className="min-h-screen bg-zinc-50 font-sans dark:bg-black">
      <div className="mx-auto max-w-4xl px-6 py-14">
        {/* Header */}
        <div className="mb-10 text-center">
          <div className="mb-3 inline-flex h-14 w-14 items-center justify-center rounded-2xl bg-emerald-600 text-3xl shadow-lg">
            📄
          </div>
          <h1 className="text-3xl font-bold tracking-tight text-zinc-900 dark:text-zinc-50">
            PDF to Excel Converter
          </h1>
          <p className="mt-2 text-zinc-600 dark:text-zinc-400">
            Upload a PDF and extract its details, text, and tabular data into an Excel workbook.
          </p>
        </div>

        {/* Drop zone */}
        <div
          className={`rounded-2xl border-2 border-dashed p-10 text-center transition-colors ${
            dragOver
              ? "border-emerald-500 bg-emerald-50 dark:bg-emerald-950/30"
              : "border-zinc-300 bg-white dark:border-zinc-700 dark:bg-zinc-900"
          }`}
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragOver(false);
            handleFile(e.dataTransfer.files?.[0] ?? null);
          }}
        >
          <input
            ref={inputRef}
            type="file"
            accept="application/pdf,.pdf"
            className="hidden"
            onChange={(e) => handleFile(e.target.files?.[0] ?? null)}
          />
          <p className="mb-4 text-lg text-zinc-500 dark:text-zinc-400">
            Drag &amp; drop your PDF here, or
          </p>
          <button
            onClick={() => inputRef.current?.click()}
            className="rounded-full bg-zinc-900 px-6 py-2.5 font-medium text-white transition-colors hover:bg-zinc-700 dark:bg-zinc-100 dark:text-black dark:hover:bg-zinc-300"
          >
            Browse Files
          </button>
          {file && (
            <div className="mt-5 flex items-center justify-center gap-3 rounded-xl bg-zinc-100 px-4 py-3 dark:bg-zinc-800">
              <span className="text-sm font-medium text-zinc-800 dark:text-zinc-200">
                {file.name}
              </span>
              <span className="text-xs text-zinc-500">
                {(file.size / 1024).toFixed(1)} KB
              </span>
              <button
                onClick={() => handleFile(null)}
                className="text-xs font-medium text-red-500 hover:underline"
              >
                Remove
              </button>
            </div>
          )}
        </div>

        {/* Convert button */}
        <button
          onClick={convert}
          disabled={!file || loading}
          className="mt-6 w-full rounded-xl bg-emerald-600 py-3.5 text-lg font-semibold text-white shadow transition-all hover:bg-emerald-500 disabled:cursor-not-allowed disabled:opacity-40"
        >
          {loading ? "Converting…" : "Convert to Excel"}
        </button>

        {/* Error */}
        {error && (
          <div className="mt-6 rounded-xl border border-red-200 bg-red-50 px-5 py-4 text-red-700 dark:border-red-900 dark:bg-red-950/40 dark:text-red-400">
            ⚠️ {error}
          </div>
        )}

        {/* Results */}
        {result && (
          <section className="mt-10">
            <div className="mb-5 flex flex-col items-start justify-between gap-4 rounded-2xl bg-white p-6 shadow dark:bg-zinc-900 sm:flex-row sm:items-center">
              <div>
                <h2 className="text-xl font-semibold text-zinc-900 dark:text-zinc-50">
                  ✅ Conversion complete
                </h2>
                <p className="mt-1 text-sm text-zinc-500">
                  {result.meta.find((m) => m.label === "Pages")?.value ?? "?"} page(s) ·{" "}
                  {result.sheets.length} sheet(s) · {result.totalRows} rows extracted
                </p>
              </div>
              <button
                onClick={downloadExcel}
                className="rounded-xl bg-emerald-600 px-6 py-3 font-semibold text-white shadow transition-colors hover:bg-emerald-500"
              >
                ⬇️ Download {result.outputName}
              </button>
            </div>

            {/* Sheet tabs */}
            <div className="mb-3 flex flex-wrap gap-2">
              {result.sheets.map((s, i) => (
                <button
                  key={s.name}
                  onClick={() => setActiveSheet(i)}
                  className={`rounded-full px-4 py-1.5 text-sm font-medium transition-colors ${
                    i === activeSheet
                      ? "bg-zinc-900 text-white dark:bg-zinc-100 dark:text-black"
                      : "bg-white text-zinc-600 shadow-sm hover:bg-zinc-100 dark:bg-zinc-900 dark:text-zinc-400 dark:hover:bg-zinc-800"
                  }`}
                >
                  {s.name}
                </button>
              ))}
            </div>

            {/* Preview table */}
            {sheet && (
              <div className="overflow-hidden rounded-2xl border border-zinc-200 bg-white shadow dark:border-zinc-800 dark:bg-zinc-900">
                <div className="max-h-105 overflow-auto">
                  <table className="w-full border-collapse text-left text-sm">
                    <thead className="sticky top-0 bg-zinc-100 dark:bg-zinc-800">
                      <tr>
                        {sheet.rows[0]?.map((cell, c) => (
                          <th
                            key={c}
                            className="whitespace-nowrap border-b border-zinc-200 px-4 py-2.5 font-semibold text-zinc-700 dark:border-zinc-700 dark:text-zinc-300"
                          >
                            {String(cell)}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {sheet.rows.slice(1, MAX_PREVIEW_ROWS + 1).map((row, r) => (
                        <tr
                          key={r}
                          className="odd:bg-white even:bg-zinc-50 dark:odd:bg-zinc-900 dark:even:bg-zinc-950/50"
                        >
                          {row.map((cell, c) => (
                            <td
                              key={c}
                              className="max-w-105 truncate border-b border-zinc-100 px-4 py-2 text-zinc-600 dark:border-zinc-800 dark:text-zinc-400"
                              title={String(cell)}
                            >
                              {String(cell)}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {sheet.rows.length - 1 > MAX_PREVIEW_ROWS && (
                  <p className="border-t border-zinc-200 px-4 py-2.5 text-xs text-zinc-500 dark:border-zinc-800">
                    Showing first {MAX_PREVIEW_ROWS} of {sheet.rows.length - 1} rows — the
                    downloaded Excel file contains all of them.
                  </p>
                )}
              </div>
            )}
          </section>
        )}

        <footer className="mt-14 text-center text-xs text-zinc-400 dark:text-zinc-600">
          Conversion happens in your browser — files are never uploaded to a server, and there is
          no size limit. Works best with text-based PDFs; scanned/image PDFs require OCR and are
          not supported.
        </footer>
      </div>
    </main>
  );
}
