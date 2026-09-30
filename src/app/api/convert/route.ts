import { NextRequest, NextResponse } from "next/server";
import { extractPdfData, buildExcelWorkbook, getOutputFileName } from "@/lib/pdfToExcel";

export const runtime = "nodejs";

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return NextResponse.json(
        { error: "No PDF file provided. Please attach a file under the 'file' field." },
        { status: 400 }
      );
    }

    if (file.type && file.type !== "application/pdf") {
      return NextResponse.json(
        { error: "Invalid file type. Only PDF files are supported." },
        { status: 400 }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const result = await extractPdfData(buffer, file.name);
    const excelBuffer = buildExcelWorkbook(result);

    return NextResponse.json({
      fileName: file.name,
      outputName: getOutputFileName(file.name),
      meta: result.meta,
      sheets: result.sheets,
      totalRows: result.totalRows,
      excelBase64: excelBuffer.toString("base64"),
    });
  } catch (err) {
    console.error("PDF conversion error:", err);
    const message =
      err instanceof Error && err.message.includes("Invalid PDF")
        ? "The uploaded file is not a valid or is a corrupted PDF."
        : "Failed to process the PDF file. Please make sure it is a valid, non-encrypted PDF.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
