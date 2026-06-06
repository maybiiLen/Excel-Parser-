import * as XLSX from "xlsx";
import type { Grid, Row } from "./types";

/**
 * Parse clipboard contents into a Grid (array of row-arrays).
 *
 * Excel and Google Sheets place an HTML `<table>` on the clipboard under the
 * `text/html` flavor; we prefer that. When only plain text is available we fall
 * back to `text/plain`, which SheetJS sniffs as TSV/CSV.
 *
 * Returns an empty Grid when the clipboard holds nothing parseable.
 */
export function parseClipboard(data: DataTransfer): Grid {
  const html = data.getData("text/html");
  if (html && html.trim()) return parseString(html);

  const text = data.getData("text/plain");
  if (text && text.trim()) return parseString(text);

  return [];
}

function parseString(input: string): Grid {
  const workbook = XLSX.read(input, { type: "string" });
  const sheetName = workbook.SheetNames[0];
  const sheet = sheetName ? workbook.Sheets[sheetName] : undefined;
  if (!sheet) return [];

  // header: 1   -> rows as arrays of cells (no header inference)
  // blankrows   -> drop fully empty rows
  // defval: ""  -> keep blank cells as "" so column positions stay aligned
  // raw: false  -> return formatted display strings
  return XLSX.utils.sheet_to_json<Row>(sheet, {
    header: 1,
    blankrows: false,
    defval: "",
    raw: false,
  });
}
