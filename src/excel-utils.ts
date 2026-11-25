// Excel 填充与网页预览共用的基础函数，避免两套几何/文本规则逐渐漂移。
import type ExcelJS from "exceljs";

export interface GridRange {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export function colNum(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}

export function addrToRowCol(addr: string): { row: number; col: number } | null {
  const match = /^([A-Z]{1,3})(\d+)$/.exec((addr || "").trim().toUpperCase());
  return match ? { row: Number(match[2]), col: colNum(match[1]) } : null;
}

/** "C10:D15" / "C10" -> 行列范围。 */
export function parseRange(value: string): GridRange | null {
  const match = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec((value || "").trim().toUpperCase());
  if (!match) return null;
  const r0 = Number(match[2]);
  const c0 = colNum(match[1]);
  const r1 = match[4] ? Number(match[4]) : r0;
  const c1 = match[3] ? colNum(match[3]) : c0;
  return {
    r0: Math.min(r0, r1),
    c0: Math.min(c0, c1),
    r1: Math.max(r0, r1),
    c1: Math.max(c0, c1),
  };
}

/** 单元格显示文本：富文本、公式结果、日期和普通值统一转字符串。 */
export function cellText(cell: ExcelJS.Cell): string {
  const value: any = cell.value;
  if (value == null) return "";
  if (typeof value === "object") {
    if (value.richText) return value.richText.map((part: any) => part.text).join("");
    if ("text" in value) return String(value.text ?? "");
    if ("result" in value) return String(value.result ?? "");
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    if (value.error) return "";
    return "";
  }
  return String(value);
}

/** Excel 列宽（字符）-> 96dpi 像素。 */
export function colWidthPx(width: number): number {
  return Math.round(width * 7) + 5;
}

/** Excel 行高（pt）-> 96dpi 像素。 */
export function rowHeightPx(height: number): number {
  return Math.round(height * (96 / 72));
}

export function getColWidthPx(ws: ExcelJS.Worksheet, col: number): number {
  return colWidthPx(ws.getColumn(col).width || 8.43);
}

export function getRowHeightPx(ws: ExcelJS.Worksheet, row: number): number {
  return rowHeightPx(ws.getRow(row).height || 15);
}

/** Uint8Array -> data URI；分块转换，避免大图片一次性展开参数导致栈溢出。 */
export function bytesToDataUri(bytes: Uint8Array, ext: string): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunkSize)) as any);
  }
  return `data:image/${ext};base64,${btoa(binary)}`;
}
