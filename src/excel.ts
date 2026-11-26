// ===== 填充引擎：先合并空白单元格 -> 居中填入 -> 插入签名图片 -> 导出 =====
// 只做三件事，不做标签猜测、不做占位符扫描：格子位置全部由 mapping.ts 指定。
import type ExcelJSNamespace from "exceljs";
import { resolveField } from "./sdk";
import type { BitableRecord, FieldMeta } from "./sdk";
import type { CellField, BlockField, CheckField, ImageField, EditableField, SheetDef } from "./mapping";
import { fmtVerdict } from "./mapping";
import {
  addrToRowCol,
  bytesToDataUri,
  cellText as cellTextValue,
  getColWidthPx,
  getRowHeightPx,
  parseRange,
} from "./excel-utils";
import type { GridRange } from "./excel-utils";

const ExcelJS = require("exceljs") as typeof ExcelJSNamespace;
const DATE_FMT = "YYYY-MM-DD";

/** 签名图来源：优先给二进制（可写进 xlsx）；拿不到时给 URL，由网页预览直接显示 */
export type ImageProvider = (fields: string[], im?: ImageField) => Promise<{ bytes?: Uint8Array; url?: string } | null>;

// ---------- 基础 ----------

/** 加载工作簿（模板 url 或 Blob） */
export async function loadWorkbook(src: string | Blob): Promise<ExcelJSNamespace.Workbook> {
  const buf = src instanceof Blob ? await src.arrayBuffer() : await (await fetch(src)).arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf as any);
  return wb;
}

/** 按一张映射表填充工作簿（行高、列宽、页面设置一律不改 —— 表单长度与源文件一致） */
export async function fillSheet(
  wb: ExcelJSNamespace.Workbook,
  sheet: SheetDef,
  record: BitableRecord,
  fieldsByName: Map<string, FieldMeta>,
  getImage?: ImageProvider,
  /** 界面输入框的值（EditableField.key -> 文本），优先于表格字段 */
  overrides?: Record<string, string>
): Promise<void> {
  const ws = wb.worksheets[0];
  for (const m of sheet.merges ?? []) {
    const rg = parseRange(m);
    if (rg) mergeBlock(ws, rg); // 仅合并，不写值
  }
  for (const b of sheet.blocks) fillBlock(ws, b, record, fieldsByName);
  for (const c of sheet.cells) fillCell(ws, c, record, fieldsByName);
  for (const ed of sheet.editable ?? []) fillEditable(ws, ed, record, fieldsByName, overrides);
  for (const ck of sheet.checks ?? []) fillCheck(ws, ck, record, fieldsByName);
  if (getImage) {
    for (const im of sheet.images) await fillImage(wb, ws, im, getImage);
  }
}

// ---------- 字段、表头与复选框 ----------

/** 合并区：先拆掉内部小合并并合成一整块，再整体居中写入 */
function fillBlock(ws: ExcelJSNamespace.Worksheet, b: BlockField, record: BitableRecord, fieldsByName: Map<string, FieldMeta>): void {
  const rg = parseRange(b.range);
  if (!rg) return;
  const master = mergeBlock(ws, rg);
  const raw = fieldText(record, b.fields, fieldsByName);
  master.value = raw ? (b.fmt ? b.fmt(raw) : raw) : "";
  master.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  if (b.fit && master.value) autoFit(ws, rg, cellTextValue(master), Number((master.font as any)?.size) || 12, true);
}

/** 单元格：居中写入；append 模式保留模板原文字（如「申请日期：」后接日期） */
function fillCell(ws: ExcelJSNamespace.Worksheet, c: CellField, record: BitableRecord, fieldsByName: Map<string, FieldMeta>): void {
  // 条件写入：同一字段对应两个分支格子（如「关连整改通知书 No.」有合格/不合格两行）
  if (c.whenField && c.whenField.length) {
    const cond = fmtVerdict(fieldText(record, c.whenField, fieldsByName)).trim();
    const hit = (c.whenValues ?? []).some((v) => cond === v);
    if (!hit) return;
  }
  const raw = c.dateFmt
    ? resolveField(record, c.fields, fieldsByName, c.dateFmt).text
    : fieldText(record, c.fields, fieldsByName);
  const cell = ws.getCell(c.cell);
  /** 写入后按本格 / 真实合并区宽度自适应缩字号，不换行、不改行高 */
  const fit = (text: string) => {
    if (!c.fit || !text) return;
    const rg = cellRange(ws, c.cell);
    if (rg) autoFit(ws, rg, text, Number((cell.font as any)?.size) || 12, !!c.wrap);
  };
  // 单段/多段填空都复用同一个居中替换函数，避免两套规则逐渐不一致
  if (c.insert && raw) {
    const text = replaceBlank(cellTextValue(cell), c.insert[0], c.insert[1], raw);
    if (text !== null) {
      cell.value = text;
      fit(text);
      return;
    }
  }
  if (c.inserts && raw) {
    const parts = c.inserts.fmt(raw);
    const text = c.inserts.slots.reduce(
      (out, [prefix, suffix], index) => replaceBlank(out, prefix, suffix, parts[index] ?? "") ?? out,
      cellTextValue(cell)
    );
    cell.value = text;
    fit(text);
    return;
  }
  if (c.append) {
    if (!raw) return;
    cell.value = cellTextValue(cell) + (c.fmt ? c.fmt(raw) : raw);
    fit(cellTextValue(cell));
    return;
  }
  // 带 fmt 时即便原文为空也要走一遍（fmt 可能从原文里挑一部分，如第 n 条附条件）
  const val = c.fmt ? c.fmt(raw) : raw;
  if (!val) return;
  cell.value = val;
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: c.align ?? "center",
    vertical: "middle",
    wrapText: !!c.wrap,
  };
  fit(val);
}

/**
 * 表头可编辑文本（单据编号/申请单编号等）：整格重写为「标签 + 值」，模板里写死的旧编号随之被覆盖。
 * 取值优先级：界面输入框 -> 表格字段 -> 空（只剩标签）。
 * fit=true 时，文本超出单元格宽度会自动缩字号，保证不溢出边框。
 */
function fillEditable(
  ws: ExcelJSNamespace.Worksheet,
  ed: EditableField,
  record: BitableRecord,
  fieldsByName: Map<string, FieldMeta>,
  overrides?: Record<string, string>
): void {
  const typed = (overrides?.[ed.key] ?? "").trim();
  const val = typed || (ed.fields ? fieldText(record, ed.fields, fieldsByName) : "");
  const text = `${ed.prefix ?? ""}${val}`;
  const cell = ws.getCell(ed.cell);
  cell.value = text;
  if (ed.fit && text) {
    const range = cellRange(ws, ed.cell);
    const base = Number((cell.font as any)?.size) || 10;
    if (range) autoFit(ws, range, text, base, false);
  }
}

/** 中文按 1 字宽、西文按 0.52 字宽估算（宋体近似）。 */
function charWidthUnits(char: string): number {
  return /[\u2e80-\u9fff\u3000-\u303f\uff00-\uffef]/.test(char) ? 1 : 0.52;
}

/** 估算 96dpi 下的文本宽度。 */
function textWidthPx(text: string, sizePt: number): number {
  const em = (sizePt * 96) / 72;
  let width = 0;
  for (const char of text) width += charWidthUnits(char);
  return width * em;
}

/** 找出覆盖该单元格的合并区（返回列范围），没合并返回 null */
function findMerge(ws: ExcelJSNamespace.Worksheet, r0: number, c0: number): { c0: number; c1: number } | null {
  for (const m of ((ws.model as any).merges || []) as string[]) {
    const rg = parseRange(m);
    if (rg && rg.r0 === r0 && rg.c0 === c0) return { c0: rg.c0, c1: rg.c1 };
  }
  return null;
}

/**
 * 填入文字的自适应：先估算按照当前字号需要几行，如果行高放得下就只换行；
 * 放不下再按 0.5pt 逐档缩字号，直到「行数 × 行高 ≤ 区域高度」或到 6pt 下限。
 * 只改字号与 wrapText，不动列宽 / 行高 / 页面设置 —— 表单几何与源文件一致。
 */
function autoFit(
  ws: ExcelJSNamespace.Worksheet,
  rg: { r0: number; c0: number; r1: number; c1: number },
  text: string,
  baseSize: number,
  canWrap: boolean
): void {
  if (!text) return;
  const cell = ws.getCell(rg.r0, rg.c0);
  let boxW = 0;
  for (let c = rg.c0; c <= rg.c1; c++) boxW += getColWidthPx(ws, c);
  let boxH = 0;
  for (let r = rg.r0; r <= rg.r1; r++) boxH += getRowHeightPx(ws, r);
  if (!boxW || !boxH) return;
  const availW = boxW - 4;      // 两侧各留 2px，别顶到边框
  const availH = boxH - 2;
  const lines = (s: number) => {
    const em = (s * 96) / 72;
    const perLine = Math.max(1, Math.floor(availW / em));
    let total = 0;
    for (const part of String(text).split("\n")) {
      // 中文按 1 字、西文按半字估算每行容量
      let units = 0;
      for (const char of part) units += charWidthUnits(char);
      total += Math.max(1, Math.ceil(units / perLine));
    }
    return total;
  };
  let size = baseSize;
  for (; size >= 6; size -= 0.5) {
    const em = (size * 96) / 72;
    if (canWrap) {
      if (lines(size) * em * 1.15 <= availH) break;
    } else if (textWidthPx(text, size) <= availW) {
      break;
    }
  }
  size = Math.max(6, size);
  setFontSize(cell, size);
  if (canWrap) cell.alignment = { ...(cell.alignment || {}), wrapText: true, vertical: "middle" };
}

/**
 * 只改自己这一格的字号。ExcelJS 里同 style id 的格共享同一个 style 对象，
 * 直接改 cell.font 会污染其他格（如 C33 缩字号连带 F33「签名：」也变小），
 * 所以整格赋一个全新的 style 对象。
 */
function setFontSize(cell: ExcelJSNamespace.Cell, size: number): void {
  cell.style = { ...(cell.style as any), font: { ...(cell.font as any), size } };
}

/** 单元格的合并范围（用于 autoFit）；自身单格时返回单格范围 */
function cellRange(ws: ExcelJSNamespace.Worksheet, addr: string) {
  const rc = addrToRowCol(addr);
  if (!rc) return null;
  const merged = findMerge(ws, rc.row, rc.col);
  return { r0: rc.row, c0: rc.col, r1: rc.row, c1: merged ? merged.c1 : rc.col };
}

/** 正则转义（模板文字里的「第」「回」等直接用） */
function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** 把前后缀之间的空白替换为值，并保持原有总宽度；找不到填空位置时返回 null。 */
function replaceBlank(text: string, prefix: string, suffix: string, value: string): string | null {
  if (!value) return null;
  const pattern = new RegExp(`${escapeRe(prefix)}(\\s*)${escapeRe(suffix)}`);
  const match = text.match(pattern);
  if (!match) return null;
  const total = match[1].length;
  const left = Math.max(0, Math.ceil((total - value.length) / 2));
  const right = Math.max(0, total - value.length - left);
  return text.replace(pattern, `${prefix}${" ".repeat(left)}${value}${" ".repeat(right)}${suffix}`);
}

/** 复选框：值等于 when 时，把该格第 box 个 □ 涂成实心黑方块 */
function fillCheck(ws: ExcelJSNamespace.Worksheet, ck: CheckField, record: BitableRecord, fieldsByName: Map<string, FieldMeta>): void {
  const raw = fieldText(record, ck.fields, fieldsByName);
  if (!raw) return;
  const val = (ck.fmt ? ck.fmt(raw) : raw).trim();
  if (val !== ck.when) return;
  const cell = ws.getCell(ck.cell);
  const master = (cell as any).master || cell;
  // 模板里画好的方框（无 □ 字符）：整格填黑，边框保持原样
  if (ck.fill) {
    master.style = {
      ...(master.style as any),
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } },
    };
    return;
  }
  const text = cellTextValue(master);
  const next = tickBox(text, ck.box ?? 0);
  if (next !== text) master.value = next;
}

/** 把第 n 个 □ 替换成 ■（实心方块 = 涂黑整个方框） */
function tickBox(text: string, box: number): string {
  let i = -1;
  return text.replace(/□/g, (m) => (++i === box ? "■" : m));
}

/** 签名：把附件图片缩放到格子里插入（不是写文字 / 不是复制模板里的图） */
async function fillImage(
  wb: ExcelJSNamespace.Workbook,
  ws: ExcelJSNamespace.Worksheet,
  im: ImageField,
  getImage: ImageProvider
): Promise<void> {
  const img = await getImage(im.fields, im);
  if (!img) return;
  // 按映射里写的格子定位（不取合并主格：F37:G37 合并后签名要压在右侧 G37 上）
  const rc = addrToRowCol(ws.getCell(im.cell).address);
  if (!rc) return;

  // 拿不到二进制（跨域/鉴权失败）时保留 URL，交给网页预览用 <img> 直接显示，
  // 至少打印时能看见签名；xlsx 导出仍需要二进制。
  if (!img.bytes || !img.bytes.length) {
    if (img.url) {
      const list = (((wb as any).__urlImages as any[]) ||= []);
      list.push({ cell: ws.getCell(im.cell).address, url: img.url, offsetX: im.offsetX ?? 0 });
    }
    return;
  }

  const { w: iw, h: ih } = imageSize(img.bytes);
  const boxW = im.maxW ?? getColWidthPx(ws, rc.col);
  const boxH = im.maxH ?? getRowHeightPx(ws, rc.row);
  const pad = 2;
  const scale = Math.min((boxW - pad * 2) / iw, (boxH - pad * 2) / ih, 1);
  const w = Math.max(8, Math.round(iw * scale));
  const h = Math.max(6, Math.round(ih * scale));

  const ext = detectImageExt(img.bytes);
  const id = wb.addImage({ base64: bytesToDataUri(img.bytes, ext), extension: ext });
  // align:"right" 推到格子右端；offsetX 指定距左边界的间距（标签右侧放图时用）；1px = 9525 EMU
  const offPx = im.align === "right" ? Math.max(0, boxW - w - pad) : im.offsetX ? Math.max(0, im.offsetX) : 0;
  const offEmu = Math.round(offPx * 9525);
  const tl = offEmu
    ? { nativeCol: rc.col - 1, nativeColOff: offEmu, nativeRow: rc.row - 1, nativeRowOff: 0 }
    : { col: rc.col - 1, row: rc.row - 1 };
  ws.addImage(id, { tl, ext: { width: w, height: h }, editAs: "oneCell" } as any);
}

// ---------- 合并单元格 ----------

/**
 * 合并成整块：先拆内部已有小合并，再合并。
 * exceljs 合并时会把覆盖格样式改成主格样式，整块的外框（右/下边）会丢，
 * 所以合并前把区域四边的边框快照下来，合并后逐格还原——打印出来与源文件一致。
 */
function mergeBlock(ws: ExcelJSNamespace.Worksheet, rg: GridRange): ExcelJSNamespace.Cell {
  const edges: Array<{ addr: string; border: any }> = [];
  const snap = (r: number, c: number) => {
    const cell = ws.getCell(r, c);
    edges.push({ addr: cell.address, border: clone(cell.border) });
  };
  for (let c = rg.c0; c <= rg.c1; c++) {
    snap(rg.r0, c);
    if (rg.r1 !== rg.r0) snap(rg.r1, c);
  }
  for (let r = rg.r0; r <= rg.r1; r++) {
    snap(r, rg.c0);
    if (rg.c1 !== rg.c0) snap(r, rg.c1);
  }

  for (const m of [...((ws.model as any).merges || [])]) {
    const inner = parseRange(m);
    if (inner && inner.r0 >= rg.r0 && inner.r1 <= rg.r1 && inner.c0 >= rg.c0 && inner.c1 <= rg.c1) {
      ws.unMergeCells(m);
    }
  }
  if (rg.r0 !== rg.r1 || rg.c0 !== rg.c1) ws.mergeCells(rg.r0, rg.c0, rg.r1, rg.c1);

  // 用「新建样式对象」而不是改 cell.border：合并后覆盖格与主格共享同一个 style 对象，
  // 直接改 border 会把改动串到同一样式编号的其他格子（实测会污染 F33/F41）。
  for (const e of edges) {
    const cell = ws.getCell(e.addr);
    cell.style = { ...(cell.style as any), border: clone(e.border) };
  }
  return ws.getCell(rg.r0, rg.c0);
}

/** 深拷贝（避免多个格子共用同一个边框对象导致写盘时样式串味） */
function clone<T>(v: T): T {
  return v == null ? v : JSON.parse(JSON.stringify(v));
}

// ---------- 工具 ----------

function fieldText(record: BitableRecord, names: string[], fieldsByName: Map<string, FieldMeta>): string {
  return resolveField(record, names, fieldsByName, DATE_FMT).text;
}

function detectImageExt(bytes: Uint8Array): "png" | "jpeg" | "gif" {
  if (bytes[0] === 0x89 && bytes[1] === 0x50) return "png";
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return "jpeg";
  if (bytes[0] === 0x47 && bytes[1] === 0x49) return "gif";
  return "png";
}

/** 读图片原始尺寸（PNG 的 IHDR / JPEG 的 SOF 段），解析不出来时给个保守默认值 */
function imageSize(bytes: Uint8Array): { w: number; h: number } {
  const ext = detectImageExt(bytes);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (ext === "png" && bytes.length > 24) {
    return { w: view.getUint32(16), h: view.getUint32(20) };
  }
  if (ext === "jpeg") {
    let i = 2;
    while (i < bytes.length - 9) {
      if (bytes[i] !== 0xff) { i++; continue; }
      const marker = bytes[i + 1];
      const len = view.getUint16(i + 2);
      // SOF0..SOF3 / SOF5..SOF7 / SOF9..SOF11 ... 含尺寸
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { h: view.getUint16(i + 5), w: view.getUint16(i + 7) };
      }
      i += 2 + len;
    }
  }
  return { w: 120, h: 40 };
}
