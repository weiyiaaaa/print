// ===== 打印预览：按 Excel 的几何尺寸精确还原（列宽/行高/合并/边框/对齐/图片）=====
// 关键点：所有尺寸先在 Excel 坐标系（96dpi 的 px）里算，再统一乘缩放系数；
// 不用 CSS zoom（打印时）、不给 td 加内边距，保证浏览器打印与 Excel 打印一致；
// 页边距由 .print-page 的 padding 还原，@page 保持 0 以去掉浏览器自动页眉页脚。
import type ExcelJSNamespace from "exceljs";
import { addrToRowCol, bytesToDataUri, cellText, colNum, colWidthPx, rowHeightPx } from "./excel-utils";

const EMU_PER_PX = 9525;      // 96dpi 下 EMU -> px
const MM_PER_PX = 25.4 / 96;  // CSS px -> mm
const PX_PER_MM = 96 / 25.4;
const DEFAULT_ROW_H = 15;     // pt
const DEFAULT_COL_W = 8.43;   // 字符宽

interface GridPos { r: number; c: number; rowSpan: number; colSpan: number; }

export interface SheetHtml {
  html: string;
  margins: { l: number; r: number; t: number; b: number };
  /** 表格实际宽度（mm），打印时正好铺满可打印区 */
  pageWmm: number;
  /** 缩放系数（内容超出一页时等比缩小，等于「适应一页」） */
  scale: number;
}

/** 渲染选项 */
export interface SheetHtmlOptions {
  /** 文字位移：单元格地址 -> 偏离字符数（负数向左）。仅打印渲染生效 */
  shift?: Record<string, number>;
}

function esc(s: string): string {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** 收集合并单元格：遍历区域内所有格子（含空格子），按 cell.master 分组 */
function collectMerges(ws: ExcelJSNamespace.Worksheet, maxRow: number, maxCol: number): Map<string, GridPos> {
  const groups = new Map<string, { r0: number; c0: number; r1: number; c1: number }>();
  for (let r = 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    if (!row) continue;
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      if (!cell) continue;
      const master = (cell as any).master || cell;
      if (!master || master.address === cell.address) continue;
      const m = /^([A-Z]+)(\d+)$/.exec(master.address);
      if (!m) continue;
      const mr = Number(m[2]);
      const mc = colNum(m[1]);
      const key = `${mr}_${mc}`;
      const g = groups.get(key) || { r0: mr, c0: mc, r1: mr, c1: mc };
      g.r0 = Math.min(g.r0, mr); g.c0 = Math.min(g.c0, mc);
      g.r1 = Math.max(g.r1, r); g.c1 = Math.max(g.c1, c);
      groups.set(key, g);
    }
  }
  const out = new Map<string, GridPos>();
  for (const [, g] of groups) {
    out.set(`${g.r0}_${g.c0}`, { r: g.r0, c: g.c0, rowSpan: g.r1 - g.r0 + 1, colSpan: g.c1 - g.c0 + 1 });
    for (let r = g.r0; r <= g.r1; r++)
      for (let c = g.c0; c <= g.c1; c++)
        if (!(r === g.r0 && c === g.c0)) out.set(`${r}_${c}`, { r, c, rowSpan: 0, colSpan: 0 });
  }
  return out;
}

function isBlankCell(cell: ExcelJSNamespace.Cell | null): boolean {
  if (!cell) return true;
  const master = (cell as any).master || cell;
  return cellText(master).trim() === "";
}

/**
 * 单元格内容 -> HTML。
 * 富文本按每段自己的字号/字体/粗斜体输出（如标题里「（产品中心）」比主标题小一号），
 * 普通文本走转义；换行统一转 <br>。字号同样乘页面缩放系数 s。
 */
function cellHtml(cell: ExcelJSNamespace.Cell, s: number): string {
  const v: any = cell.value;
  if (v && typeof v === "object" && Array.isArray(v.richText)) {
    return v.richText
      .map((run: any) => {
        const f: any = run.font || {};
        const css: string[] = [];
        if (f.size) css.push(`font-size:${(f.size * s).toFixed(2)}pt`);
        if (f.name) css.push(`font-family:'${f.name}','Microsoft YaHei',sans-serif`);
        if (f.bold) css.push("font-weight:bold");
        if (f.italic) css.push("font-style:italic");
        const color = argbToCss(f.color?.argb);
        if (color && color !== "#ff000000") css.push(`color:${color}`);
        const txt = esc(run.text ?? "").replace(/\n/g, "<br>");
        return css.length ? `<span style="${css.join(";")}">${txt}</span>` : txt;
      })
      .join("");
  }
  return esc(cellText(cell)).replace(/\n/g, "<br>");
}

function argbToCss(argb?: string): string | undefined {
  if (!argb) return undefined;
  if (argb.length === 8) return "#" + argb.slice(2);
  if (argb.length === 6) return "#" + argb;
  return undefined;
}

/** 单边框样式映射（仅渲染模板定义的边，不画假网格线） */
function borderCss(side: any): string {
  if (!side || !side.style) return "";
  const color = argbToCss(side.color?.argb) || "#000";
  const map: Record<string, string> = {
    thin: "1px solid", medium: "2px solid", thick: "3px solid",
    dashed: "1px dashed", dotted: "1px dotted", double: "3px double",
    hair: "1px solid", mediumDashed: "2px dashed", dashDot: "1px dashed",
    mediumDashDot: "2px dashed", dashDotDot: "1px dashed", mediumDashDotDot: "2px dashed",
    slantDashDot: "2px dashed",
  };
  return `${map[side.style] || "1px solid"} ${color}`;
}

function cellCss(cell: ExcelJSNamespace.Cell, s: number, border?: any, canOverflow = false, canOverflowLeft = false): string {
  const st: any = (cell as any).style || {};
  const parts: string[] = ["padding:0"];
  const b = border || st.border || {};
  const t = borderCss(b.top), bo = borderCss(b.bottom), l = borderCss(b.left), r = borderCss(b.right);
  if (t) parts.push(`border-top:${t}`);
  if (bo) parts.push(`border-bottom:${bo}`);
  if (l) parts.push(`border-left:${l}`);
  if (r) parts.push(`border-right:${r}`);
  const a = st.alignment || {};
  // Excel 默认：水平靠左、垂直靠下
  if (a.horizontal === "center") parts.push("text-align:center");
  else if (a.horizontal === "right") parts.push("text-align:right");
  else parts.push("text-align:left");
  if (a.vertical === "top") parts.push("vertical-align:top");
  else if (a.vertical === "middle") parts.push("vertical-align:middle");
  else parts.push("vertical-align:bottom");
  // Excel 的「缩进」：靠右对齐时文字从右边内缩（看起来左移），靠左时从左边内缩
  const indent = Number(a.indent) || 0;
  if (indent > 0) {
    if (a.horizontal === "right") parts.push(`padding-right:${indent}ch`);
    else if (a.horizontal !== "center") parts.push(`padding-left:${indent}ch`);
  }
  // 换行格自动折行并裁掉溢出；不换行格按 Excel 行为外溢（左对齐看右邻是否为空，右/居中看左邻），否则裁切
  if (a.wrapText) parts.push("white-space:pre-wrap;overflow:hidden;word-break:break-all");
  else parts.push(canOverflow || canOverflowLeft ? "white-space:pre;overflow:visible" : "white-space:pre;overflow:hidden");
  parts.push("line-height:1.15");
  const f = st.font || {};
  const size = (f.size || 11) * s;
  parts.push(`font-size:${size.toFixed(2)}pt`);
  // 字体名用单引号，避免把 style="..." 属性提前截断
  parts.push(`font-family:'${f.name || "宋体"}','Microsoft YaHei',sans-serif`);
  if (f.bold) parts.push("font-weight:bold");
  if (f.italic) parts.push("font-style:italic");
  const color = argbToCss(f.color?.argb);
  if (color && color !== "#ff000000") parts.push(`color:${color}`);
  const fill = st.fill;
  if (fill && fill.type === "pattern" && fill.fgColor && fill.pattern !== "none") {
    const fg = argbToCss(fill.fgColor.argb);
    if (fg) parts.push(`background:${fg}`);
  }
  return parts.join(";");
}

/**
 * 合并区边框：Excel 把整块的外框记在四条边上的格子里（右/下边常在"被覆盖"的格子），
 * 这里按区域四边扫描合成，否则合并块会少画右、下两条线。
 */
function mergedBorder(ws: ExcelJSNamespace.Worksheet, r: number, c: number, g: GridPos): any {
  const pick = (arr: any[]) => arr.find((b) => b && b.style);
  const rowSpan = g.rowSpan || 1;
  const colSpan = g.colSpan || 1;
  const top: any[] = [], bottom: any[] = [], left: any[] = [], right: any[] = [];
  for (let i = 0; i < colSpan; i++) {
    top.push((ws.getCell(r, c + i).border || {}).top);
    bottom.push((ws.getCell(r + rowSpan - 1, c + i).border || {}).bottom);
  }
  for (let i = 0; i < rowSpan; i++) {
    left.push((ws.getCell(r + i, c).border || {}).left);
    right.push((ws.getCell(r + i, c + colSpan - 1).border || {}).right);
  }
  return { top: pick(top), bottom: pick(bottom), left: pick(left), right: pick(right) };
}

/** 提取工作表图片（LOGO / 签名）为绝对定位 <img>；按列宽行高还原锚点 */
function imagesHtml(wb: ExcelJSNamespace.Workbook, ws: ExcelJSNamespace.Worksheet, colPx: number[], rowPx: number[], s: number): string {
  let imgs: any[] = [];
  try { imgs = (ws as any).getImages ? (ws as any).getImages() : []; } catch { imgs = []; }
  const media: any[] = ((wb as any).model && (wb as any).model.media) || [];
  const sum = (arr: number[], n: number) => arr.slice(0, Math.max(0, n)).reduce((a, b) => a + b, 0);
  const out: string[] = [];
  for (const im of imgs) {
    const m = media.find((x) => x.name === im.imageId || x.key === im.imageId) || media[Number(im.imageId)];
    if (!m || !m.buffer) continue;
    let src = "";
    try {
      const buf: any = m.buffer;
      if (typeof buf === "string") {
        src = buf.startsWith("data:") ? buf : `data:image/${m.extension || "png"};base64,${buf}`;
      } else {
        const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf.buffer || buf);
        src = bytesToDataUri(bytes, m.extension || "png");
      }
    } catch { continue; }
    const tl = im.range?.tl || im.range;
    const br = im.range?.br;
    const ext = im.range?.ext;
    const left = (sum(colPx, tl?.nativeCol ?? 0) + (tl?.nativeColOff || 0) / EMU_PER_PX) * s;
    const top = (sum(rowPx, tl?.nativeRow ?? 0) + (tl?.nativeRowOff || 0) / EMU_PER_PX) * s;

    let width = 120, height = 60;
    if (ext && ext.cx) {
      // 读盘后 exceljs 给的是 EMU
      width = (ext.cx / EMU_PER_PX) * s;
      height = (ext.cy / EMU_PER_PX) * s;
    } else if (ext && ext.width) {
      // addImage 传入的 ext 读回是像素
      width = ext.width * s;
      height = (ext.height || ext.width * 0.5) * s;
    } else if (br && br.nativeCol != null) {
      const right = (sum(colPx, br.nativeCol) + (br.nativeColOff || 0) / EMU_PER_PX) * s;
      const bottom = (sum(rowPx, br.nativeRow) + (br.nativeRowOff || 0) / EMU_PER_PX) * s;
      width = Math.max(10, right - left);
      height = Math.max(10, bottom - top);
    }
    out.push(`<img src="${src}" style="position:absolute;left:${left.toFixed(1)}px;top:${top.toFixed(1)}px;width:${width.toFixed(1)}px;height:${height.toFixed(1)}px;object-fit:contain">`);
  }

  // 兜底：签名图没能下载成二进制（跨域/鉴权），但拿到了 URL，直接交给浏览器显示
  for (const u of (((wb as any).__urlImages as any[]) || [])) {
    const pos = addrToRowCol(String(u.cell || ""));
    if (!pos || !u.url) continue;
    const { row: r, col: c } = pos;
    if (c < 1 || r < 1 || c > colPx.length || r > rowPx.length) continue;
    const off = (u.offsetX || 0);
    const left = (sum(colPx, c - 1) + off) * s;
    const top = sum(rowPx, r - 1) * s;
    // 宽度用「偏移之后的剩余格宽」，避免越到右边一列；靠左显示以紧贴标签
    const width = Math.max(10, (colPx[c - 1] - off - 4) * s);
    const height = Math.max(8, (rowPx[r - 1] - 4) * s);
    out.push(`<img src="${esc(u.url)}" style="position:absolute;left:${left.toFixed(1)}px;top:${top.toFixed(1)}px;width:${width.toFixed(1)}px;height:${height.toFixed(1)}px;object-fit:contain;object-position:left center">`);
  }
  return out.join("");
}

/**
 * 预览与网页打印共用的页面样式：
 * - 屏幕：整页 A4，内容按源文件页边距内缩；面板比 A4 窄时整体等比缩小（zoom），只看效果
 * - 打印：@page margin 为 0（去掉浏览器页眉页脚），留白由 .print-page 的 padding 还原；逐页分页，一律 100%
 * - td 内边距为 0，保证与 Excel 的单元格几何一致
 */
export function pageCss(m: { l: number; r: number; t: number; b: number }, screenZoom = 1): string {
  return `
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; }
    body { background: #e8e8ec; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
    table { border-collapse: collapse; }
    td { padding: 0; }
    @media screen {
      body { zoom: ${screenZoom}; }
      .print-page {
        width: 210mm;
        padding: ${m.t}mm ${m.r}mm ${m.b}mm ${m.l}mm;
        margin: 10px auto;
        background: #fff;
        box-shadow: 0 1px 4px rgba(0,0,0,.15);
        overflow: hidden;
      }
    }
    /* 页边距交给 .print-page 自己留白，@page 设为 0：
       浏览器会在 @page 的 margin 区域自动打印「时间 + 文档标题」页眉页脚，
       margin 归零后这些抬头就没有绘制空间（打印样式里再由 padding 还原留白，版面不变）。 */
    @page { size: A4 portrait; margin: 0; }
    @media print {
      body { zoom: 1; background: #fff; }
      .print-page {
        width: 210mm;
        padding: ${m.t}mm ${m.r}mm ${m.b}mm ${m.l}mm;
        margin: 0;
        box-shadow: none;
        overflow: visible;
        page-break-after: always;
      }
      .print-page:last-child { page-break-after: auto; }
    }
  `;
}

/**
 * worksheet -> 单页 HTML。
 * 列宽/行高按 Excel 几何换算成 px，再统一乘缩放系数（超出一页时等比缩小，等于「适应一页」）。
 * 返回 pageWmm = 表格实际宽度，配合页面左右 padding 正好对齐源文件版面。
 */
export function sheetToHtml(wb: ExcelJSNamespace.Workbook, ws: ExcelJSNamespace.Worksheet, opts: SheetHtmlOptions = {}): SheetHtml {
  const props: any = (ws as any).properties || {};
  const defRowH = props.defaultRowHeight || DEFAULT_ROW_H;
  const defColW = props.defaultColWidth || DEFAULT_COL_W;

  const maxRow = Math.min(ws.rowCount || 50, 150);
  const maxCol = Math.min(ws.columnCount || 16, 30);

  // Excel 几何
  const colPx: number[] = [];
  let contentW = 0;
  for (let c = 1; c <= maxCol; c++) {
    const px = colWidthPx(ws.getColumn(c).width || defColW);
    colPx.push(px);
    contentW += px;
  }
  const rowPx: number[] = [];
  let contentH = 0;
  for (let r = 1; r <= maxRow; r++) {
    const px = rowHeightPx(ws.getRow(r).height || defRowH);
    rowPx.push(px);
    contentH += px;
  }

  // 页边距（mm）与可打印区
  const mgRaw: any = (ws as any).pageSetup?.margins || {};
  const mg = {
    l: Math.round((mgRaw.left ?? 0.7) * 25.4),
    r: Math.round((mgRaw.right ?? 0.7) * 25.4),
    t: Math.round((mgRaw.top ?? 0.75) * 25.4),
    b: Math.round((mgRaw.bottom ?? 0.75) * 25.4),
  };
  const availW = (210 - mg.l - mg.r) * PX_PER_MM;
  const availH = (297 - mg.t - mg.b) * PX_PER_MM;
  const s = Math.min(1, availW / contentW, availH / contentH);

  const merges = collectMerges(ws, maxRow, maxCol);
  const colDefs = colPx.map((px) => `<col style="width:${(px * s).toFixed(2)}px">`).join("");

  const rowsHtml: string[] = [];
  for (let r = 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    const h = rowPx[r - 1] * s;
    let tds = "";
    for (let c = 1; c <= maxCol; c++) {
      const g = merges.get(`${r}_${c}`);
      if (g && g.rowSpan === 0) continue;
      const cell = row.getCell(c);
      const isMerged = !!g && (g.rowSpan > 1 || g.colSpan > 1);
      const border = isMerged ? mergedBorder(ws, r, c, g) : ((cell as any).border || {});
      // 不换行且靠左的格子，右侧为空时按 Excel 行为向外溢出，否则裁切
      const al: any = ((cell as any).style || {}).alignment || {};
      const rightCol = c + (g?.colSpan || 1);
      // 超出表格右边界视为空（不要 ws.getCell 造出表外单元格，否则会撑大 columnCount）
      const rightCell = rightCol <= maxCol ? ws.getCell(r, rightCol) : null;
      const rightEmpty = rightCol > maxCol || isBlankCell(rightCell);
      const canOverflow = rightEmpty && (al.horizontal == null || al.horizontal === "left" || al.horizontal === "general");
      // 右对齐/居中的长文字，Excel 会向左溢出到空格子；这里同样允许，否则会被裁成「Rev.」这类残缺文本
      const leftCell = c > 1 ? ws.getCell(r, c - 1) : null;
      const canOverflowLeft = isBlankCell(leftCell) && (al.horizontal === "right" || al.horizontal === "center");
      const style = cellCss(cell, s, border, canOverflow, canOverflowLeft);
      const span = isMerged ? ` rowspan="${g!.rowSpan}" colspan="${g!.colSpan}"` : "";
      let inner = cellHtml(cell, s);
      // 文字位移（字符）：左对齐无法向左移，用 transform 实现，只影响打印渲染
      const shift = opts.shift?.[cell.address];
      if (shift) inner = `<span style="display:inline-block;transform:translateX(${shift}ch)">${inner}</span>`;
      tds += `<td${span} style="${style}">${inner}</td>`;
    }
    rowsHtml.push(`<tr style="height:${h.toFixed(2)}px">${tds}</tr>`);
  }

  const tableW = (contentW * s).toFixed(2);
  const table = `<table style="width:${tableW}px;border-collapse:collapse;table-layout:fixed;margin:0">${colDefs}${rowsHtml.join("")}</table>`;
  const imgs = imagesHtml(wb, ws, colPx, rowPx, s);
  const inner = imgs
    ? `<div style="position:relative;width:${tableW}px;height:${(contentH * s).toFixed(2)}px">${imgs}<div style="position:absolute;left:0;top:0">${table}</div></div>`
    : table;

  return {
    html: `<div class="print-page">${inner}</div>`,
    margins: mg,
    pageWmm: Math.round(contentW * s * MM_PER_PX * 10) / 10,
    scale: s,
  };
}
