// 验收脚本：用固定的映射表分别填充「产品发机评审单」与「内部FAT申请表」，
// 与填写样例逐项比对，并校验几何保真（行高/列宽/页面设置/边框/样式）。
// 运行：node scripts/verify.js
const { execSync } = require("child_process");
const path = require("path");
const fs = require("fs");
const zlib = require("zlib");
const ExcelJS = require("exceljs");

const ROOT = path.resolve(__dirname, "..");
const BUILD = path.join(ROOT, "scripts/.build");
const REF = "C:/Users/WeiYi/Desktop/飞书打印插件/打印实例excel/产品发机评审单.xlsx";

// ---------- 1. 编译 TS -> CJS ----------
execSync(
  `npx tsc src/excel.ts src/sdk.ts src/mapping.ts src/preview.ts src/assets.d.ts --outDir scripts/.build --module commonjs --target ES2020 --esModuleInterop --skipLibCheck --moduleResolution node`,
  { cwd: ROOT, stdio: "inherit" }
);

// TS 里 import xlsx 是资源 url；node 下用路径桩替代（先把模板拷到编译产物同级目录）
require.extensions[".xlsx"] = (mod, filename) => { mod.exports = filename; };
fs.mkdirSync(path.join(BUILD, "templates"), { recursive: true });
for (const f of ["template_review.xlsx", "template_fat.xlsx"]) {
  fs.copyFileSync(path.join(ROOT, "src/templates", f), path.join(BUILD, "templates", f));
}

const { fillSheet } = require(path.join(BUILD, "excel.js"));
const { SHEETS } = require(path.join(BUILD, "mapping.js"));

// ---------- 2. 造一张可见的 PNG（400x150），验证签名按比例缩放 ----------
function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}
function makePng(w, h) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8; ihdr[9] = 6;
  const stride = 1 + w * 4;
  const raw = Buffer.alloc(h * stride);
  // 画可见内容：2px 深色边框 + 中间一条斜线（模拟手写笔画），便于预览里确认签名位置
  for (let y = 0; y < h; y++) {
    raw[y * stride] = 0; // 滤波类型
    for (let x = 0; x < w; x++) {
      const o = y * stride + 1 + x * 4;
      const border = x < 2 || x >= w - 2 || y < 2 || y >= h - 2;
      const stroke = Math.abs(y / h - x / w) < 0.08;
      const ink = border || stroke;
      raw[o] = ink ? 40 : 255;
      raw[o + 1] = ink ? 40 : 255;
      raw[o + 2] = ink ? 40 : 255;
      raw[o + 3] = ink ? 255 : 0;
    }
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}
const PNG = new Uint8Array(makePng(400, 150)); // 接近真实手写签名比例

// ---------- 3. 测试数据（对照样例里的值） ----------
const sheet = SHEETS.find((s) => s.id === "review");
const values = {
  // 基本信息
  "项目号": "C250910",
  "客户名称": "印度信实热线H10",
  "机型": "吹瓶机",
  "申请评审部门": "吹灌部装一组",
  "单据编号": "188",
  "申请日期": 1735689600000, // 2025-01-01
  "添附资料": "有 1 张",
  // 残留问题点
  "残留问题点": "壳体划伤\n密封圈渗油",
  "原因&应对方案": "更换外壳并抛光处理,更换密封圈并做气密测试",
  "预完成日期": 1736035200000, // 2025-01-05
  // 品管判定
  "判定结果": "附条件合格（按评审结果执行）",
  "附条件合格影响等级": "二级",
  "判定理由/说明/处置建议": "问题已进行整改可以进行拆机",
  "问题对应的主管总监": "吴国华",
  "销售代表": "张三",
  // 评审环节
  "主管总监评审结果": "允许发机",
  "主管总监评审意见": "同意发机，后续整改",
  "销售代表评审结果": "不允许发机（客户未确认）",
  // 结果跟踪
  "结案时间": 1736294400000, // 2025-01-08
};
const record = { recordId: "rec1", fields: {} };
const fieldsByName = new Map();
for (const [k, v] of Object.entries(values)) {
  record.fields[k] = v;
  fieldsByName.set(k, { fieldId: "f_" + k, name: k, type: k.includes("日期") ? 5 : 1 });
}

// ---------- 4. 填充 ----------
/** 规范化序列化：键排序，避免键顺序造成的假差异 */
function canon(v) {
  if (v === null || typeof v !== "object") return JSON.stringify(v);
  if (Array.isArray(v)) return "[" + v.map(canon).join(",") + "]";
  return "{" + Object.keys(v).sort().map((k) => JSON.stringify(k) + ":" + canon(v[k])).join(",") + "}";
}

const results = [];

function check(name, cond, detail) {
  results.push({ name, ok: !!cond, detail });
  console.log(`${cond ? "PASS" : "FAIL"}  ${name}${cond || !detail ? "" : "  <-- " + detail}`);
}

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(sheet.file);
  const srcHeights = [];
  for (let r = 1; r <= wb.worksheets[0].rowCount; r++) srcHeights.push(wb.worksheets[0].getRow(r).height ?? null);

  let imgCalls = 0;
  await fillSheet(wb, sheet, record, fieldsByName, async (names) => {
    imgCalls++;
    return names[0] === "记录人签名" ? null : { bytes: PNG }; // 最后一格留空，验证「无图不插」
  }, { docNo: "TL/TR-RC008-A0" });

  // 先落盘再重新读取：后续所有校验都针对「最终交付的文件」，而不是内存态
  fs.mkdirSync(path.join(ROOT, "scripts/out"), { recursive: true });
  const outPath = path.join(ROOT, "scripts/out/review-filled.xlsx");
  await wb.xlsx.writeFile(outPath);
  const finalWb = new ExcelJS.Workbook();
  await finalWb.xlsx.readFile(outPath);
  const ws = finalWb.worksheets[0];
  const txt = (a) => {
    const v = ws.getCell(a).value;
    if (v == null) return "";
    if (typeof v === "object") {
      if (v.richText) return v.richText.map((s) => s.text).join("");
      if ("text" in v) return String(v.text);
      if ("result" in v) return String(v.result);
      return "";
    }
    return String(v);
  };

  // 4.1 合并区（对照样例：C10:D15 / E10:E15 / F10:F15 / E20:G22）
  const merges = ws.model.merges || [];
  const has = (m) => merges.includes(m);
  check("合并区 C10:D15", has("C10:D15"), merges.join(","));
  check("合并区 E10:E15", has("E10:E15"));
  check("合并区 F10:F15", has("F10:F15"));
  check("合并区 E20:G22", has("E20:G22"));
  check("明细区逐行小合并已拆除", !merges.some((m) => /^C1[0-5]:D1[0-5]$/.test(m) && m.slice(1, 3) === m.slice(4, 6)), merges.join(","));
  check("判定区逐行小合并已拆除", !merges.some((m) => /^E2[0-2]:G2[0-2]$/.test(m) && m.slice(1, 3) === m.slice(4, 6)), merges.join(","));

  // 4.2 取值（含 append 与 fmt）
  check("D5 项目号", txt("D5") === "C250910", txt("D5"));
  check("D6 客户名称", txt("D6") === "印度信实热线H10", txt("D6"));
  check("E6 机型", txt("E6") === "吹瓶机", txt("E6"));
  check("F6 申请评审部门", txt("F6") === "吹灌部装一组", txt("F6"));
  check("F2 单据编号（界面可编辑，默认 TL/TR-RC008-A0）", txt("F2") === "单据编号：TL/TR-RC008-A0", txt("F2"));
  check("F3 申请日期追加+格式化", txt("F3") === "申请日期：2025-01-01", txt("F3"));
  check("F23 主管总监", txt("F23") === "吴国华", txt("F23"));
  check("C10 残留问题点(多行拆分)", txt("C10") === "壳体划伤\n密封圈渗油", JSON.stringify(txt("C10")));
  check("E10 原因&应对方案", txt("E10").startsWith("更换外壳并抛光处理"), txt("E10"));
  check("F10 预完成日期格式化", txt("F10") === "2025-01-05", txt("F10"));
  check("E20 判定理由", txt("E20") === "问题已进行整改可以进行拆机", txt("E20"));

  // 4.2c 单据编号：界面留空时回退到表格字段，模板里写死的旧编号必须被覆盖
  {
    const wbFb = new ExcelJS.Workbook();
    await wbFb.xlsx.readFile(sheet.file);
    await fillSheet(wbFb, sheet, record, fieldsByName, async () => null, { docNo: "" });
    const v = wbFb.worksheets[0].getCell("F2").value;
    const t = typeof v === "string" ? v : "";
    check("F2 留空时回退表格「单据编号」字段", t === "单据编号：188", t);
    check("模板旧编号 C/E-20- 已被覆盖", !t.includes("C/E-20"), t);
  }

  // 4.2b 复选框 / 追加意见 / 结果跟踪
  check("E8 添附资料「有」已勾选", txt("E8") === "添附资料：有 ■（共    张）；无 □", txt("E8"));
  check("C22 附条件合格已勾选", txt("C22") === "附条件合格（按评审结果执行）■", txt("C22"));
  check("C20 合格未勾选", txt("C20") === "合格（允许发机）□", txt("C20"));
  check("C21 不合格未勾选", txt("C21") === "不合格（禁止发机）□", txt("C21"));
  check("C24 影响度等级 II 已勾选", txt("C24") === "I:□  /II:■  /III:□  /IV:□", txt("C24"));
  check("C32 主管总监「允许」已勾选", txt("C32") === "评审结果：允许发机 ■；不允许发机 □", txt("C32"));
  check("C36 销售代表「不允许」已勾选", txt("C36") === "评审结果：允许发机 □；不允许发机 ■", txt("C36"));
  check("C33 处理意见追加", txt("C33") === "处理意见/说明：同意发机，后续整改", txt("C33"));
  check("F24 销售代表", txt("F24") === "张三", txt("F24"));
  check("G46 结案时间格式化", txt("G46") === "2025-01-08", txt("G46"));

  // 4.3 居中 + 换行
  const al = (a) => ws.getCell(a).alignment || {};
  check("D5 居中", al("D5").horizontal === "center" && al("D5").vertical === "middle", JSON.stringify(al("D5")));
  check("C10 居中且换行", al("C10").horizontal === "center" && al("C10").wrapText === true, JSON.stringify(al("C10")));
  check("F3 追加格不改对齐", !al("F3").horizontal || al("F3").horizontal === "left", JSON.stringify(al("F3")));

  // 4.4 表单长度：行高与源文件逐一相同（R29 例外：为放手写签名图已加高）
  const heightDiff = [];
  for (let r = 1; r <= srcHeights.length; r++) {
    if (r === 29) continue;
    const now = ws.getRow(r).height ?? null;
    if (srcHeights[r - 1] !== now) heightDiff.push(`R${r}:${srcHeights[r - 1]}->${now}`);
  }
  check("行高与源文件一致（仅 R29 为签名加高）", heightDiff.length === 0, heightDiff.join(","));
  check("R29 已加高以容纳签名图", (ws.getRow(29).height || 0) >= 24, `R29=${ws.getRow(29).height}pt`);
  check("行数未变化", ws.rowCount === srcHeights.length, `${ws.rowCount} vs ${srcHeights.length}`);

  // 4.5 签名图片
  const imgs = ws.getImages ? ws.getImages() : [];
  const media = (finalWb.model && wb.model.media) || [];
  check("图片总数 = 模板LOGO1 + 签名7", imgs.length === 8, `imgs=${imgs.length} media=${media.length}`);
  check("图片像素已写入工作簿", media.length >= 8, `media=${media.length}`);
  const sig = imgs[imgs.length - 1];
  const sigAnchor = sig && sig.range && sig.range.tl ? `col=${sig.range.tl.nativeCol},row=${sig.range.tl.nativeRow}` : "none";
  check("签名锚点为单元格位置", !!(sig && sig.range && sig.range.tl), sigAnchor);
  // 两张签名图各紧贴对应标签右侧：图1 在 D29 内偏移 68px，图2 在 G29 左端
  const imgAt = (col, row) => imgs.find((i) => i.range && i.range.tl && i.range.tl.nativeCol === col && i.range.tl.nativeRow === row);
  const d29img = imgAt(3, 28), g29img = imgAt(6, 28);
  const d29off = ((d29img && d29img.range.tl.nativeColOff) || 0) / 9525;
  const g29off = ((g29img && g29img.range.tl.nativeColOff) || 0) / 9525;
  check("品管经理签名图在 D29（紧贴标签右侧）", !!d29img && d29off >= 60, `offset=${d29off.toFixed(1)}px`);
  check("品管总监签名图在 G29 左端", !!g29img && g29off < 10, `offset=${g29off.toFixed(1)}px`);
  check("无图字段未插入", imgCalls === 8, `calls=${imgCalls}`);

  // 4.6 与样例文件的结构比对
  const refWb = new ExcelJS.Workbook();
  await refWb.xlsx.readFile(REF);
  const refMerges = refWb.worksheets[0].model.merges || [];
  const missing = refMerges.filter((m) => !merges.includes(m));
  check("样例的合并结构全部具备", missing.length === 0, "缺:" + missing.join(","));

  // 4.7 预览/网页打印用 HTML（几何按 Excel 换算，页边距与源文件一致）
  const { sheetToHtml } = require(path.join(BUILD, "preview.js"));
  const { html, margins, pageWmm, scale } = sheetToHtml(finalWb, ws, { shift: sheet.textShift });
  check("预览 HTML 生成", html.includes('class="print-page"'), html.slice(0, 60));
  check("单页自适应缩放（85%~100%）", scale >= 0.85 && scale <= 1, "scale=" + scale.toFixed(4));
  check("不含 CSS zoom（尺寸已按比例换算）", !html.includes("zoom:"), "html 里仍有 zoom");
  const tablePx = parseFloat((/width:([\d.]+)px;border-collapse/.exec(html) || [])[1] || "0");
  const tableMm = (tablePx * 25.4) / 96;
  check("表格宽度 = 可打印宽度（mm）", Math.abs(tableMm - pageWmm) < 0.6, `table=${tableMm.toFixed(1)}mm page=${pageWmm}mm`);
  const trH = [...html.matchAll(/<tr style="height:([\d.]+)px"/g)].reduce((a, m) => a + parseFloat(m[1]), 0);
  const contentMm = (trH * 25.4) / 96;
  check("内容高度不超 A4 可打印高度", contentMm <= 297 - margins.t - margins.b + 0.6, `${contentMm.toFixed(1)}mm`);
  check("内容宽度不超可打印宽度", pageWmm <= 210 - margins.l - margins.r + 0.6, `page=${pageWmm} 可打印=${210 - margins.l - margins.r}`);

  // 评审区块之间的间隔行必须有高度，否则各方框会贴在一起（曾为给折行腾高度被压到 1pt）
  const rowPx = [...html.matchAll(/<tr style="height:([\d.]+)px"/g)].map((m) => parseFloat(m[1]));
  const gaps = [34, 38, 42].map((r) => rowPx[r - 1] || 0);
  check("评审区块之间有间距（34/38/42 行）", gaps.every((g) => g >= 6), gaps.map((g) => g + "px").join("/"));

  // 合并块必须四条边都画出来（右/下边原本记在被覆盖的格子上，容易丢）
  const mergedTd = html.match(/<td rowspan="6" colspan="2" style="([^"]*)"/);
  const mb = mergedTd ? mergedTd[1] : "";
  check(
    "合并块四边边框齐全（C10:D15）",
    !!mergedTd && /border-top/.test(mb) && /border-bottom/.test(mb) && /border-left/.test(mb) && /border-right/.test(mb),
    mb.slice(0, 140)
  );

  // 富文本分段字号：标题里「（产品中心）」按 10pt 渲染，主标题仍是 20pt
  check(
    "标题「（产品中心）」按富文本字号渲染",
    /<span style="font-size:10\.\d+pt[^"]*">（产品中心）<\/span>/.test(html),
    (html.match(/<span style="font-size:[\d.]+pt[^"]*">（产品中心）/) || ["未找到"])[0]
  );

  // 表头左移一位（前导空格 5 -> 4）；第 29 行两个签名标签改左对齐 + 左缩进，右侧留出签名图位置
  const d2run = (ws.getCell("D2").value && ws.getCell("D2").value.richText) || [];
  check("表头前导空格已减为 4 个", d2run[0] && d2run[0].text === "    产品发机评审单", JSON.stringify(d2run[0] && d2run[0].text));
  const row29 = html.split("<tr ")[29] || "";
  check("第 29 行签名标签左对齐无缩进（与下方签名行一致）", /text-align:left/.test(row29) && !/padding-(left|right):\d/.test(row29), row29.slice(0, 120));
  const shift29 = (row29.match(/translateX\(-4ch\)/g) || []).length;
  check("第 29 行两个签名标签各左移 4 个字符", shift29 === 2, `命中 ${shift29} 处`);

  // 右侧文字不得被裁：Rev.2018 与签名标签都应完整显示
  const tdsOf = (rowIdx) => ((html.split("<tr ")[rowIdx] || "").match(/<td[^>]*>[\s\S]*?<\/td>/g) || []);
  const revTd = tdsOf(48).find((t) => /Rev\./.test(t) && t.includes("2018")) || "";
  check("Rev.2018 不再被裁切", /overflow:visible/.test(revTd), revTd.slice(0, 140));
  const labelTd = tdsOf(29).find((t) => t.includes("品管主管总监签名")) || "";
  check("「品管主管总监签名：」完整显示", labelTd.includes("品管主管总监签名："), labelTd.slice(0, 140));
  check("「品管主管总监签名：」字号 11pt（与下方签名行一致）", /font-size:11(\.\d+)?pt/.test(labelTd), (labelTd.match(/font-size:[\d.]+pt/) || [""])[0]);
  // 右侧括号说明字号 10pt 且不换行（按 Excel 单行显示）
  const noteTd = tdsOf(36).find((t) => t.includes("当异常发机申请由销售部门")) || "";
  check("右侧括号说明字号 10pt", /font-size:10(\.\d+)?pt/.test(noteTd), (noteTd.match(/font-size:[\d.]+pt/) || [""])[0]);
  check("括号说明不换行（单行显示）", !/pre-wrap/.test(noteTd), noteTd.slice(0, 140));
  check("括号说明文本完整", noteTd.includes("由项目管理部评审"), noteTd.slice(0, 140));
  check("Rev.2018 已在表格内（第48行右对齐格）", /text-align:right/.test(revTd), revTd.slice(0, 140));
  // 左侧小标题不再贴边（缩进 1 字符）
  const titleTd = tdsOf(18).find((t) => t.includes("品管部门判定")) || "";
  check("左侧小标题已缩进留白", /padding-left:1ch/.test(titleTd), titleTd.slice(0, 120));

  // 同时产出一份可直接用浏览器打开/打印的预览页，便于与 xlsx 逐项目视比对
  const { pageCss } = require(path.join(BUILD, "preview.js"));
  const previewPath = path.join(ROOT, "scripts/out/preview-review.html");
  fs.writeFileSync(
    previewPath,
    `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>预览 · ${sheet.name}</title><style>${pageCss(margins)}</style></head><body>${html}</body></html>`,
    "utf8"
  );

  // 4.8 保真度：列宽 / 页面设置 / 未填写单元格的样式必须与源文件一致
  const srcWb2 = new ExcelJS.Workbook();
  await srcWb2.xlsx.readFile(sheet.file);
  const sws = srcWb2.worksheets[0];

  const widthDiff = [];
  for (let c = 1; c <= Math.max(10, sws.columnCount); c++) {
    if ((sws.getColumn(c).width ?? null) !== (ws.getColumn(c).width ?? null)) {
      widthDiff.push(`C${c}:${sws.getColumn(c).width}->${ws.getColumn(c).width}`);
    }
  }
  check("列宽与源文件一致", widthDiff.length === 0, widthDiff.join(","));

  // 只比对影响打印结果的键（exceljs 写盘时固定带 useFirstPageNumber=true，起始页码仍为 1，与源文件等效）
  const PS_KEYS = ["paperSize", "orientation", "margins", "scale", "fitToPage", "fitToWidth", "fitToHeight",
    "printArea", "printTitlesRow", "printTitlesColumn", "horizontalCentered", "verticalCentered",
    "pageOrder", "blackAndWhite", "draft", "errors"];
  const pickPs = (ps) => canon(PS_KEYS.reduce((o, k) => ((o[k] = (ps || {})[k]), o), {}));
  const psSrc = pickPs(sws.pageSetup);
  const psOut = pickPs(ws.pageSetup);
  check("页面设置与源文件一致", psSrc === psOut, `src=${psSrc.slice(0, 80)} out=${psOut.slice(0, 80)}`);

  // 合并区四角边框（合并后 Excel 仍按边缘格绘制外框，必须保持源文件定义）
  const borderDiff = [];
  for (const addr of ["C10", "D10", "C15", "D15", "E10", "E15", "F10", "F15", "E20", "G20", "E22", "G22"]) {
    if (canon(sws.getCell(addr).border || {}) !== canon(ws.getCell(addr).border || {})) borderDiff.push(addr);
  }
  check("合并区边缘边框与源文件一致", borderDiff.length === 0, "差异:" + borderDiff.join(","));

  // 被映射写入/合并的格子不参与样式比对（它们本来就要变）
  const touched = new Set();
  for (const c of sheet.cells) touched.add(c.cell);
  for (const e of sheet.editable ?? []) touched.add(e.cell);
  for (const ck of sheet.checks ?? []) touched.add(ck.cell);
  for (const im of sheet.images) touched.add(im.cell);
  const addRange = (str) => {
    const [a, b] = str.split(":");
    const p = (x) => (/^([A-Z]+)(\d+)$/.exec(x) || []).slice(1);
    const [c1s, r1s] = p(a);
    const [c2s, r2s] = b ? p(b) : [c1s, r1s];
    const cn = (s) => s.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    for (let r = +r1s; r <= +r2s; r++) for (let c = cn(c1s); c <= cn(c2s); c++) touched.add(`${c}_${r}`);
  };
  for (const b of sheet.blocks) addRange(b.range);
  for (const m of sheet.merges ?? []) addRange(m);

  const styleOf = (cell) => canon({
    font: cell.font || null,
    border: cell.border || null,
    fill: cell.fill || null,
    numFmt: cell.numFmt || null,
    align: cell.alignment || null,
  });
  const styleDiff = [];
  let compared = 0;
  for (let r = 1; r <= sws.rowCount; r++) {
    for (let c = 1; c <= 10; c++) {
      const addr = ws.getCell(r, c).address;
      if (touched.has(addr) || touched.has(`${c}_${r}`)) continue;
      // 被合并覆盖的格子不单独渲染（Excel 按主格画），样式不参与比对
      if (ws.getCell(r, c).master !== ws.getCell(r, c)) continue;
      if (sws.getCell(r, c).master !== sws.getCell(r, c)) continue;
      compared++;
      const a = styleOf(sws.getCell(r, c));
      const b = styleOf(ws.getCell(r, c));
      if (a !== b) styleDiff.push(addr);
    }
  }
  check(`未填写单元格样式一致（比对 ${compared} 格）`, styleDiff.length === 0, "差异:" + styleDiff.slice(0, 8).join(","));

  console.log("\n输出文件:", outPath);

  // 4.9 URL 兜底：拿不到图片二进制（跨域/鉴权失败）时，预览用 <img src=url> 直接显示
  {
    const wbUrl = new ExcelJS.Workbook();
    await wbUrl.xlsx.readFile(path.join(ROOT, "src/templates/template_review.xlsx"));
    let k = 0;
    await fillSheet(wbUrl, sheet, { recordId: "recUrl", fields: {} }, new Map(), async () => ({ url: `https://example.com/s${++k}.png` }));
    const rUrl = sheetToHtml(wbUrl, wbUrl.worksheets[0], { shift: sheet.textShift });
    const cnt = (rUrl.html.match(/<img src="https:\/\/example\.com\/s\d+\.png"/g) || []).length;
    check("URL 兜底：8 张签名以 <img src> 渲染", cnt === 8, `count=${cnt}`);
    check(
      "URL 兜底：D29 图按 offsetX 68 定位（left=230）",
      /<img src="https:\/\/example\.com\/s3\.png"[^>]*left:230\.0px/.test(rUrl.html),
      (rUrl.html.match(/<img src="https:\/\/example\.com\/s3\.png"[^>]*style="[^"]*"/) || ["未命中"])[0].slice(0, 120)
    );
  }

  // 4.10 上传签名兜底：签批状态命中「已确认」才允许用本地上传的签名
  {
    const { matchStatus, SHEETS: S } = require(path.join(BUILD, "mapping.js"));
    check("状态「品管经理已确认」判定为已签", matchStatus("品管经理已确认") === true, "");
    check("状态「品管总监已确认」判定为已签", matchStatus("品管总监已确认") === true, "");
    check("状态「升级处理」判定为未签", matchStatus("升级处理") === false, "");
    check("状态为空判定为未签", matchStatus("") === false, "");
    check("自定义词表：允许发机", matchStatus("允许发机", ["允许发机"]) === true, "");
    check("自定义词表：不允许发机不应命中", matchStatus("不允许发机", ["允许发机"]) === false, "");
    const review = S.find((x) => x.id === "review");
    const noKey = review.images.filter((im) => !im.uploadKey);
    check("所有签名位都配了上传键（可在界面管理签名）", noKey.length === 0, noKey.map((i) => i.cell).join(","));
    const qc = review.images.find((im) => im.cell === "D29");
    check("品管经理签名位绑定「品管经理签批」", !!qc && (qc.statusField || []).includes("品管经理签批"), JSON.stringify(qc && qc.statusField));
  }

  // 4.11 内部FAT申请表：近几轮新增映射的回归验收（字段结果 / 涂黑 / 条件 / 签名 / 版面）
  {
    const { sheetToHtml: fatToHtml, pageCss: fatPageCss } = require(path.join(BUILD, "preview.js"));
    const fat = SHEETS.find((s) => s.id === "fat");
    const fvals = {
      "项目号": "C250911",
      "客户/机型": "农夫山泉广西大明山",
      "申请次数": "3",
      "申请单编号": "22-FAT20261105002",
      "发行日期": new Date("2026-11-05T00:00:00").getTime(),
      "FAT检开始时间": new Date("2026-11-06T14:30:00").getTime(),
      "缺件及未达事项": "缺防尘盖 2 件\n未张贴铭牌",
      "发生原因说明": "供应商到货延迟",
      "预期处理完成时间": new Date("2026-11-20T00:00:00").getTime(),
      "内部FAT检验结果": "附条件合格",
      "附条件事项清单": "补齐防尘盖并重新固定\n铭牌内容按图纸复核\n门锁开合顺畅性确认",
      "关连整改通知书编号": "ZG-2026-0118",
      "整改处理结果": "已完成整改",
      "验收记录确认": "完成",
      "工艺调试记录确认": "附条件完成",
      "瓶样检测合格表确认": "完成",
      "申请人": [{ name: "黄永鹏" }],
      "FAT检验人": [{ name: "吴浩龙" }, { name: "李晓宋" }],
    };
    const frecord = { recordId: "rec_fat", fields: {} };
    const fmap = new Map();
    for (const [k, v] of Object.entries(fvals)) {
      frecord.fields[k] = v;
      fmap.set(k, { fieldId: "f_" + k, name: k, type: /日期|时间/.test(k) ? 5 : 1 });
    }
    const fwb = new ExcelJS.Workbook();
    await fwb.xlsx.readFile(fat.file);
    await fillSheet(fwb, fat, frecord, fmap, async () => null);
    const fws = fwb.worksheets[0];
    const ft = (a) => {
      const v = fws.getCell(a).value;
      if (v == null) return "";
      if (typeof v === "object") return v.richText ? v.richText.map((x) => x.text).join("") : (v.text != null ? String(v.text) : "");
      return String(v);
    };
    const blackAt = (ws, a) => {
      const f = ws.getCell(a).fill;
      return !!(f && f.fgColor && String(f.fgColor.argb).toUpperCase().endsWith("000000"));
    };
    const black = (a) => blackAt(fws, a);

    check("FAT D5 项目号", ft("D5") === "C250911", ft("D5"));
    check("FAT G5 客户/机型", ft("G5") === "农夫山泉广西大明山", ft("G5"));
    check("FAT 申请单编号整格重写（模板写死前缀已去掉）", ft("L2") === "22-FAT20261105002" && !ft("L2").includes("C/E-22"), ft("L2"));
    check("FAT 长编号自动缩字号（6~10pt）", fws.getCell("L2").font && fws.getCell("L2").font.size <= 10 && fws.getCell("L2").font.size >= 6, String(fws.getCell("L2").font && fws.getCell("L2").font.size));
    check("FAT 发行日期写入 K1 标签右侧的值格", /^\d{4}-\d{2}-\d{2}$/.test(ft("L1")) && ft("K1") === "发行日期：", `K1=${ft("K1")} L1=${ft("L1")}`);
    check("FAT 现申请时间三处填空且整句宽度不变", /^现申请：\s+11\s+月\s+06\s+日\s+14\s+时 做FAT，请实施$/.test(ft("C16")) && ft("C16").length === "现申请：     月     日        时 做FAT，请实施".length, ft("C16"));
    {
      const { fmtApplyTime } = require(path.join(BUILD, "mapping.js"));
      check("FAT 开始时间解析：日期字符串含时分", fmtApplyTime("2026-11-06 14:30").join("/") === "11/06/14", fmtApplyTime("2026-11-06 14:30").join("/"));
      check("FAT 开始时间解析：下午 2 点转 24 小时制", fmtApplyTime("2026/11/6 下午2:05").join("/") === "11/06/14", fmtApplyTime("2026/11/6 下午2:05").join("/"));
      check("FAT 开始时间解析：只有日期时「时」留空", fmtApplyTime("2026-11-06").join("/") === "11/06/", fmtApplyTime("2026-11-06").join("/"));
      check("FAT 开始时间解析：空值不填", fmtApplyTime("").length === 0, "");
    }
    const fatRender = fatToHtml(fwb, fws, { shift: fat.textShift });
    check("FAT 预览含现申请时间", fatRender.html.includes("现申请："), "");
    const fatCssOut = fatPageCss(fatRender.margins);
    check("打印样式 @page margin 已归零（去掉浏览器页眉页脚）", /@page\s*\{[^}]*margin:\s*0\s*;/.test(fatCssOut), (fatCssOut.match(/@page\s*\{[^}]*\}/) || [""])[0]);
    check("打印留白改由 .print-page 的 padding 还原", /@media print[\s\S]*?\.print-page\s*\{[^}]*padding:/.test(fatCssOut), "");
    check("FAT 申请次数填空（第   3   回，原宽度不变）", /^第\s+3\s+回$/.test(ft("L5")) && ft("L5").length === 12, JSON.stringify(ft("L5")));
    check("FAT 缺件事项合并块", ft("C20").includes("缺防尘盖 2 件"), ft("C20"));
    check("FAT 发生原因合并块", ft("I20") === "供应商到货延迟", ft("I20"));
    check("FAT 预完成日期合并块", /^\d{4}-\d{2}-\d{2}$/.test(ft("L20")), ft("L20"));

    check("FAT 申请人接在「申请人：」后面", ft("C23") === "申请人：黄永鹏", ft("C23"));
    check("FAT 发行确认填 FAT 检验人（多人按、连接）", ft("L3") === "吴浩龙、李晓宋", ft("L3"));
    check("FAT 整改处理结果追加在标签后", ft("C40") === "整改通知书处理及附条件事项处理的结果：已完成整改", ft("C40"));
    check("FAT 附条件 1~3 条逐行写入", ft("F31") === "补齐防尘盖并重新固定" && ft("F32") === "铭牌内容按图纸复核" && ft("F33") === "门锁开合顺畅性确认", [ft("F31"), ft("F32"), ft("F33")].join("|"));
    check("FAT 附条件 4/5 行留空", ft("F34") === "" && ft("F35") === "", [ft("F34"), ft("F35")].join("|"));

    check("FAT 附条件合格只涂黑 C29", black("C29") && !black("C27") && !black("C37"), ["C27", "C29", "C37"].filter(black).join(","));
    check("FAT 附条件合格：编号写 H29，H37 只有标签", ft("H29") === "关连整改通知书No.:ZG-2026-0118" && ft("H37") === "关连整改通知书No.:", `H29=${ft("H29")} H37=${ft("H37")}`);
    check("FAT 三项确认涂黑 H9/H12/H13（附条件完成不误涂「完成」）", ["H9", "H12", "H13"].every(black) && !["H10", "H11", "H14"].some(black), ["H9", "H10", "H11", "H12", "H13", "H14"].filter(black).join(","));

    check("FAT 表头三标签按实测字宽分别补偿", JSON.stringify(fat.textShift) === JSON.stringify({ K1: -3, K2: -5, K3: -3 }), JSON.stringify(fat.textShift));
    const appImg = fat.images.find((im) => im.cell === "E23");
    check("FAT 申请人手签图挪到 E23 且不占用姓名文字字段", !!appImg && !appImg.fields.includes("申请人"), JSON.stringify(appImg));

    // 三种检验结果的分支 + 空值 + 带括号说明，逐一重新填充
    for (const [verdict, wantBox, wantNo] of [
      ["合格", "C27", null],
      ["附条件合格", "C29", "H29"],
      ["不合格", "C37", "H37"],
      ["附条件合格（按评审执行）", "C29", "H29"],
      ["", null, null],
    ]) {
      const r2 = { recordId: "r2", fields: { ...fvals, "内部FAT检验结果": verdict } };
      const w2 = new ExcelJS.Workbook();
      await w2.xlsx.readFile(fat.file);
      await fillSheet(w2, fat, r2, fmap, async () => null);
      const s2 = w2.worksheets[0];
      const boxes = ["C27", "C29", "C37"].filter((a) => blackAt(s2, a));
      const t2 = (a) => { const v = s2.getCell(a).value; return v == null ? "" : String(typeof v === "object" && v.richText ? v.richText.map((x) => x.text).join("") : v); };
      const withNo = ["H29", "H37"].filter((a) => t2(a).includes("ZG-2026-0118"));
      check(`FAT 结果「${verdict || "空"}」涂黑 ${wantBox || "无"}`, (wantBox ? boxes.length === 1 && boxes[0] === wantBox : boxes.length === 0), boxes.join(","));
      check(`FAT 结果「${verdict || "空"}」编号写在 ${wantNo || "无"}`, (wantNo ? withNo.length === 1 && withNo[0] === wantNo : withNo.length === 0), withNo.join(","));
    }

    // 保真：行高 / 列宽 / 页面设置与源模板一致
    const fsrc = new ExcelJS.Workbook();
    await fsrc.xlsx.readFile(fat.file);
    const sws = fsrc.worksheets[0];
    const hd = [], wd = [];
    for (let r = 1; r <= Math.max(sws.rowCount, fws.rowCount); r++) if ((sws.getRow(r).height || 0) !== (fws.getRow(r).height || 0)) hd.push(r);
    for (let c = 1; c <= Math.max(sws.columnCount, fws.columnCount); c++) if ((sws.getColumn(c).width || 0) !== (fws.getColumn(c).width || 0)) wd.push(c);
    check("FAT 行高与源文件一致", hd.length === 0, hd.join(","));
    check("FAT 列宽与源文件一致", wd.length === 0, wd.join(","));
    check("FAT 页面设置与源文件一致", JSON.stringify(sws.pageSetup) === JSON.stringify(fws.pageSetup), "");

    const fh = fatToHtml(fwb, fws, { shift: fat.textShift });
    check("FAT 预览含申请人姓名", fh.html.includes("申请人：黄永鹏"), "");
    check("FAT 预览含编号与整改结果", fh.html.includes("22-FAT20261105002") && fh.html.includes("已完成整改"), "");
    check("FAT 预览含发行确认与通知书编号", fh.html.includes("吴浩龙、李晓宋") && fh.html.includes("ZG-2026-0118"), "");
    check("FAT 预览表头位移命中 3 个标签", (fh.html.match(/translateX\(-[\d.]+ch\)/g) || []).length === 3, String((fh.html.match(/translateX\(-[\d.]+ch\)/g) || []).length));
  }

  // 4.12 评审单填入文字防溢出：长文本自动换行/缩字号，且不污染相邻的「签名：」标签
  {
    const vals = {
      "客户名称": "【E2E测试】走查专用客",  // 用户截图里溢出的实际值
      "申请评审部门": "吹灌部装一组很长很长的部门名称",
      "主管总监评审意见": "这是一段非常非常长的评审意见内容用于测试自动缩小字号会不会溢出边框",
    };
    const rec = { recordId: "rec_long", fields: {} };
    const map = new Map();
    for (const [k, v] of Object.entries(vals)) {
      rec.fields[k] = v;
      map.set(k, { fieldId: "f_" + k, name: k, type: 1 });
    }
    const w = new ExcelJS.Workbook();
    await w.xlsx.readFile(sheet.file);
    await fillSheet(w, sheet, rec, map, async () => null);
    const wsf = w.worksheets[0];
    check("客户名称按本格宽度缩字号（不再溢出到机型列）", (wsf.getCell("D6").font.size || 0) < 12 && (wsf.getCell("D6").font.size || 0) >= 9, String(wsf.getCell("D6").font.size));
    check("长部门名只缩字号、不换行（避免撑高分页）", (wsf.getCell("F6").font.size || 0) < 12 && !(wsf.getCell("F6").alignment || {}).wrapText, `size=${wsf.getCell("F6").font.size} wrap=${!!(wsf.getCell("F6").alignment || {}).wrapText}`);
    check("长评审意见缩字号且不溢出签名格", (wsf.getCell("C33").font.size || 0) < 12 && (wsf.getCell("F33").font.size || 0) === 11, `C33=${wsf.getCell("C33").font.size} F33=${wsf.getCell("F33").font.size}`);
    check("填入格均未开启换行（行高不变、不会分两页）", ["D5", "D6", "E6", "F6", "C33", "C37", "C41"].every((a) => !(wsf.getCell(a).alignment || {}).wrapText), ["D5", "D6", "E6", "F6", "C33", "C37", "C41"].filter((a) => (wsf.getCell(a).alignment || {}).wrapText).join(","));
  }

  const fail = results.filter((r) => !r.ok);  console.log(`\n断言 ${results.length} 项，通过 ${results.length - fail.length}，失败 ${fail.length}`);
  process.exit(fail.length ? 1 : 0);
})();
