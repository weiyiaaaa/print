// 生成「内部FAT申请单」的预览 HTML（填一条样例数据），用于核对表头位置与版面。
// 用法：先 tsc 编译到 scripts/.build（见 verify.js 第一步），再 node scripts/preview-fat.js
const path = require("path");
const fs = require("fs");

const ROOT = path.resolve(__dirname, "..");
const BUILD = path.join(ROOT, "scripts/.build");
const OUT = path.join(ROOT, "scripts/out");

require.extensions[".xlsx"] = (mod, filename) => { mod.exports = filename; };

const { fillSheet } = require(path.join(BUILD, "excel.js"));
const { SHEETS } = require(path.join(BUILD, "mapping.js"));
const { sheetToHtml, pageCss } = require(path.join(BUILD, "preview.js"));
const ExcelJS = require("exceljs");

const values = {
  "项目号": "C250911",
  "客户/机型": "农夫山泉广西大明山",
  "申请次数": "3",
  "申请单编号": "22-FAT20261105002",
  "发行日期": new Date("2026-11-05T00:00:00").getTime(),
  "FAT检开始时间": new Date("2026-11-06T14:30:00").getTime(),
  "缺件及未达事项": "缺防尘盖 2 件\n未张贴铭牌",
  "发生原因说明": "供应商到货延迟",
  "预期处理完成时间": new Date("2026-11-20T00:00:00").getTime(),
  // 检验结果三选一 + 附条件清单 + 关连整改通知书编号
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

const record = { recordId: "rec_fat", fields: {} };
const fieldsByName = new Map();
for (const [k, v] of Object.entries(values)) {
  record.fields[k] = v;
  fieldsByName.set(k, { fieldId: "f_" + k, name: k, type: k.includes("日期") || k.includes("时间") ? 5 : 1 });
}

(async () => {
  const sheet = SHEETS.find((s) => s.id === "fat");
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(sheet.file);
  await fillSheet(wb, sheet, record, fieldsByName, async () => null);

  const { html, margins, pageWmm } = sheetToHtml(wb, wb.worksheets[0], { shift: sheet.textShift });
  fs.mkdirSync(OUT, { recursive: true });
  const file = path.join(OUT, "preview-fat.html");
  fs.writeFileSync(
    file,
    `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>预览 · ${sheet.name}</title><style>${pageCss(margins, pageWmm)}</style></head><body>${html}</body></html>`,
    "utf8"
  );

  const shifted = (html.match(/translateX\(-[\d.]+ch\)/g) || []).length;
  console.log("预览已生成：", file);
  console.log(`带位移的表头单元格：${shifted} 个（标签 K1/K2/K3）`);

  const noCell = wb.worksheets[0].getCell("L2");
  console.log("L2 内容：", JSON.stringify(noCell.value), " 字号：", noCell.font && noCell.font.size);

  // 检验结果：方框涂黑 / 附条件 / 通知书编号
  const ws = wb.worksheets[0];
  const filled = ["C27", "C29", "C37"].filter((a) => {
    const f = ws.getCell(a).fill;
    return f && f.fgColor && String(f.fgColor.argb).toUpperCase().endsWith("000000");
  });
  console.log("涂黑的方框：", filled.join(",") || "无", "（期望只有 C29 = 附条件合格）");
  for (const a of ["F31", "F32", "F33", "F34", "F35"]) {
    console.log(`  ${a}:`, JSON.stringify(ws.getCell(a).value ?? ""));
  }
  console.log("H29 关连整改通知书：", JSON.stringify(ws.getCell("H29").value ?? ""));
  console.log("H37 关连整改通知书（合格/附条件时应为空）：", JSON.stringify(ws.getCell("H37").value ?? ""));
  console.log("C40 整改处理结果：", JSON.stringify(ws.getCell("C40").value ?? ""));
  console.log("L5 申请回数：", JSON.stringify(ws.getCell("L5").value ?? ""));
  const okBoxes=["H9","H10","H11","H12","H13","H14"].filter(a=>{const f=ws.getCell(a).fill;return f&&f.fgColor&&String(f.fgColor.argb).toUpperCase().endsWith("000000");});
  console.log("确认结果涂黑：", okBoxes.join(",")||"无", "（期望 H9,H12,H13）");
  console.log("L3 发行确认：", JSON.stringify(ws.getCell("L3").value ?? ""));
  console.log("C23 申请人：", JSON.stringify(ws.getCell("C23").value ?? ""));
  console.log("C16 现申请时间：", JSON.stringify(ws.getCell("C16").value ?? ""));
})();
