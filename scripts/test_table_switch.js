// 表格切换监听的冒烟测试：用桩替身顶替 SDK，验证「换表才刷新、同表选区变化不触发」
// 运行：node scripts/test_table_switch.js
const { execSync } = require("child_process");
const path = require("path");
const fs = require("fs");
const Module = require("module");

const ROOT = path.resolve(__dirname, "..");
const BUILD = path.join(ROOT, "scripts/.build");

execSync(
  `npx tsc src/sdk.ts src/mapping.ts src/assets.d.ts --outDir scripts/.build --module commonjs --target ES2020 --esModuleInterop --skipLibCheck --moduleResolution node`,
  { cwd: ROOT, stdio: "inherit" }
);

// mapping.ts 里 import 了 xlsx（webpack 下是资源 url），node 下用路径桩 + 拷贝模板
require.extensions[".xlsx"] = (mod, filename) => { mod.exports = filename; };
fs.mkdirSync(path.join(BUILD, "templates"), { recursive: true });
for (const f of fs.readdirSync(path.join(ROOT, "src/templates"))) {
  fs.copyFileSync(path.join(ROOT, "src/templates", f), path.join(BUILD, "templates", f));
}

// ---------- 桩：两套表 + 事件注册表 ----------
const handlers = { onSelectionChange: [], onTableAdd: [], onTableDelete: [] };
const tableA = {
  id: "tblA",
  getName: async () => "产品发机评审单（主表）",
  getFieldMetaList: async () => [{ id: "f1", name: "项目号", type: 1, isPrimary: true }],
  getRecordListByPage: async () => ({ records: [{ recordId: "r1", fields: { f1: "A-001" } }], hasMore: false }),
};
const tableB = {
  id: "tblB",
  getName: async () => "FAT 申请表",
  getFieldMetaList: async () => [{ id: "g1", name: "客户名称", type: 1, isPrimary: true }],
  getRecordListByPage: async () => ({ records: [{ recordId: "r9", fields: { g1: "客户X" } }], hasMore: false }),
};
const tableReviewTarget = {
  id: "tblReviewTarget",
  getName: async () => "FAT发机评审单",
  getFieldMetaList: async () => [{ id: "f1", name: "项目号", type: 1, isPrimary: true }],
  getRecordListByPage: async () => ({ records: [{ recordId: "review1", fields: { f1: "R-001" } }], hasMore: false }),
};
const tableFatTarget = {
  id: "tblFatTarget",
  getName: async () => "FAT申请单",
  getFieldMetaList: async () => [],
  getRecordListByPage: async () => ({ records: [], hasMore: false }),
};
const otherTable = {
  id: "tblOther",
  getName: async () => "项目资料",
  getFieldMetaList: async () => { throw new Error("非目标表不应读取字段"); },
  getRecordListByPage: async () => { throw new Error("非目标表不应读取记录"); },
};
let active = tableA;

const fakeBase = {
  getActiveTable: async () => active,
  getTableMetaList: async () => Promise.all([tableA, tableB, tableReviewTarget, tableFatTarget, otherTable].map(async (table) => ({ id: table.id, name: await table.getName() }))),
  onSelectionChange: (cb) => { handlers.onSelectionChange.push(cb); return () => handlers.onSelectionChange.pop(); },
  onTableAdd: (cb) => { handlers.onTableAdd.push(cb); return () => handlers.onTableAdd.pop(); },
  onTableDelete: (cb) => { handlers.onTableDelete.push(cb); return () => handlers.onTableDelete.pop(); },
};
const fakeUi = {
  switchToTable: async (tableId) => {
    active = [tableA, tableB, tableReviewTarget, tableFatTarget, otherTable].find((table) => table.id === tableId);
    return !!active;
  },
};
const fakeApi = { base: fakeBase, ui: fakeUi };

// 拦截 SDK 的动态 import（TS 编译为 CJS 后会走 require）
const origLoad = Module._load;
Module._load = function (req, ...rest) {
  if (req === "@lark-base-open/js-sdk") return { bitable: fakeApi };
  return origLoad.call(this, req, ...rest);
};

const fire = (tableId) => handlers.onSelectionChange.forEach((cb) => cb({ data: { tableId } }));

let pass = 0, fail = 0;
const check = (name, ok, extra) => {
  if (ok) { pass++; console.log("PASS", name); }
  else { fail++; console.log("FAIL", name, extra === undefined ? "" : `-> ${extra}`); }
};

(async () => {
  const sdk = require(path.join(BUILD, "sdk.js"));

  // 1. 初始读取
  let t = await sdk.getTableData();
  check("读到当前表名", t.name === "产品发机评审单（主表）", t.name);
  check("读到 tableId", t.tableId === "tblA", t.tableId);
  check("读到记录（字段按名索引）", t.records.length === 1 && t.records[0].fields["项目号"] === "A-001", JSON.stringify(t.records));

  // 2. 订阅
  const got = [];
  const off = await sdk.onActiveTableChange((id) => got.push(id));

  // 3. 同一张表内的选区变化 -> 不触发
  fire("tblA");
  fire("tblA");
  check("同表选区变化不触发刷新", got.length === 0, got.length);

  // 4. 切到 B -> 触发一次
  active = tableB;
  fire("tblB");
  check("切换到 B 触发一次", got.length === 1 && got[0] === "tblB", JSON.stringify(got));

  // 5. 重复事件 -> 不重复触发
  fire("tblB");
  check("重复事件不重复触发", got.length === 1, got.length);

  // 6. 切换后重新读取拿到 B 的数据
  t = await sdk.getTableData();
  check("刷新后表名更新为 B", t.name === "FAT 申请表", t.name);
  check("刷新后记录更新为 B", t.records.length === 1 && t.records[0].recordId === "r9", JSON.stringify(t.records));

  // 7. 空 tableId（记录变更类事件可能不带）不触发
  fire(null);
  fire(undefined);
  check("事件缺 tableId 时不触发", got.length === 1, got.length);

  // 8. 取消监听
  off();
  active = tableA;
  fire("tblA");
  check("取消监听后不再回调", got.length === 1, got.length);

  // 9. 模板随表名显示：切到哪张表只出现对应模板
  const { sheetsForTable } = require(path.join(BUILD, "mapping.js"));
  const ids = (name) => sheetsForTable(name).map((s) => s.id).join(",");
  check("FAT申请单 -> 只出 FAT 模板", ids("FAT申请单") === "fat", ids("FAT申请单"));
  check("内部FAT申请表 -> 只出 FAT 模板", ids("内部FAT申请表") === "fat", ids("内部FAT申请表"));
  check("内部FAT实施申请单 -> 只出 FAT 模板", ids("内部FAT实施申请单") === "fat", ids("内部FAT实施申请单"));
  check("产品发机评审单 -> 只出评审单模板", ids("产品发机评审单") === "review", ids("产品发机评审单"));
  check("发机评审信息 -> 只出评审单模板", ids("发机评审信息") === "review", ids("发机评审信息"));
  check("FAT发机评审单 -> 评审单模板（FAT 不干扰）", ids("FAT发机评审单") === "review", ids("FAT发机评审单"));
  check("只写 FAT 的表名 -> FAT 模板兜底", ids("FAT") === "fat", ids("FAT"));
  check("未匹配的表名 -> 模板全显示", ids("表1") === "review,fat", ids("表1"));
  check("表名未知(空) -> 模板全显示", ids("") === "review,fat", ids(""));

  // 10. 非目标表：只读取轻量表信息，返回两个固定入口，不触碰字段/记录
  active = otherTable;
  const gateway = await sdk.getPrintTableData();
  check("其他表进入选表分支", gateway.supported === false && gateway.currentTable.name === "项目资料", JSON.stringify(gateway));
  check("选表分支提供两个指定目标", gateway.targetTables.map((t) => t.name).join(",") === "FAT发机评审单,FAT申请单", JSON.stringify(gateway.targetTables));
  check("非目标表不返回业务数据", gateway.data === undefined, typeof gateway.data);
  check("无关申请表不能绕过选表分支", sdk.isSupportedPrintTable("其他申请表") === false, "");
  check("FAT发机评审单允许直接打印", sdk.isSupportedPrintTable("FAT发机评审单") === true, "");
  check("FAT申请单允许直接打印", sdk.isSupportedPrintTable("FAT申请单") === true, "");

  // 11. 用户选择目标表：调用宿主切表，再读取目标表的打印数据
  const switched = await sdk.switchToTable("tblReviewTarget");
  const targetData = await sdk.getPrintTableData();
  check("选择后切到 FAT发机评审单", switched && targetData.supported && targetData.currentTable.id === "tblReviewTarget", targetData.currentTable.id);
  check("切表后沿用原打印数据链路", targetData.data && targetData.data.records[0].recordId === "review1", JSON.stringify(targetData.data));

  console.log(`\n结果：${pass} PASS / ${fail} FAIL`);
  process.exit(fail ? 1 : 0);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
