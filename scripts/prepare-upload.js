// 上传前准备：校验 app.json / block.json 是否已填真实值，并把 block.json 复制进 dist。
// opdev upload ./dist 时，工具在项目根目录找 app.json（appId），在待上传目录找 block.json（blockTypeID），
// 所以 dist 里必须有一份 block.json。
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const DIST = path.join(ROOT, "dist");

function readJson(p, label) {
  if (!fs.existsSync(p)) {
    console.error(`[准备上传] 缺少 ${label}：${p}`);
    process.exit(1);
  }
  try {
    return JSON.parse(fs.readFileSync(p, "utf8"));
  } catch (e) {
    console.error(`[准备上传] ${label} 不是合法 JSON：${e.message}`);
    process.exit(1);
  }
}

const app = readJson(path.join(ROOT, "app.json"), "app.json");
const block = readJson(path.join(ROOT, "block.json"), "block.json");

const problems = [];

const appId = String(app.appId || "").trim();
if (!/^cli_[A-Za-z0-9]+$/.test(appId)) {
  problems.push(
    "app.json 里的 appId 没填。到飞书开发者后台 → 你的应用 → 凭证与基础信息，复制 App ID（形如 cli_axxxxxxxxxxxxxxxx）填进去。"
  );
}

const blockTypeId = String(block.blockTypeID || "").trim();
if (!/^blk_[A-Za-z0-9]{6,}$/.test(blockTypeId) || /^blk_x+$/i.test(blockTypeId)) {
  problems.push(
    "block.json 里的 blockTypeID 还是占位符。到开发者后台 → 添加应用能力 → 多维表格插件 → 选类型后，页面会给出 BlockTypeID（形如 blk_xxxxxxxx），复制填进去。"
  );
}

if (!fs.existsSync(path.join(DIST, "index.html"))) {
  problems.push("dist/index.html 不存在，请先执行 npm run build。");
}

if (problems.length) {
  console.error("\n[准备上传] 还差这些才能上传：\n");
  problems.forEach((p, i) => console.error(`  ${i + 1}) ${p}`));
  console.error("");
  process.exit(1);
}

fs.copyFileSync(path.join(ROOT, "block.json"), path.join(DIST, "block.json"));

const files = fs.readdirSync(DIST);
console.log("[准备上传] OK");
console.log("  appId       :", appId);
console.log("  blockTypeID :", blockTypeId);
console.log("  待上传目录  :", DIST, `(${files.length} 个文件)`);
console.log("");
console.log("下一步：npx opdev upload ./dist");
