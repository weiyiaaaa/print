// 一键打包：build -> 同步 dist 到部署目录 -> 生成 print-plugin-server.zip
// 用法：npm run package
// 产物：deploy/print-plugin-server.zip（拷到 Windows 服务器解压即用，无需 Node / npm）
const { execFileSync } = require("child_process");
const path = require("path");
const fs = require("fs");

const ROOT = path.resolve(__dirname, "..");
const DIST = path.join(ROOT, "dist");
const PKG_DIR = path.join(ROOT, "deploy", "print-plugin-server");
const PKG_DIST = path.join(PKG_DIR, "dist");
const ZIP = path.join(ROOT, "deploy", "print-plugin-server.zip");

function log(msg) { console.log("[打包] " + msg); }

function copyDir(src, dst) {
  fs.mkdirSync(dst, { recursive: true });
  for (const name of fs.readdirSync(src)) {
    const s = path.join(src, name);
    const d = path.join(dst, name);
    const stat = fs.statSync(s);
    if (stat.isDirectory()) copyDir(s, d);
    else fs.copyFileSync(s, d);
  }
}

function main() {
  if (!fs.existsSync(DIST)) throw new Error("找不到 dist，请先执行 npm run build");
  if (!fs.existsSync(PKG_DIR)) throw new Error("找不到部署目录 " + PKG_DIR);

  log("1/3 清空并同步 dist ...");
  fs.rmSync(PKG_DIST, { recursive: true, force: true });
  copyDir(DIST, PKG_DIST);

  log("2/3 生成 zip（Compress-Archive）...");
  if (fs.existsSync(ZIP)) fs.rmSync(ZIP, { force: true });
  const ps = [
    "$ProgressPreference='SilentlyContinue';",
    `Compress-Archive -Path '${path.join(PKG_DIR, "*")}' -DestinationPath '${ZIP}' -CompressionLevel Optimal -Force`,
  ].join(" ");
  execFileSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", ps], { cwd: ROOT, stdio: "inherit" });

  log("3/3 完成");
  const kb = (fs.statSync(ZIP).size / 1024).toFixed(1);
  console.log(`\n产物：${ZIP}\n大小：${kb} KB`);
  console.log("拷贝这个 zip 到服务器，解压后按《使用说明.txt》操作即可。");
}

try { main(); } catch (e) { console.error("[打包失败]", e.message); process.exit(1); }
