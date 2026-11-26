// 一键打包：build -> 同步 dist 与服务器源码 -> 生成 print-plugin-server.zip
// 用法：npm run package
// 产物：deploy/print-plugin-server.zip（拷到 Windows 服务器解压即用，无需 Node / npm）
const { execFileSync } = require("child_process");
const path = require("path");
const fs = require("fs");

const ROOT = path.resolve(__dirname, "..");
const DIST = path.join(ROOT, "dist");
const SERVER = path.join(ROOT, "server");
const PKG_DIR = path.join(ROOT, "deploy", "print-plugin-server");
const PKG_DIST = path.join(PKG_DIR, "dist");
const PKG_SERVER = path.join(PKG_DIR, "server");
const ZIP = path.join(ROOT, "deploy", "print-plugin-server.zip");

function log(msg) { console.log("[打包] " + msg); }

function main() {
  if (!fs.existsSync(DIST)) throw new Error("找不到 dist，请先执行 npm run build");
  fs.mkdirSync(PKG_DIR, { recursive: true });

  log("1/3 同步页面与服务器源码 ...");
  fs.rmSync(PKG_DIST, { recursive: true, force: true });
  fs.cpSync(DIST, PKG_DIST, { recursive: true });
  fs.rmSync(PKG_SERVER, { recursive: true, force: true });
  fs.cpSync(SERVER, PKG_SERVER, { recursive: true });

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
