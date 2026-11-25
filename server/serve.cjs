// ===== 生产用静态服务器：托管 dist/ 给飞书「自定义插件」加载 =====
// 零第三方依赖，只用 Node 内置模块，方便在 Windows 上做成开机自启的服务。
//
// 用法：
//   node server/serve.cjs                 # 默认端口 8666
//   node server/serve.cjs 9100            # 指定端口
//   set PORT=9100 && node server/serve.cjs
//   set HOST=127.0.0.1 && node server/serve.cjs   # 只允许本机访问
//
// 环境变量：
//   PORT  监听端口（默认 8666，不用 5173/3000 这类开发默认端口，避免和别的东西撞）
//   HOST  监听地址（默认 0.0.0.0，局域网内其他电脑也能访问；只想本机用就设 127.0.0.1）
const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");

const ROOT = path.resolve(__dirname, "..", "dist");
const PORT = Number(process.argv[2] || process.env.PORT || 8666);
const HOST = process.env.HOST || "0.0.0.0";

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".map": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".ico": "image/x-icon",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
  ".ttf": "font/ttf",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".xls": "application/vnd.ms-excel",
  ".txt": "text/plain; charset=utf-8",
};

/** 日志带时间戳，方便在服务日志里看谁什么时候拉了页面 */
function log(...args) {
  console.log(`[${new Date().toISOString()}]`, ...args);
}

function send(res, code, body, headers) {
  res.writeHead(code, headers || {});
  res.end(body);
}

const server = http.createServer((req, res) => {
  // 只接受 GET / HEAD（静态资源足够）
  if (req.method !== "GET" && req.method !== "HEAD") {
    return send(res, 405, "Method Not Allowed");
  }

  let pathname;
  try {
    pathname = decodeURIComponent(new URL(req.url, "http://localhost").pathname);
  } catch {
    return send(res, 400, "Bad Request");
  }
  if (pathname === "/" || pathname === "") pathname = "/index.html";

  // 防目录穿越：拼完再确认仍在 dist 里面
  const filePath = path.join(ROOT, path.normalize(pathname).replace(/^(\.\.[/\\])+/, ""));
  if (!filePath.startsWith(ROOT)) return send(res, 403, "Forbidden");

  fs.stat(filePath, (err, st) => {
    if (err || !st.isFile()) {
      log("404", req.url);
      return send(res, 404, "Not Found", { "Content-Type": "text/plain; charset=utf-8" });
    }
    const ext = path.extname(filePath).toLowerCase();
    const type = MIME[ext] || "application/octet-stream";
    // 带 hash 的静态资源可长缓存；入口 html 必须每次回源，否则改完版本飞书还加载旧页面
    const cache = ext === ".html" ? "no-store, must-revalidate" : "public, max-age=31536000, immutable";
    const headers = {
      "Content-Type": type,
      "Content-Length": st.size,
      "Cache-Control": cache,
      // 插件是 iframe 加载的，放开跨域省事（不涉及敏感数据，页面本身是公开的）
      "Access-Control-Allow-Origin": "*",
      "X-Content-Type-Options": "nosniff",
    };
    if (req.method === "HEAD") return send(res, 200, undefined, headers);
    res.writeHead(200, headers);
    fs.createReadStream(filePath).pipe(res);
  });
});

function lanAddresses() {
  const out = [];
  for (const list of Object.values(os.networkInterfaces())) {
    for (const it of list || []) {
      if (it.family === "IPv4" && !it.internal) out.push(it.address);
    }
  }
  return out;
}

if (!fs.existsSync(path.join(ROOT, "index.html"))) {
  console.error(`[错误] 找不到 ${path.join(ROOT, "index.html")}，请先执行 npm run build`);
  process.exit(1);
}

server.on("error", (e) => {
  if (e.code === "EADDRINUSE") {
    console.error(`[错误] 端口 ${PORT} 已被占用，换一个：node server/serve.cjs 8667`);
  } else {
    console.error("[错误]", e.message);
  }
  process.exit(1);
});

server.listen(PORT, HOST, () => {
  log(`打印插件已启动：http://localhost:${PORT}/`);
  if (HOST === "0.0.0.0") {
    for (const ip of lanAddresses()) log(`局域网访问：      http://${ip}:${PORT}/`);
  }
  log(`静态目录：${ROOT}`);
});

// 作为服务运行时，被要求停止就干净退出
for (const sig of ["SIGINT", "SIGTERM", "SIGBREAK"]) {
  process.on(sig, () => {
    log(`收到 ${sig}，正在关闭…`);
    server.close(() => process.exit(0));
    setTimeout(() => process.exit(0), 3000);
  });
}
