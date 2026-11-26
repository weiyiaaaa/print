// 验证生产构建与静态服务器都只通过 /print/ 提供打印插件。
const assert = require("assert");
const fs = require("fs");
const http = require("http");
const net = require("net");
const path = require("path");
const { spawn } = require("child_process");

const ROOT = path.resolve(__dirname, "..");

function getFreePort() {
  return new Promise((resolve, reject) => {
    const socket = net.createServer();
    socket.once("error", reject);
    socket.listen(0, "127.0.0.1", () => {
      const { port } = socket.address();
      socket.close(() => resolve(port));
    });
  });
}

function request(port, pathname) {
  return new Promise((resolve, reject) => {
    const req = http.get({ host: "127.0.0.1", port, path: pathname }, (res) => {
      let body = "";
      res.setEncoding("utf8");
      res.on("data", (chunk) => { body += chunk; });
      res.on("end", () => resolve({ status: res.statusCode, headers: res.headers, body }));
    });
    req.setTimeout(3000, () => req.destroy(new Error("请求超时")));
    req.on("error", reject);
  });
}

async function waitUntilReady(port, child) {
  for (let i = 0; i < 50; i += 1) {
    if (child.exitCode !== null) throw new Error(`静态服务器提前退出：${child.exitCode}`);
    try {
      return await request(port, "/print/");
    } catch {
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
  }
  throw new Error("静态服务器启动超时");
}

(async () => {
  const html = fs.readFileSync(path.join(ROOT, "dist", "index.html"), "utf8");
  const script = html.match(/<script[^>]+src=(?:"([^"]+)"|([^\s>]+))/);
  const scriptPath = script && (script[1] || script[2]);
  assert(scriptPath && scriptPath.startsWith("/print/"), "构建后的 JS 地址必须带 /print/ 前缀");
  console.log("PASS 构建资源使用 /print/ 前缀");

  const port = await getFreePort();
  const child = spawn(process.execPath, [path.join(ROOT, "server", "serve.cjs"), String(port)], {
    env: { ...process.env, HOST: "127.0.0.1" },
    stdio: ["ignore", "ignore", "inherit"],
  });

  try {
    const page = await waitUntilReady(port, child);
    assert.strictEqual(page.status, 200);
    assert(page.body.includes(scriptPath), "打印入口应引用构建后的 JS");
    console.log("PASS /print/ 返回打印插件页面");

    const asset = await request(port, scriptPath);
    assert.strictEqual(asset.status, 200);
    assert(asset.body.length > 1000, "JS 资源不应为空");
    console.log("PASS /print/ 下的 JS 资源可访问");

    const redirect = await request(port, "/print?from=test");
    assert.strictEqual(redirect.status, 308);
    assert.strictEqual(redirect.headers.location, "/print/?from=test");
    console.log("PASS /print 自动补斜杠并保留查询参数");

    assert.strictEqual((await request(port, "/")).status, 404);
    assert.strictEqual((await request(port, "/factory/")).status, 404);
    console.log("PASS 根路径和其他插件路径不加载打印页面");
  } finally {
    child.kill();
  }
})().catch((error) => {
  console.error("FAIL", error.message);
  process.exitCode = 1;
});
