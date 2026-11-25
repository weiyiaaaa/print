<#
  前台启动（给「启动服务.cmd」调用）
  powershell -ExecutionPolicy Bypass -File server\start.ps1 [端口]
#>
param([int]$Port = 8666)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   飞书打印插件 · 本地服务" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path (Join-Path $Root "dist\index.html"))) {
  Write-Host "[错误] 找不到 dist\index.html" -ForegroundColor Red
  Write-Host "请确认 dist 文件夹与本文件在同一目录下。" -ForegroundColor Yellow
  exit 1
}

$nodeCmd = Get-Command node.exe -ErrorAction SilentlyContinue
if (-not $nodeCmd) {
  Write-Host "[错误] 这台电脑没装 Node.js，或者装完没重开窗口。" -ForegroundColor Red
  Write-Host "请先安装 Node.js LTS：https://nodejs.org/zh-cn" -ForegroundColor Yellow
  Write-Host "一路下一步即可，装完关掉本窗口再双击运行。" -ForegroundColor Yellow
  exit 1
}

Write-Host ("端口      : " + $Port)
Write-Host ("本机访问  : http://127.0.0.1:" + $Port + "/")
Write-Host ("别的电脑  : http://<本机IP>:" + $Port + "/   （地址必须带端口）")
Write-Host "停止服务  : 按 Ctrl+C，或直接关掉本窗口"
Write-Host ""

& node (Join-Path $Root "server\serve.cjs") $Port
Write-Host ""
Write-Host "服务已停止。" -ForegroundColor DarkGray
