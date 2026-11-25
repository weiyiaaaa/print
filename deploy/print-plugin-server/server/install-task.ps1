<#
  把打印插件注册成 Windows 开机自启的后台服务（用「计划任务」实现，不需要装 NSSM 之类的东西）。

  用法（必须用【管理员身份】打开 PowerShell）：
    cd D:\workplace\print-plugin-unified
    powershell -ExecutionPolicy Bypass -File server\install-task.ps1              # 安装，端口 8666
    powershell -ExecutionPolicy Bypass -File server\install-task.ps1 -Port 9100   # 换端口
    powershell -ExecutionPolicy Bypass -File server\install-task.ps1 -Uninstall   # 卸载

  做了什么：
    1) 建一个开机自启的计划任务（SYSTEM 账户，不用登录也在跑），崩了自动重启
    2) 加一条防火墙入站规则，局域网其他电脑才能访问
    3) 立即启动，并自检能不能打开页面
#>
param(
  [int]$Port = 8666,
  [string]$TaskName = "PrintPluginServer",
  [string]$NodePath = "",
  [switch]$Uninstall,
  [switch]$NoFirewall
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot          # 项目根目录
$Serve = Join-Path $Root "server\serve.cjs"
$RuleName = "打印插件 TCP $Port"

function Fail($msg) {
  Write-Host ""
  Write-Host $msg -ForegroundColor Red
  Write-Host ""
  exit 1
}

# ---------- 管理员检查 ----------
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Fail "请右键「Windows PowerShell」-> 以管理员身份运行，再执行本脚本。"
}

# ---------- 卸载 ----------
if ($Uninstall) {
  Write-Host "正在卸载计划任务 $TaskName ..."
  Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
  Get-NetFirewallRule -DisplayName "打印插件 TCP *" -ErrorAction SilentlyContinue | Remove-NetFirewallRule -ErrorAction SilentlyContinue
  Write-Host "已卸载（任务 + 防火墙规则）。" -ForegroundColor Green
  exit 0
}

# ---------- 前置检查 ----------
if (-not (Test-Path $Serve)) { Fail "找不到 $Serve，请在项目根目录执行。" }
if (-not (Test-Path (Join-Path $Root "dist\index.html"))) {
  Fail "dist\index.html 不存在。请先在本目录执行：npm run build"
}

if (-not $NodePath) {
  # 优先用正常的 Node 安装；跳过 .workbuddy 自带的私有运行时（版本随工具升级会变，不适合做服务）
  $all = @(Get-Command node.exe -All -ErrorAction SilentlyContinue | ForEach-Object { $_.Source })
  if ($all.Count -eq 0) { Fail "没找到 node.exe。请先安装 Node.js，或用 -NodePath 指定路径。" }
  $preferred = @($all | Where-Object { $_ -notmatch '\\\.workbuddy\\' })
  if ($preferred.Count -gt 0) { $NodePath = $preferred[0] } else { $NodePath = $all[0] }
}
if (-not (Test-Path $NodePath)) { Fail "Node 路径不存在：$NodePath" }

Write-Host "项目目录：$Root"
Write-Host "Node：    $NodePath"
Write-Host "端口：    $Port"
Write-Host ""

# ---------- 重建计划任务 ----------
Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

$action = New-ScheduledTaskAction -Execute $NodePath -Argument "`"$Serve`" $Port" -WorkingDirectory $Root
$trigger = New-ScheduledTaskTrigger -AtStartup
$whoami = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet `
  -AllowStartIfOnBatteries `
  -DontStopIfGoingOnBatteries `
  -StartWhenAvailable `
  -RestartCount 999 `
  -RestartInterval (New-TimeSpan -Minutes 1) `
  -ExecutionTimeLimit ([TimeSpan]::Zero) `
  -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $whoami -Settings $settings `
  -Description "飞书多维表格打印插件静态服务（端口 $Port）" -Force | Out-Null

Start-ScheduledTask -TaskName $TaskName
Write-Host "计划任务已创建并启动：$TaskName" -ForegroundColor Green

# ---------- 防火墙 ----------
if (-not $NoFirewall) {
  Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue | Remove-NetFirewallRule -ErrorAction SilentlyContinue
  New-NetFirewallRule -DisplayName $RuleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $Port -Profile Any | Out-Null
  Write-Host "防火墙已放行 TCP $Port" -ForegroundColor Green
}

# ---------- 自检 ----------
Start-Sleep -Seconds 3
$state = (Get-ScheduledTask -TaskName $TaskName).State
try {
  $res = Invoke-WebRequest -Uri "http://127.0.0.1:$Port/" -UseBasicParsing -TimeoutSec 5
  $ok = $res.StatusCode -eq 200
} catch {
  $ok = $false
}

Write-Host ""
Write-Host "任务状态：$state"
if ($ok) {
  Write-Host "自检通过：http://127.0.0.1:$Port/ 返回 200" -ForegroundColor Green
  Write-Host ""
  Write-Host "把这个地址填到飞书多维表格「插件 -> 自定义插件 -> 新增插件」："
  Write-Host "  本机：    http://127.0.0.1:$Port/"
  $ips = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254.*" }).IPAddress
  foreach ($ip in $ips) { Write-Host "  局域网：  http://${ip}:$Port/" }
} else {
  Write-Host "自检失败：端口 $Port 打不开。可能被其他程序占用，换个端口重试：" -ForegroundColor Yellow
  Write-Host "  powershell -ExecutionPolicy Bypass -File server\install-task.ps1 -Port 8667"
}
Write-Host ""
Write-Host "常用操作："
Write-Host "  看状态：Get-ScheduledTask -TaskName $TaskName | Get-ScheduledTaskInfo"
Write-Host "  重启：  Restart-ScheduledTask -TaskName $TaskName"
Write-Host "  停止：  Stop-ScheduledTask -TaskName $TaskName"
Write-Host "  卸载：  powershell -ExecutionPolicy Bypass -File server\install-task.ps1 -Uninstall"
Write-Host ""
