<#
  打印插件服务诊断（在跑服务的那台 Windows 上运行）
  用法：双击同目录的「诊断.cmd」，或
        powershell -ExecutionPolicy Bypass -File server\diagnose.ps1 -Port 8666
#>
param([int]$Port = 8666)

$ErrorActionPreference = "SilentlyContinue"
$Root = Split-Path -Parent $PSScriptRoot

function Line($text, $color) { Write-Host $text -ForegroundColor $color }
function Title($t) { Line ("`n==== " + $t + " ====") "Cyan" }

Line "打印插件服务诊断（端口 $Port）" "White"
Line ("时间：" + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) "DarkGray"

# ---------- 1. 本机地址 ----------
Title "1. 本机 IP"
$ips = @()
foreach ($a in @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue)) {
  if ($a.IPAddress -and $a.IPAddress -notlike "127.*" -and $a.IPAddress -notlike "169.254.*") { $ips += $a.IPAddress }
}
if ($ips.Count -eq 0) { Line "[X] 没有可用的局域网 IP" "Red" }
else { foreach ($ip in $ips) { Line ("    " + $ip) "White" } }

# ---------- 2. 文件是否齐 ----------
Title "2. 文件检查"
foreach ($f in @("dist\index.html", "server\serve.cjs")) {
  $p = Join-Path $Root $f
  if (Test-Path $p) { Line ("[OK] " + $f) "Green" } else { Line ("[X] 缺少 " + $f) "Red" }
}
$nodeCmd = Get-Command node.exe -ErrorAction SilentlyContinue
if ($nodeCmd) { Line ("[OK] node：" + $nodeCmd.Source) "Green" }
else { Line "[X] 找不到 node.exe（没装 Node.js，或装完没重开窗口）" "Red" }

# ---------- 3. 端口监听 ----------
Title "3. 端口监听"
$listen = @()
try { $listen = @(Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop) } catch { $listen = @() }
$boundAll = $false
if ($listen.Count -eq 0) {
  Line ("[X] 端口 " + $Port + " 没有进程监听 —— 服务根本没在跑（窗口关了？常驻任务没装？）") "Red"
} else {
  foreach ($l in $listen) {
    $procName = ""
    try { $procName = (Get-Process -Id $l.OwningProcess -ErrorAction Stop).ProcessName } catch { }
    Line ("[OK] 正在监听 " + $l.LocalAddress + ":" + $l.LocalPort + "（进程 " + $procName + "，PID " + $l.OwningProcess + "）") "Green"
    if ($l.LocalAddress -eq "0.0.0.0" -or $l.LocalAddress -eq "::") { $boundAll = $true }
  }
  if (-not $boundAll) {
    Line "[!] 只监听在 127.0.0.1 —— 本机能打开，其他电脑一定连不上" "Yellow"
    Line "    解决：关掉窗口后重新双击「启动服务.cmd」（默认监听 0.0.0.0）" "Yellow"
  }
}

# ---------- 4. 本机访问自测 ----------
Title "4. 访问自测"
try {
  $r = Invoke-WebRequest -Uri ("http://127.0.0.1:" + $Port + "/") -UseBasicParsing -TimeoutSec 5
  Line ("[OK] http://127.0.0.1:" + $Port + "/ 返回 " + $r.StatusCode) "Green"
} catch {
  Line ("[X] http://127.0.0.1:" + $Port + "/ 打不开：" + $_.Exception.Message) "Red"
}
foreach ($ip in $ips) {
  try {
    $r2 = Invoke-WebRequest -Uri ("http://" + $ip + ":" + $Port + "/") -UseBasicParsing -TimeoutSec 5
    Line ("[OK] http://" + $ip + ":" + $Port + "/ 返回 " + $r2.StatusCode) "Green"
  } catch {
    Line ("[X] http://" + $ip + ":" + $Port + "/ 打不开（用局域网地址访问不通）") "Red"
  }
}

# ---------- 5. 防火墙 ----------
Title "5. 防火墙"
foreach ($p in @(Get-NetFirewallProfile -ErrorAction SilentlyContinue)) {
  if ($p.Enabled) { Line ("[!] " + $p.Name + " 配置：已开启（未放行的入站会被拦）") "Yellow" }
  else { Line ("[OK] " + $p.Name + " 配置：已关闭") "Green" }
}
$rule = @()
try { $rule = @(Get-NetFirewallRule -DisplayName "打印插件 TCP *" -ErrorAction Stop) } catch { $rule = @() }
if ($rule.Count -eq 0) {
  Line ("[X] 没有放行 " + $Port + " 的防火墙规则 —— 别的电脑连不上，八成是这里") "Red"
  Line "    解决：右键「安装为开机自启.cmd」→「以管理员身份运行」（它同时会放行端口）" "Yellow"
} else {
  foreach ($r in $rule) {
    Line ("[OK] 放行规则：" + $r.DisplayName + "（" + $r.Enabled + " / " + $r.Direction + " / " + $r.Action + "）") "Green"
  }
}

# ---------- 6. 开机自启任务 ----------
Title "6. 开机自启任务"
$task = $null
try { $task = Get-ScheduledTask -TaskName "PrintPluginServer" -ErrorAction Stop } catch { $task = $null }
if ($task) {
  $info = Get-ScheduledTaskInfo -TaskName "PrintPluginServer" -ErrorAction SilentlyContinue
  Line ("[OK] 任务状态：" + $task.State + "，上次运行结果：" + $info.LastTaskResult) "Green"
} else {
  Line "[X] 没装常驻任务（现在这个服务是窗口里跑的，关窗口就没了）" "Red"
  Line "    解决：右键「安装为开机自启.cmd」→「以管理员身份运行」" "Yellow"
}

# ---------- 7. 结论 ----------
Title "结论"
$ok = ($listen.Count -gt 0) -and $boundAll -and ($rule.Count -gt 0)
if ($ok) {
  Line "本机监听正常（0.0.0.0）、防火墙已放行。" "Green"
  Line "如果别的电脑还是连不上，用下面这个地址试（必须在同一网段）：" "White"
  foreach ($ip in $ips) { Line ("    http://" + $ip + ":" + $Port + "/") "White" }
  Line "再不行就查：无线 AP 隔离、交换机端口隔离、跨网段没路由。" "White"
} else {
  Line "按上面标 [X] 的项逐个处理，最常见的两个原因：" "Yellow"
  Line "  1) 服务没在跑 —— 双击「启动服务.cmd」，窗口不要关" "Yellow"
  Line "  2) 防火墙没放行 —— 右键「安装为开机自启.cmd」以管理员运行" "Yellow"
}
Line ""
Line "把本窗口内容截图发我，我按结果给具体命令。" "DarkGray"
