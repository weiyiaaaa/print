<#
  一键修网络：放行防火墙 + 自检（在跑服务的那台 Windows 上以管理员身份运行）

  用法：
    右键「修复网络.cmd」→「以管理员身份运行」
    或 powershell -ExecutionPolicy Bypass -File server\fix-net.ps1 -Port 8666

  做三件事：
    1) 删掉 Windows 里针对 node.exe 的「阻止」规则（这类规则会让端口放行失效）
    2) 添加 node.exe 与端口两条「允许入站」规则
    3) 自测 127.0.0.1 与局域网地址，并打印该填进飞书的地址
#>
param([int]$Port = 8666)

$ErrorActionPreference = "SilentlyContinue"
$Root = Split-Path -Parent $PSScriptRoot

function Line($t, $c) { Write-Host $t -ForegroundColor $c }
function Title($t) { Line ("`n==== " + $t + " ====") "Cyan" }

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Line "需要管理员权限：请右键「修复网络.cmd」→「以管理员身份运行」。" "Red"
  exit 1
}

# ---------- 1. 当前监听 ----------
Title "1. 当前监听"
$listen = @()
try { $listen = @(Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop) } catch { }
if ($listen.Count -eq 0) {
  Line ("[!] 端口 " + $Port + " 现在没有进程监听 —— 先把服务跑起来（双击「启动服务.cmd」），再跑本脚本") "Yellow"
} else {
  foreach ($l in $listen) { Line ("[OK] " + $l.LocalAddress + ":" + $l.LocalPort + "  PID " + $l.OwningProcess) "Green" }
}

# ---------- 2. 清理 node.exe 的阻止规则 ----------
Title "2. 清理阻止规则"
$killed = 0
foreach ($r in @(Get-NetFirewallRule -Direction Inbound -Action Block -Enabled True)) {
  $app = $null
  try { $app = $r | Get-NetFirewallApplicationFilter -ErrorAction Stop } catch { }
  if ($app -and $app.Program -and ($app.Program -like "*node.exe")) {
    Line ("[-] 发现阻止 node.exe 的规则：" + $r.DisplayName + "，已删除") "Yellow"
    Remove-NetFirewallRule -Name $r.Name
    $killed++
  }
}
if ($killed -eq 0) { Line "[OK] 没有针对 node.exe 的阻止规则" "Green" }

# ---------- 3. 添加放行规则 ----------
Title "3. 添加放行规则"
Get-NetFirewallRule -DisplayName "打印插件 *" | Remove-NetFirewallRule
New-NetFirewallRule -DisplayName ("打印插件 TCP " + $Port) -Direction Inbound -Action Allow -Protocol TCP -LocalPort $Port -Profile Any -Enabled True | Out-Null
Line ("[OK] 已放行 TCP " + $Port + "（所有网络配置）") "Green"

$nodeCmd = Get-Command node.exe
if ($nodeCmd) {
  New-NetFirewallRule -DisplayName "打印插件 node.exe" -Direction Inbound -Action Allow -Program $nodeCmd.Source -Profile Any -Enabled True | Out-Null
  Line ("[OK] 已放行 node.exe：" + $nodeCmd.Source) "Green"
} else {
  Line "[!] 没找到 node.exe，跳过程序放行（只做了端口放行）" "Yellow"
}

if ($listen.Count -gt 0) {
  foreach ($l in $listen) {
    $procId = $l.OwningProcess
    $exe = (Get-Process -Id $procId).Path
    if ($exe) {
      New-NetFirewallRule -DisplayName "打印插件 进程放行" -Direction Inbound -Action Allow -Program $exe -Profile Any -Enabled True | Out-Null
      Line ("[OK] 已放行当前服务进程：" + $exe) "Green"
    }
  }
}

# ---------- 4. 自测 ----------
Title "4. 自测"
$ips = @()
foreach ($a in @(Get-NetIPAddress -AddressFamily IPv4)) {
  if ($a.IPAddress -and $a.IPAddress -notlike "127.*" -and $a.IPAddress -notlike "169.254.*") { $ips += $a.IPAddress }
}
try {
  $r = Invoke-WebRequest -Uri ("http://127.0.0.1:" + $Port + "/") -UseBasicParsing -TimeoutSec 5
  Line ("[OK] http://127.0.0.1:" + $Port + "/ 返回 " + $r.StatusCode) "Green"
} catch {
  Line ("[X] http://127.0.0.1:" + $Port + "/ 打不开（服务没跑？）") "Red"
}
foreach ($ip in $ips) {
  try {
    $r2 = Invoke-WebRequest -Uri ("http://" + $ip + ":" + $Port + "/") -UseBasicParsing -TimeoutSec 5
    Line ("[OK] http://" + $ip + ":" + $Port + "/ 返回 " + $r2.StatusCode) "Green"
  } catch {
    Line ("[X] http://" + $ip + ":" + $Port + "/ 本机不通") "Red"
  }
}

# ---------- 5. 填什么地址 ----------
Title "5. 飞书里该填的地址（必须带端口）"
foreach ($ip in $ips) { Line ("    http://" + $ip + ":" + $Port + "/") "White" }
Line ("    只写 IP 不带 :" + $Port + " 的话，浏览器会去连 80 端口，必然超时") "Yellow"
Line ""
Line "改完防火墙后，从【另一台电脑】的浏览器打开上面地址验证；能打开再到飞书里填。" "White"
