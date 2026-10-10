# install_ph450m.ps1 — 把菲律宾 450M 交付包放进本地 workspace（默认演练，不写盘）
#
# 两种来源都支持，先解压再运行：
#   A. 成品小包  PH450M_deliverables_20261010.zip（聊天里发的，含 PPT / 讲稿 / 对照表 / 预览图）
#   B. 完整包    GitHub 分支 claude/project-file-reading-improvement-e3mec3 → Code → Download ZIP
#                （含构建脚本、素材、图标、地图，可在本地重跑构建）
#
# 用法（PowerShell）：
#   Expand-Archive .\PH450M_deliverables_20261010.zip -DestinationPath .\ph450m
#   .\install_ph450m.ps1 -Src .\ph450m            # 演练：只打印将要做的事
#   .\install_ph450m.ps1 -Src .\ph450m -Apply     # 实际执行
#   加 -Root "D:\Workspace\01_projects\2026-10_菲律宾450M方案" 可改目标目录
param(
    [Parameter(Mandatory = $true)][string]$Src,
    [string]$Root = "F:\Workspace\01_projects\2026-10_菲律宾450M方案",
    [switch]$Apply
)
$ErrorActionPreference = 'Stop'
$mode  = if ($Apply) { 'APPLY' } else { 'DRY-RUN' }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

# GitHub ZIP 解压后会多一层 "Philippines-450M-<分支名>" 目录，自动向下找含 01_交付物 的那一层
$base = Get-ChildItem -Path $Src -Directory -Recurse -Depth 2 -ErrorAction SilentlyContinue |
        Where-Object { Test-Path (Join-Path $_.FullName '01_交付物') } | Select-Object -First 1 -ExpandProperty FullName
if (Test-Path (Join-Path $Src '01_交付物')) { $base = (Resolve-Path $Src).Path }
if (-not $base) { throw "在 $Src 下没找到 01_交付物 目录，确认已解压" }
$full = Test-Path (Join-Path $base '03_生成脚本')
Write-Host "[$mode] 来源：$base（$(if ($full) {'完整包'} else {'成品小包'})）"
Write-Host "[$mode] 目标：$Root"

# 1) 完整包：整体复制到 handoff\（独立文件夹，不碰现有文件）
if ($full) {
  $handoff = Join-Path $Root 'handoff\菲律宾450M方案_完整交付包_20261010'
  if (Test-Path $handoff) { throw "已存在：$handoff —— 为避免覆盖，请先改名或删除" }
  Write-Host "[$mode] 复制完整包 -> $handoff"
  if ($Apply) { New-Item -ItemType Directory -Force -Path $handoff | Out-Null; Copy-Item -Path (Join-Path $base '*') -Destination $handoff -Recurse }
}

# 2) 成品复制到 deliverables\（同名文件先备份为 .bak_<时间戳>）
$deliv = Join-Path $Root 'deliverables'
$files = @('01_交付物\菲律宾450M方案_v8.pptx', '01_交付物\菲律宾450M方案_v8_对外版.pptx', '01_交付物\菲律宾450M讲稿_v8.md',
           '01_交付物\菲律宾450M方案_数据来源对照表.md', '01_交付物\菲律宾450M方案_数据来源对照表.docx', '04_项目文档\v8_逐页大纲.md')
foreach ($f in $files) {
  $src = Join-Path $base $f
  if (-not (Test-Path $src)) { Write-Host "[$mode] 跳过（包内没有）：$f"; continue }
  $dst = Join-Path $deliv (Split-Path $f -Leaf)
  if (Test-Path $dst) { Write-Host "[$mode] 备份 $dst -> $dst.bak_$stamp"; if ($Apply) { Copy-Item $dst "$dst.bak_$stamp" } }
  Write-Host "[$mode] 复制 $f -> $dst"
  if ($Apply) { New-Item -ItemType Directory -Force -Path $deliv | Out-Null; Copy-Item $src $dst -Force }
}

# 3) 预览图复制到 working\_preview\v8\
$pv = Join-Path $base '02_页面预览\v8'
if (Test-Path $pv) {
  $pvDst = Join-Path $Root 'working\_preview\v8'
  Write-Host "[$mode] 复制预览图 -> $pvDst"
  if ($Apply) { New-Item -ItemType Directory -Force -Path $pvDst | Out-Null; Copy-Item (Join-Path $pv '*') $pvDst -Force }
}

Write-Host ""
Write-Host "项目根目录下你本地的 brief / status / decisions / sources.md 不会被覆盖；包内更新版在 04_项目文档\，请自行对比合并。"
if (-not $Apply) { Write-Host "演练结束。确认路径无误后加 -Apply 再跑一次。" }
