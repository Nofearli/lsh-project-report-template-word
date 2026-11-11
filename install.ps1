# 将本技能安装到 ZCode 用户级技能目录（%USERPROFILE%\.zcode\skills）
$ErrorActionPreference = "Stop"

$src  = $PSScriptRoot
$dest = Join-Path $env:USERPROFILE ".zcode\skills\lsh-project-report-template-word"

New-Item -ItemType Directory -Force -Path (Split-Path $dest) | Out-Null
if (Test-Path $dest) { Remove-Item -Recurse -Force $dest }
New-Item -ItemType Directory -Force -Path $dest | Out-Null
robocopy $src $dest /E /XD .git __pycache__ /NFL /NDL /NJH | Out-Null

Write-Host "Installed to $dest"
Write-Host "在任意工作区重新打开 ZCode 即可自动发现该技能。"
