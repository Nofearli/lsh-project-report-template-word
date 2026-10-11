#!/usr/bin/env bash
# 将本技能安装到 ZCode 用户级技能目录（~/.zcode/skills）
set -euo pipefail

src="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
dest="$HOME/.zcode/skills/project-report-template-word"

mkdir -p "$(dirname "$dest")"
rm -rf "$dest"
mkdir -p "$dest"
tar -C "$src" --exclude=.git --exclude=__pycache__ -cf - . | tar -C "$dest" -xf -

echo "Installed to $dest"
echo "在任意工作区重新打开 ZCode 即可自动发现该技能。"
