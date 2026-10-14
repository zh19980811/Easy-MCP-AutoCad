#!/bin/bash
# 无头模式启动 FreeCAD bridge（供测试或服务器环境用）
# GUI 用法见 使用说明.md
cd "$(dirname "$0")"
trap 'pkill -P $$ 2>/dev/null; exit 0' TERM INT EXIT
{ printf "exec(open('%s/bridge.py').read())\n" "$PWD"; sleep infinity; } | ./squashfs-root/usr/bin/freecadcmd
