#!/bin/bash

INSTALL_DIR="$HOME/.excelkit"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
PLIST_DST="$HOME/Library/LaunchAgents/com.excelkit.server.plist"

echo "╔══════════════════════════════════════╗"
echo "║   ExcelKit 卸载脚本                  ║"
echo "╚══════════════════════════════════════╝"
echo ""

# 1. 卸载 launchd
echo "🛑 停止自启服务..."
launchctl unload "$PLIST_DST" 2>/dev/null || true
rm -f "$PLIST_DST"
echo "  ✅ launchd 已移除"

# 2. 停服务器
kill $(lsof -ti :3000) 2>/dev/null || true

# 3. 退 Excel
pkill -9 "Microsoft Excel" 2>/dev/null || true
sleep 1

# 4. 清缓存
rm -rf "$WEF_DIR"/*
rm -rf "$HOME/Library/Containers/com.Microsoft.OsfWebHost/Data"/* 2>/dev/null
echo "  ✅ 缓存已清除"

# 5. 删代码（可选）
if [ -d "$INSTALL_DIR" ]; then
    read -p "🧹 删除安装目录 ($INSTALL_DIR)? [y/N] " yn
    if [ "$yn" = "y" ] || [ "$yn" = "Y" ]; then
        rm -rf "$INSTALL_DIR"
        echo "  ✅ 已删除"
    else
        echo "  ⏭️  保留"
    fi
fi

echo ""
echo "✅ ExcelKit 已卸载"
