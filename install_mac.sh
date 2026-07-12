#!/bin/bash
set -e

REPO="https://github.com/RouE-559/ExcelKit.git"
INSTALL_DIR="$HOME/Library/Application Support/ExcelKit"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
PLIST_DST="$HOME/Library/LaunchAgents/com.excelkit.server.plist"
LOG="$HOME/Library/Logs/excelkit-server.log"

echo "╔══════════════════════════════════════╗"
echo "║   ExcelKit 安装脚本                  ║"
echo "╚══════════════════════════════════════╝"
echo ""

# 1. 检查环境
echo "━━━ 1/5 环境 ━━━"
command -v node &>/dev/null || { echo "❌ 请先安装 Node.js: brew install node"; exit 1; }
command -v git &>/dev/null || { echo "❌ 请先安装 Git: xcode-select --install"; exit 1; }
echo "  ✅ Node.js $(node -v)"

# 2. 下载/更新代码
echo "━━━ 2/5 代码 ━━━"
if [ -d "$INSTALL_DIR/.git" ]; then
    cd "$INSTALL_DIR" && git pull --ff-only 2>/dev/null
    echo "  ✅ 已更新"
else
    rm -rf "$INSTALL_DIR"
    git clone "$REPO" "$INSTALL_DIR"
    echo "  ✅ 已下载"
fi

# 3. 安装依赖
echo "━━━ 3/6 依赖 ━━━"
cd "$INSTALL_DIR"
npm install --silent 2>/dev/null
echo "  ✅ node_modules 就绪"

# 4. 安装 HTTPS 证书（新机器必需）
echo "━━━ 4/6 证书 ━━━"
npx office-addin-dev-certs install 2>/dev/null || true
echo "  ✅ HTTPS 证书就绪"

# 5. 验证构建
echo "━━━ 5/6 构建 ━━━"
npx webpack --mode development 2>/dev/null
echo "  ✅ dist 已生成"

# 6. 停旧 + 清缓存 + 侧载
echo "━━━ 6/8 注册 ━━━"
launchctl unload "$PLIST_DST" 2>/dev/null || true
pkill -9 "Microsoft Excel" 2>/dev/null || true
kill $(lsof -ti :3000) 2>/dev/null || true
sleep 2
# 彻底清除各级缓存
rm -rf "$WEF_DIR"
rm -rf "$HOME/Library/Containers/com.Microsoft.OsfWebHost/Data" 2>/dev/null
rm -rf "$HOME/Library/Containers/com.microsoft.Excel/Data/Library/Caches" 2>/dev/null
rm -rf "$HOME/Library/Caches/com.microsoft.Excel" 2>/dev/null
rm -rf "$HOME/Library/WebKit/com.microsoft.Excel" 2>/dev/null
# 杀掉残留的 Office 进程
pkill -9 "OsfWebHost" 2>/dev/null || true
mkdir -p "$WEF_DIR"
cp "$INSTALL_DIR/manifest.xml" "$WEF_DIR/"
echo "  ✅ manifest.xml → WEF"

# 7. 安装 launchd 自启
echo "━━━ 7/8 自启 ━━━"
# 修正 node 路径
NODE_PATH=$(which node)
sed -i '' "s|/opt/homebrew/bin/node|$NODE_PATH|g" "$INSTALL_DIR/com.excelkit.server.plist" 2>/dev/null || true
sed -i '' "s|/usr/local/bin/node|$NODE_PATH|g" "$INSTALL_DIR/com.excelkit.server.plist" 2>/dev/null || true
# 修正安装路径
sed -i '' "s|/Users/lly/trae-project/ExcelKit|$INSTALL_DIR|g" "$INSTALL_DIR/com.excelkit.server.plist"

cp "$INSTALL_DIR/com.excelkit.server.plist" "$PLIST_DST"
launchctl load "$PLIST_DST"
sleep 4

if lsof -i :3000 2>/dev/null | grep -q LISTEN; then
    echo "  ✅ 服务器已启动 (localhost:3000)"
else
    echo "  ⚠️  检查日志: $LOG"
fi

# 8. 验证
echo "━━━ 8/8 验证 ━━━"
echo "  ⏳ 等待服务器启动..."

# 轮询等待最多 30 秒
for i in $(seq 1 15); do
  if curl -sk https://localhost:3000/commands.js -o /dev/null 2>/dev/null; then
    break
  fi
  sleep 2
done

if curl -sk https://localhost:3000/commands.js -o /dev/null 2>/dev/null; then
  echo "  ✅ 服务器在线"
  ICONS_OK=0
  for icon in addRound_16 removeRound_16 toggleSign_16 convertFormat_16; do
    STATUS=$(curl -sk -o /dev/null -w "%{http_code}" "https://localhost:3000/assets/${icon}.png" 2>/dev/null)
    if [ "$STATUS" = "200" ]; then ICONS_OK=$((ICONS_OK+1)); fi
  done
  if [ $ICONS_OK -eq 4 ]; then
    echo "  ✅ 4/4 图标正常"
  else
    echo "  ⚠️  $ICONS_OK/4 图标 404"
    echo "  修复: cd '$INSTALL_DIR' && npm run build:dev"
  fi
else
  echo "  ❌ 30秒后仍未响应"
  echo "  检查日志: cat '$LOG'"
  echo "  手动修复:"
  echo "    cd '$INSTALL_DIR' && npm run build:dev"
  echo "    cd '$INSTALL_DIR' && npm run dev-server &"
fi

echo ""
echo "✅ ExcelKit 安装完成"
echo "   打开 Excel → 开始选项卡 → Mac Excel 加载项工具集"
echo ""
echo "   卸载命令:"
echo "   curl -fsSL https://raw.githubusercontent.com/RouE-559/ExcelKit/main/uninstall_mac.sh | bash"
