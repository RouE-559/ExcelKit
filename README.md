# ExcelKit

Mac 版 Excel 加载项工具集 — 为审计师提供公式处理、格式转换等常用操作。

## 功能

| 按钮 | 功能 | 说明 |
|------|------|------|
| **添加Round** ▼ | 为选中区域添加 `ROUND(x, N)` | 下拉可选精度 0/1/2/3/4 位或自定义 (0-15) |
| **移除Round** | 移除最外层 `ROUND(expr, N)` | 公式自动还原，纯数字解除精度限制 |
| **正负转换** | 正数变负、负数变正 | 对公式包裹 `-(expr)`，对纯数字直接取反 |
| **格式转换** ▼ | 文本 ↔ 数字互转 | 见下方详细说明 |

### 格式转换

| 操作 | 效果 |
|------|------|
| **文本转数字** | 文本 `"123"` → 数值 `123`；文本日期 `"2025年11月10日"` / `"2025-06-11"` → 日期值 |
| **数字转文本** | 数值 `123.45` → 文本 `"123.45"`；日期值 → 文本日期 |

**文本转数字** 会自动识别并处理：
- 货币符号（¥/$/€/£）、千位分隔符、全角逗号
- 会计负数括号 `(123.45)` → `-123.45`
- 中文日期 `2025年6月11日`、短日期 `2025-06-11`、`2025/6/11`

> ⚠️ **已知限制**：短日期格式（`2025-06-11`、`2025/6/11`）的「数字转文本」目前无法正确转换为文本日期，会显示为裸日期序列号。受限于 Mac WKWebView 对 `range.text` / `range.numberFormat` 的支持缺陷，暂无完善修复方案。**临时规避**：先将短日期通过「文本转数字」转为长日期格式，再执行「数字转文本」。

## 安装

```bash
curl -fsSL https://raw.githubusercontent.com/RouE-559/ExcelKit/main/install_mac.sh | bash
```

脚本会自动完成：
1. 安装 npm 依赖
2. 将加载项注册到 Excel（WEF 侧载）
3. 安装 launchd 服务，**开机自动启动本地服务器**

安装后，打开 Excel → 「开始」选项卡 → 「Mac Excel 加载项工具集」即可使用。

## 卸载

```bash
curl -fsSL https://raw.githubusercontent.com/RouE-559/ExcelKit/main/uninstall_mac.sh | bash
```

停止服务器、移除加载项、清除缓存。

## 开发

```bash
# 构建
npm run build:dev

# 手动启动/停止开发服务器
npm start
npm stop

# 验证 manifest
npm run validate

# 查看服务器状态
launchctl list | grep excelkit
tail -f ~/Library/Logs/excelkit-server.log
```

## 项目结构

```
ExcelKit/
├── manifest.xml              # Office Add-in 清单（功能区按钮定义）
├── package.json              # npm 配置
├── webpack.config.js         # webpack 打包
├── com.excelkit.server.plist # launchd 自启配置
├── install_mac.sh            # 一键安装脚本
├── uninstall_mac.sh          # 卸载脚本
├── src/
│   └── commands/
│       ├── commands.html     # 功能区按钮入口（FunctionFile）
│       └── commands.js       # 全部按钮逻辑
├── assets/                   # 图标资源 (PNG 16/32/80)
└── public/
    └── dialog-precision.html # 自定义精度弹框
```

## 技术栈

- **运行时**：独立 Browser Runtime（FunctionFile），非 Shared Runtime
- **API**：Office JavaScript API（`Excel.run` / `context.sync` 批处理模式）
- **构建**：webpack 5 + Babel（Yo Office 脚手架）
- **部署**：macOS launchd 自启 + WEF 侧载
- **目标**：Excel for Mac (Microsoft 365)

## 架构决策

- **放弃 Shared Runtime**：Mac Excel 的 XML 解析器不识别 `<Runtimes>` 元素
- **FunctionFile 模式**：所有按钮通过 `Office.actions.associate()` 注册，在 `commands.js` 中实现
- **批量操作**：全部读写遵循 `load → sync → 改数组 → 赋值 → sync`，不在循环内 `sync()`
- **交集优化**：选中整列时自动缩小为「选区 ∩ 已用区域」，避免加载百万空行

## 兼容性

| 平台 | 状态 |
|------|------|
| Excel for Mac (Microsoft 365) | ✅ 开发与测试平台 |
| Excel on Windows | ⚠️ 理论兼容 |
| Excel on the web | ⚠️ 理论兼容 |
| WPS Office for Mac | ❌ 不适用（WPS 使用 JSAPI，非 Office.js） |

## 许可

MIT
