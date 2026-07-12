# ExcelKit

Mac 版 Excel 加载项工具集 — 一键完成公式处理、格式转换等常用操作。

## 功能

| 按钮 | 功能 |
|------|------|
| **添加Round** ▼ | 为选中单元格添加 `ROUND(x, N)`，下拉菜单可选精度 0-4 位或自定义 |
| **移除Round** | 移除最外层 `ROUND(expr, N)` |
| **正负转换** | 正数变负、负数变正 |

## 开发

```bash
# 安装依赖
npm install

# 启动开发服务器 + 侧载到 Excel
npm start

# 仅构建
npm run build:dev

# 停止
npm stop
```

## 项目结构

```
ExcelKit/
├── manifest.xml              # Office Add-in 清单
├── package.json              # npm 配置
├── webpack.config.js         # webpack 打包
├── src/
│   ├── commands/
│   │   ├── commands.html     # 功能区按钮入口
│   │   └── commands.js       # 按钮逻辑
│   └── taskpane/
│       ├── taskpane.html     # 任务窗格
│       ├── taskpane.js       # 任务窗格脚本
│       └── taskpane.css      # 样式
├── assets/                   # 图标资源
└── public/
    └── dialog-precision.html # 自定义精度弹框
```

## 技术栈

- Office JavaScript API（Office.js）
- webpack + Babel
- Excel 功能区按钮（ExecuteFunction）
- 独立运行时（Browser Runtime）

## 兼容性

| 平台 | 状态 |
|------|------|
| Excel for Mac (Microsoft 365) | ✅ 已测试 |
| Excel on Windows | ⚠️ 理论上兼容，未测试 |
| Excel on the web | ⚠️ 理论上兼容，未测试 |

> 本项目基于 Office JavaScript API 标准开发，Windows 和 Web 端理论上均可运行，但目前仅在 Mac 上完成功能验证。

## 许可

MIT
