# Web UI - PDF→DOCX 转换器可视化界面

> 基于 Next.js 16 + React 19 + Tailwind CSS 4 + shadcn/ui 的可视化转换界面。

## 截图功能

- 📤 **拖拽上传** PDF文件
- ⚙️ **13个转换选项** 开关 (表格/列表/公式/超链接/表单域/TOC/元数据/注释等)
- 📊 **实时统计** 页数/段落/表格/图片/列表/公式/链接/表单域/目录条目/元数据
- 👁️ **对比预览** 原PDF vs 转换后DOCX并排渲染
- 📥 **一键下载** DOCX和原PDF
- 📜 **历史记录** 自动记录转换历史
- 📋 **转换日志** 实时输出
- 📐 **进度反馈** 实时转换进度条

## 快速开始

### 前置条件

- Node.js 18+ 
- Python 3.10+
- 已安装中文字体 (运行 `../setup_fonts.sh`)

### 安装和运行

```bash
# 1. 安装Python依赖 (转换引擎)
cd ..
pip install PyMuPDF python-docx pdfplumber lxml Pillow

# 2. 安装Web UI依赖
cd web-ui
npm install --legacy-peer-deps

# 3. 启动开发服务器
npm run dev
# 访问 http://localhost:3000

# 4. 生产构建
npm run build
npm start
```

### Docker一键部署

```bash
# 在项目根目录
docker-compose up -d
# Web UI: http://localhost:3000
# API: http://localhost:8000
```

## 界面说明

### 左侧面板

| 区域 | 功能 |
|------|------|
| **上传区** | 拖拽或点击上传PDF文件 |
| **转换选项** | 13个开关控制转换行为 |
| **历史记录** | 自动记录的转换历史 |

### 右侧面板

| 区域 | 功能 |
|------|------|
| **统计卡片** | 转换结果统计 (页数/段落/表格/图片/列表/公式/链接/表单域/目录/元数据) |
| **对比预览** | 原PDF和DOCX并排渲染对比, 支持翻页 |
| **下载按钮** | 下载DOCX或原PDF |
| **转换日志** | Python输出日志 |

### 13个转换选项

| 选项 | 说明 | 默认 |
|------|------|------|
| 表格检测 | pdfplumber精确提取 | ✅ |
| 列表/编号检测 | 7种列表类型自动识别 | ✅ |
| 矢量图形还原 | DrawingML浮动形状 | ✅ |
| 数学公式检测 | OMML原生公式(可编辑) | ✅ |
| 超链接/书签 | URL/Email自动识别 | ✅ |
| 表单域识别 | 签名/复选框/文本框 | ✅ |
| 分页对齐 | 标题与下段同页/表格行不分裂 | ✅ |
| 页面填充对齐 | 强制每页内容边界(可能增页) | ❌ |
| 自动生成目录 | 基于标题级别插入TOC字段 | ✅ |
| PDF元数据还原 | 标题/作者/日期/生产者 | ✅ |
| PDF注释还原 | 高亮/批注/删除线/波浪线 | ✅ |
| 保留分页符 | 维持原PDF分页 | ✅ |
| 生成质量报告 | PDF/DOCX对比评分 | ✅ |

## 技术栈

| 技术 | 版本 | 用途 |
|------|------|------|
| Next.js | 16.0 | App Router + API Routes |
| React | 19.0 | UI框架 |
| TypeScript | 5.6 | 类型安全 |
| Tailwind CSS | 4.0 | 样式 |
| shadcn/ui | New York | UI组件库 |
| Lucide Icons | 0.460 | 图标 |
| PyMuPDF | 1.26 | PDF解析 |
| python-docx | 1.2 | DOCX生成 |

## API路由

Web UI内置以下API路由 (位于 `src/app/api/pdf2docx/`):

| 路由 | 方法 | 说明 |
|------|------|------|
| `/api/pdf2docx/convert` | POST | 上传并转换PDF |
| `/api/pdf2docx/download` | GET | 下载DOCX/PDF |
| `/api/pdf2docx/preview` | GET | 渲染页面为PNG |
| `/api/pdf2docx/history` | GET/POST/DELETE | 历史记录管理 |
| `/api/pdf2docx/list` | GET | 已转换文件列表 |
| `/api/pdf2docx/progress` | GET (SSE) | 实时进度推送 |

详见 [API.md](../API.md) 了解完整API文档。

## 目录结构

```
web-ui/
├── src/
│   ├── app/
│   │   ├── api/pdf2docx/      # API路由
│   │   │   ├── convert/       # 转换API
│   │   │   ├── download/      # 下载API
│   │   │   ├── history/       # 历史记录API
│   │   │   ├── list/          # 文件列表API
│   │   │   ├── preview/       # 页面预览API
│   │   │   └── progress/     # 进度推送API(SSE)
│   │   ├── globals.css       # 全局样式
│   │   ├── layout.tsx         # 布局
│   │   └── page.tsx           # 主页面
│   ├── components/ui/         # shadcn/ui组件
│   ├── hooks/                 # React Hooks
│   └── lib/                   # 工具函数
├── next.config.ts             # Next.js配置
├── tailwind.config.ts         # Tailwind配置
├── tsconfig.json              # TypeScript配置
└── package.json               # 依赖
```
