# PDF→DOCX 企业级 1:1 复刻转换器

> **自研内核，不依赖 pdf2docx。** 基于 PyMuPDF + python-docx + lxml 的多模块协同架构。
> 
> 支持 **CLI命令行** + **Web UI可视化界面** + **REST API程序调用** 三种使用方式。

## 📦 三种使用方式

### 1. CLI 命令行 (单文件/批量)

```bash
# 单文件
python3 main.py -i input.pdf -o output.docx

# 批量并行
python3 batch.py -i /input/dir -o /output/dir -w 4
```

### 2. Web UI 可视化界面

```bash
cd web-ui && npm install && npm run dev
# 访问 http://localhost:3000
```

详见 [Web UI README](web-ui/README.md)

### 3. REST API 程序调用

```bash
python3 api_server.py --port 8000
# POST /api/convert 上传转换
# GET /api/download/:id 下载
```

详见 [API文档](API.md)

## 🐳 Docker一键部署

```bash
docker-compose up -d
# Web UI: http://localhost:3000
# API: http://localhost:8000
```

## ⚠️ 重要说明

### 字体依赖
本工具**不附带任何字体文件**。运行前请先安装字体：

```bash
chmod +x setup_fonts.sh && ./setup_fonts.sh
```

详见 [FONTS.md](FONTS.md) 了解字体来源和授权。

### 已知限制
- **页数差异**: DOCX流式排版导致页数略多于PDF（97→106页）
- **字体回退**: 未安装字体会被系统回退字体替代
- **下划线**: 基于矢量线段几何匹配，字体自带字形可能漏判
- **复杂矢量图**: 曲线/弧形用矩形包围盒近似
- **旋转/竖排文本**: 不支持
- **加密PDF**: 需提供密码（`--password`）

## 核心能力

| 能力 | 保真度 | 说明 |
|------|--------|------|
| **字体** | ★★★★☆ | CJK+西文双字体映射，不再强制Times New Roman |
| **字号** | ★★★★★ | 精确磅值 w:sz + w:szCs |
| **颜色** | ★★★★★ | 位移提取RGB，验证正确 |
| **加粗/斜体** | ★★★★★ | flags位+字体名双重判定 |
| **下划线** | ★★★☆☆ | 9种样式，矢量线段几何匹配 |
| **页眉** | ★★★★☆ | 自动检测对齐(右对齐)+分隔线 |
| **页脚** | ★★★★☆ | PAGE域字段+居中对齐 |
| **表格** | ★★★★☆ | pdfplumber+边框+合并 |
| **图片** | ★★★★☆ | xref提取+浮动定位 |
| **印章** | ★★★★☆ | 签名域bbox浮动定位 |
| **列表编号** | ★★★☆☆ | 7种类型+DOCX numbering |
| **数学公式** | ★★★☆☆ | OMML原生公式(可编辑) |
| **超链接** | ★★★☆☆ | w:hyperlink+URL识别 |
| **表单域** | ★★★☆☆ | 签名/复选框/文本框 |
| **TOC目录** | ★★★☆☆ | TOC字段+书签 |
| **元数据** | ★★★☆☆ | core_properties+日期解析 |
| **注释** | ★★☆☆☆ | 高亮/批注/删除线 |

## 项目结构

```
pdf2word-webui-cli/
├── main.py                  # CLI入口(单文件)
├── batch.py                 # 批量处理(支持并行)
├── api_server.py            # 独立REST API服务
├── setup_fonts.sh           # 字体安装脚本
├── Dockerfile               # Docker构建文件
├── docker-compose.yml       # Docker Compose
├── src/                     # Python转换引擎(16模块)
│   ├── formatter.py         # 格式映射(字体/颜色/字号)
│   ├── extractor.py         # PDF提取(含加密检测)
│   ├── analyzer.py          # 布局分析(页眉页脚对齐)
│   ├── table_extractor.py   # 表格提取
│   ├── docx_builder.py      # DOCX构建
│   ├── converter.py         # 主编排
│   ├── reporter.py          # 质量报告
│   ├── shape_builder.py     # 矢量图形
│   ├── list_detector.py     # 列表检测
│   ├── formula_detector.py  # 数学公式(OMML)
│   ├── link_handler.py      # 超链接
│   ├── form_handler.py      # 表单域
│   ├── toc_builder.py       # TOC目录
│   ├── metadata_handler.py  # 元数据
│   ├── annotation_handler.py # 注释
│   └── page_aligner.py      # 分页对齐
├── web-ui/                  # Web UI (Next.js)
│   ├── src/app/             # 页面+API路由
│   ├── src/components/ui/   # shadcn/ui组件
│   ├── package.json
│   └── README.md            # Web UI文档
├── README.md                # 本文件
├── API.md                   # API接口文档
├── FONTS.md                 # 字体安装指南
├── CODE_REVIEW.md           # 代码审查回应
├── LICENSE                  # MIT
└── .gitignore
```

## 安装

### 完整安装

```bash
# 1. Python依赖
pip install PyMuPDF python-docx pdfplumber lxml Pillow aiohttp

# 2. 中文字体
chmod +x setup_fonts.sh && ./setup_fonts.sh

# 3. Web UI依赖
cd web-ui && npm install --legacy-peer-deps && cd ..
```

### Docker安装

```bash
docker-compose up -d
```

## 字体来源

| 来源 | 说明 |
|------|------|
| [StellarCN/scp_zh](https://github.com/StellarCN/scp_zh) | SimSun, SimHei |
| [zhyounger/FontsFromWindows](https://github.com/zhyounger/FontsFromWindows) | Windows字体集合 |
| [DoveOutland字体库](https://github.com/DoveOutland/Common-Chinese-office-fonts-font-library-) | 中文办公字体 |
| [晋城市财政局字体包](http://czj.jcgov.gov.cn/ggfw/xzzx/202506/P020250624600205358001.zip) | 政府公文标准字体 |

> **免责声明**: 字体版权归各自所有者。本工具不附带字体文件，仅提供下载链接。商业使用请确认授权许可。

## License

MIT
