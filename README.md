# PDF→DOCX 企业级 1:1 复刻转换器

> **自研内核，不依赖 pdf2docx。** 基于 PyMuPDF + python-docx + lxml 的多模块协同架构。

## ⚠️ 重要说明

### 字体依赖
本工具**不附带任何字体文件**。转换后的 DOCX 保留了 PDF 中的原始字体名（如仿宋、黑体），系统需安装对应字体才能正确渲染。

**请先运行字体安装脚本：**
```bash
chmod +x setup_fonts.sh && ./setup_fonts.sh
```

详见 [FONTS.md](FONTS.md) 了解字体下载来源和授权说明。

### 已知限制（诚实声明）
- **页数差异**: DOCX 流式排版导致页数略多于 PDF（97页PDF→106页DOCX），非严格1:1
- **字体回退**: 未安装的字体会被系统回退字体替代，视觉上可能不同
- **下划线**: 基于矢量线段几何匹配，字体自带下划线字形可能漏判
- **复杂矢量图**: 曲线/弧形用矩形包围盒近似
- **旋转/竖排文本**: 不支持
- **RTL文本**: 阿拉伯/希伯来等从右到左文本阅读顺序可能错误
- **加密PDF**: 需提供密码（已添加检测和报错）

## 核心能力

| 能力 | 实现方式 | 保真度 | 说明 |
|------|----------|--------|------|
| **字体** | w:rFonts + w:eastAsia 双字体 | ★★★★☆ | CJK和西文使用相同字体名(已修复) |
| **字号** | w:sz + w:szCs | ★★★★★ | 精确磅值 |
| **颜色** | 位移提取RGB → w:color | ★★★★★ | 正确转换(已验证) |
| **加粗/斜体** | flags位 + 字体名双重判定 | ★★★★★ | |
| **下划线** | 矢量线段几何匹配 + w:u | ★★★☆☆ | 9种样式，字体自带字形可能漏判 |
| **页眉** | 自动检测对齐 + 分隔线 | ★★★★☆ | 右对齐/居中自动检测 |
| **页脚** | PAGE域字段自动编号 | ★★★★☆ | 居中对齐自动检测 |
| **页眉分隔线** | 段落底边框 w:pBdr | ★★★★☆ | 颜色/粗细保留 |
| **表格** | pdfplumber + w:tblBorders | ★★★★☆ | 含合并单元格 |
| **图片** | xref提取 + 区域裁剪 + wp:anchor | ★★★★☆ | 印章浮动定位 |
| **列表编号** | 7种类型 + DOCX numbering | ★★★☆☆ | |
| **数学公式** | OMML原生公式 | ★★★☆☆ | 可编辑 |
| **超链接** | w:hyperlink + URL识别 | ★★★☆☆ | |
| **表单域** | 签名/复选框/文本框 | ★★★☆☆ | |
| **TOC目录** | w:fldChar TOC + 书签 | ★★★☆☆ | Word中"更新域"生成 |
| **元数据** | core_properties | ★★★☆☆ | |
| **注释** | 高亮/批注/删除线 | ★★☆☆☆ | 基础支持 |

## 技术架构

```
pdf2word-webui-cli/
├── main.py                  # CLI入口
├── batch.py                 # 批量处理(支持并行)
├── setup_fonts.sh           # 字体安装脚本
├── src/
│   ├── formatter.py         # 格式映射(字体/颜色/字号)
│   ├── extractor.py         # PDF提取(含加密检测)
│   ├── analyzer.py          # 布局分析(页眉页脚对齐自动检测)
│   ├── table_extractor.py   # 表格提取
│   ├── docx_builder.py       # DOCX构建
│   ├── converter.py         # 主编排
│   ├── reporter.py          # 质量报告
│   ├── shape_builder.py     # 矢量图形
│   ├── list_detector.py     # 列表检测
│   ├── formula_detector.py  # 数学公式(OMML)
│   ├── link_handler.py      # 超链接
│   ├── form_handler.py      # 表单域
│   ├── toc_builder.py        # TOC目录
│   ├── metadata_handler.py  # 元数据
│   ├── annotation_handler.py # 注释
│   └── page_aligner.py      # 分页对齐
├── FONTS.md                 # 字体安装指南
├── CODE_REVIEW.md           # 代码审查回应
├── LICENSE                  # MIT
└── .gitignore
```

## 安装

```bash
# 1. 安装Python依赖
pip install PyMuPDF python-docx pdfplumber lxml Pillow

# 2. 安装中文字体(必需)
chmod +x setup_fonts.sh && ./setup_fonts.sh
```

## 使用

```bash
# 单文件转换
python3 main.py -i input.pdf -o output.docx

# 指定页码
python3 main.py -i input.pdf -o output.docx --pages 1-10

# 质量报告
python3 main.py -i input.pdf -o output.docx --report

# 批量处理(并行)
python3 batch.py -i /input/dir -o /output/dir -w 4

# 加密PDF
python3 main.py -i encrypted.pdf -o output.docx --password YOUR_PASSWORD
```

## 字体来源

| 来源 | 说明 |
|------|------|
| [StellarCN/scp_zh](https://github.com/StellarCN/scp_zh) | SimSun, SimHei |
| [zhyounger/FontsFromWindows](https://github.com/zhyounger/FontsFromWindows) | Windows字体集合 |
| [DoveOutland字体库](https://github.com/DoveOutland/Common-Chinese-office-fonts-font-library-) | 中文办公字体 |
| [晋城市财政局字体包](http://czj.jcgov.gov.cn/ggfw/xzzx/202506/P020250624600205358001.zip) | 政府公文标准字体 |

> **免责声明**: 字体版权归各自所有者。本工具不附带字体文件，仅提供下载链接参考。商业使用请确认授权许可。

## License

MIT
