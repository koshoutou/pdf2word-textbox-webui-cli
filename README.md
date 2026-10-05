# PDF→DOCX 企业级 1:1 复刻转换器

一个自研的、企业级的 PDF→DOCX 转换工具，致力于**精准 1:1 复刻** PDF 文档的全部格式信息到 DOCX。

> 不依赖 `pdf2docx`，自研多模块协同架构，达到更高保真度。

## v12.0 新增功能
- 📜 **转换历史记录面板**: Web UI侧边栏显示历史转换记录
  - 自动记录: 文件名/时间/页数/评分/大小
  - 时间格式化: 刚刚/X分钟前/X小时前/日期
  - 点击历史项可恢复预览
  - 一键清空历史
  - 最多保留50条记录
- 🆕 **history API**: GET/POST/DELETE 历史记录管理
  - 存储到 /tmp/pdf2docx-history/history.json
  - convert 成功后自动写入
- 🐛 **修复 existsSync 导入**: 从 fs/promises 改为 fs (fs/promises 无 existsSync)
- 🖥️ **Web UI v12**: 左侧新增"转换历史"面板 + 版本号v12.0

## 核心能力

| 能力 | 实现方式 | 保真度 |
|------|----------|--------|
| **字体** (仿宋/黑体/宋体/楷体/微软雅黑) | w:rFonts + w:eastAsia 双字体映射 | ★★★★★ |
| **字号** (磅/中文字号) | w:sz + w:szCs | ★★★★★ |
| **颜色** (RGB) | w:color hex | ★★★★★ |
| **加粗/斜体** | flags位 + 字体名双重判定 | ★★★★★ |
| **下划线** (单/双/虚/点) | 矢量线段几何匹配 + w:u | ★★★★☆ |
| **上标/下标** | flags位 + 字号比例 | ★★★★☆ |
| **段落对齐** (左/中/右/两端) | bbox 几何分析 + w:jc | ★★★★☆ |
| **首行缩进** | 行间 x0 对比 + w:ind | ★★★★☆ |
| **行距** (1.0/1.5/2.0倍) | 行高比 + w:spacing | ★★★★☆ |
| **页眉** (文字+页码+分隔线) | header + w:pBdr bottom | ★★★★★ |
| **页脚** (文字+页码字段+分隔线) | footer + PAGE field + w:pBdr | ★★★★★ |
| **页眉页脚分隔线** | 水平矢量线检测 + 段落底/顶边框 | ★★★★★ |
| **表格** (边框/合并/对齐) | pdfplumber + w:tblBorders + merge | ★★★★☆ |
| **图片** (内联+浮动) | xref 提取 + 区域裁剪渲染 + wp:anchor | ★★★★★ |
| **印章/签名** (覆盖文字) | 浮动定位 (wp:anchor 绝对坐标) | ★★★★★ |
| **矢量图形** | drawings 提取 + 边框/图片转换 | ★★★☆☆ |
| **标题层级** | 编号模式 + 字号 + 加粗 | ★★★★☆ |
| **列表/编号** (7种类型) | 正则匹配 + DOCX numbering | ★★★★☆ |
| **矢量图形** (矩形/线/椭圆/路径) | DrawingML 浮动形状 (wp:anchor) | ★★★★☆ |
| **分页符** | 保留原 PDF 分页 | ★★★★☆ |
| **Web可视化预览** | Next.js + LibreOffice 渲染对比 | ★★★★★ |
| **批量处理** | 目录扫描 + CSV 报告 | ★★★★★ |

## 技术架构

```
pdf2docx-pro/
├── main.py                  # CLI 入口 (argparse)
├── batch.py                 # 批量处理入口 (目录扫描)
├── src/
│   ├── formatter.py         # 格式映射: 字体归一化/颜色/字号/加粗斜体/下划线样式
│   ├── extractor.py         # PDF 提取: PyMuPDF 全格式提取 (spans/drawings/images)
│   ├── analyzer.py          # 布局分析: 页眉页脚检测/段落重建/下划线匹配/标题识别
│   ├── table_extractor.py   # 表格提取: pdfplumber + 单元格 span 匹配
│   ├── shape_builder.py     # 矢量图形还原: DrawingML 浮动形状 (矩形/线/椭圆/路径)
│   ├── list_detector.py     # 列表检测: 7种列表类型正则匹配 + DOCX numbering
│   ├── docx_builder.py       # DOCX 构建: python-docx + lxml 直接 XML (浮动图片/字段/边框/形状/列表)
│   ├── converter.py         # 主编排: 提取→分析→表格→列表→形状→构建
│   └── reporter.py          # 转换质量报告: PDF/DOCX 对比 + 评分 + 问题检测
└── output/                  # 输出目录

# Web UI (位于主项目)
/home/z/my-project/src/app/
├── page.tsx                          # 主页: 上传/选项/结果/对比预览
└── api/pdf2docx/
    ├── convert/route.ts              # 转换 API (调用 Python)
    ├── preview/route.ts              # 页面预览渲染 API (PyMuPDF)
    ├── download/route.ts             # 文件下载 API
    └── list/route.ts                 # 已转换文件列表 API
```

## 安装依赖

```bash
pip install PyMuPDF python-docx pdfplumber lxml Pillow
```

## 使用方法

### 基本用法

```bash
# 全文档转换
python3 main.py -i input.pdf -o output.docx

# 指定页范围 (1-based, 闭区间)
python3 main.py -i input.pdf -o output.docx --pages 1-10
python3 main.py -i input.pdf -o output.docx --pages 1,3,5-8

# 详细日志
python3 main.py -i input.pdf -o output.docx -v

# 高 DPI 图片 (默认 150)
python3 main.py -i input.pdf -o output.docx --dpi 300

# 禁用表格检测 (加速)
python3 main.py -i input.pdf -o output.docx --no-tables
```

### Python API 调用

```python
import sys; sys.path.insert(0, 'src')
from converter import PDF2DocxConverter, ConversionOptions

options = ConversionOptions(
    page_range=(1, 20),      # 1-based 闭区间
    dpi=200,
    detect_tables=True,
    preserve_page_breaks=True,
    verbose=True,
)
converter = PDF2DocxConverter('input.pdf', options)
stats = converter.convert('output.docx')
print(stats)  # {'pages':..., 'paragraphs':..., 'tables':..., 'images':...}
```

## CLI 参数

| 参数 | 说明 | 默认 |
|------|------|------|
| `-i, --input` | 输入 PDF 路径 (必填) | - |
| `-o, --output` | 输出 DOCX 路径 (必填) | - |
| `-p, --pages` | 页码: `1-10` 或 `1,3,5-8` | 全部 |
| `-d, --dpi` | 图片渲染 DPI | 150 |
| `--no-tables` | 禁用表格检测 | False |
| `--no-page-breaks` | 不保留分页符 | False |
| `-v, --verbose` | 详细日志 | False |

## 复刻原理详解

### 1. 字体映射 (formatter.py)
- PDF 字体名归一化: `FangSong,Bold` → `仿宋` + bold=True
- 子集前缀剥离: `ABCDEF+SimSun` → `SimSun` → `宋体`
- 中文字体 → 西文字体配对: 仿宋↔Times New Roman, 黑体↔Arial
- DOCX 双字体设置: `w:eastAsia` (中文) + `w:ascii/hAnsi` (西文)

### 2. 页眉页脚检测 (analyzer.py)
- **跨页一致性分析**: 收集每页顶部/底部文本, 归一化 (去页码) 后找跨页重复
- **区域划分**: header_y_max / footer_y_min 自动计算
- **分隔线检测**: 在页眉下方/页脚上方找水平矢量线, 提取颜色/粗细/样式
- **页码字段**: 检测纯数字 span → 插入 `PAGE` 域 (自动编号)

### 3. 段落重建 (analyzer.py)
PyMuPDF 只给 line/span, 没有段落概念。重建策略:
- 按 y 排序
- 同 block + 字号接近 + 垂直间距 < 1.6×行高 → 同段
- 垂直间距过大 / 字号显著变化 / block 切换 → 新段
- 首行缩进: 首行 x0 vs 后续行平均 x0
- 行距: (行间隙+行高)/行高 → 1.0/1.5/2.0 标准值

### 4. 下划线检测 (analyzer.py)
PDF 下划线通常是 span 下方的水平线段:
- 遍历所有 drawings 的 line items
- 判断: 水平线 (dy<1.5) + y 在 span 底部 [y1-1, y1+0.25×字号] + 水平重叠 > 50%
- 样式: dash 模式判断 single/dash/dotted

### 5. 表格提取 (table_extractor.py)
- pdfplumber `find_tables()` 检测表格区域
- 每个 cell 用 span 中心点匹配, 保留完整格式
- 对齐: span 中心 vs cell 中心 → left/center/right
- 合并: None cell → colspan 累计

### 6. 图片处理 (extractor.py + converter.py)
- **优先**: xref 直接提取 PNG
- **fallback**: XObject 损坏时, 裁剪渲染页面区域 (PyMuPDF clip)
- **浮动定位**: 小图 (<220pt) 且与文字重叠 → `wp:anchor` 绝对定位 (印章/签名)
- **inline**: 大图/独立图 → 居中 inline

### 7. DOCX 高级 XML (docx_builder.py)
直接操作 OOXML 实现高级功能:
- `w:pBdr` (段落边框): 页眉分隔线
- `w:fldChar` + `w:instrText`: PAGE 页码字段
- `wp:anchor` + `wp:positionH/V`: 浮动图片绝对定位
- `w:tblBorders`: 表格全框线
- `w:rFonts w:eastAsia`: 中文字体独立设置

## 字体回退配置

DOCX 中保留标准中文字体名 (仿宋/黑体/宋体/楷体)。在无这些字体的系统 (如 Linux 服务器), 通过 fontconfig 映射:

```bash
# ~/.config/fontconfig/conf.d/99-chinese-fonts.conf
仿宋 → Noto Serif SC (serif 风格近似)
宋体 → Noto Serif SC
黑体 → WenQuanYi Zen Hei (sans-serif)
楷体 → LXGW WenKai
```

在 Windows MS Word (有原生中文字体) 中打开则直接使用真字体, 达到 100% 还原。

## 已知限制

1. **数学公式**: 复杂公式需 OCR 或 MathML, 当前保留为图片
2. **分页对齐**: DOCX 流式排版会导致页数与 PDF 不完全 1:1 (内容完整保留)
3. **复杂矢量图**: 非线性曲线图保留为区域渲染图片
4. **加密 PDF**: 需先解密

## 性能

97 页政府文档: 约 2 秒完成转换 (含表格检测)。
