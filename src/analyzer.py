"""
analyzer.py - 布局分析器 (核心大脑)
负责:
  1. 页眉/页脚检测 (跨页一致性 + 区域分析)
  2. 段落重建 (从 spans/lines 聚合成段落)
  3. 下划线检测 (匹配 span 与下方矢量线)
  4. 分隔线检测 (页眉页脚分割线)
  5. 列表/编号检测
  6. 表格区域识别 (与 pdfplumber 协作)
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Set
from collections import defaultdict
from extractor import PageData, SpanData, LineData, DrawingData
import math


# ============ 数据结构 ============
@dataclass
class RunFormat:
    """run 级格式"""
    font: str = '宋体'
    latin_font: str = 'Times New Roman'
    size: float = 12.0
    color: str = '000000'
    bold: bool = False
    italic: bool = False
    underline: Optional[str] = None  # None / 'single' / 'double' / 'dash' / 'dotted'
    superscript: bool = False
    highlight: Optional[str] = None
    # 段落级别 (仅 run 的第一个有意义)
    alignment: Optional[str] = None
    line_spacing: Optional[float] = None
    first_line_indent: Optional[float] = None  # pt
    left_indent: Optional[float] = None


@dataclass
class ParagraphData:
    """段落"""
    runs: List[Tuple[str, RunFormat]] = field(default_factory=list)  # (text, format)
    alignment: str = 'left'
    line_spacing: Optional[float] = None
    space_before: float = 0.0
    space_after: float = 0.0
    first_line_indent: float = 0.0  # pt
    left_indent: float = 0.0
    right_indent: float = 0.0
    # 段落级别下划线 (整段下划线 / 分隔线)
    bottom_border: Optional[Dict] = None  # {'style':'single','sz':'6','color':'000000'}
    top_border: Optional[Dict] = None
    bbox: Tuple[float, float, float, float] = (0, 0, 0, 0)
    is_heading: bool = False
    heading_level: int = 0
    page_no: int = 0
    # ★ 列表信息
    is_list_item: bool = False
    list_type: Optional[str] = None     # 'bullet' / 'decimal' / 'lowerLetter' / 'upperLetter' / 'lowerRoman' / 'upperRoman' / 'chinese'
    list_level: int = 0                 # 0-based 嵌套层级
    list_marker_text: Optional[str] = None  # 原始标记文本 (如 "1." / "（一）" / "●")
    # ★ 列表组 ID (同组连续项共享)
    list_group_id: Optional[int] = None


@dataclass
class TableRegion:
    """表格区域"""
    bbox: Tuple[float, float, float, float]
    page_no: int
    rows: int = 0
    cols: int = 0
    cells: List[List[Dict]] = field(default_factory=list)  # [row][col] = {text, format, spans}
    # 边框信息
    borders: Dict = field(default_factory=dict)


@dataclass
class HeaderFooterData:
    """页眉/页脚数据"""
    text_lines: List[ParagraphData] = field(default_factory=list)
    separator_line: Optional[Dict] = None  # {'y':..., 'style':..., 'width':...}
    page_number_span: Optional[SpanData] = None
    bbox: Tuple[float, float, float, float] = (0, 0, 0, 0)
    region: str = ''  # 'header' / 'footer'


@dataclass
class PageLayout:
    """单页布局分析结果"""
    page_no: int
    width: float
    height: float
    header: Optional[HeaderFooterData] = None
    footer: Optional[HeaderFooterData] = None
    body_paragraphs: List[ParagraphData] = field(default_factory=list)
    tables: List[TableRegion] = field(default_factory=list)
    images: List[Dict] = field(default_factory=list)  # {bbox, xref}
    standalone_drawings: List[DrawingData] = field(default_factory=list)


# ============ 几何工具 ============
def bbox_overlap(a, b, tol=1.0) -> bool:
    return not (a[2] < b[0] - tol or a[0] > b[2] + tol or
                a[3] < b[1] - tol or a[1] > b[3] + tol)


def bbox_contains(outer, inner, tol=1.0) -> bool:
    return (outer[0] - tol <= inner[0] and outer[1] - tol <= inner[1] and
            outer[2] + tol >= inner[2] and outer[3] + tol >= inner[3])


def v_distance(a, b) -> float:
    """垂直距离"""
    return abs(a[1] - b[3]) if a[1] > b[3] else abs(b[1] - a[3])


def h_overlap_ratio(a, b) -> float:
    """水平重叠比例"""
    if a[2] <= b[0] or b[2] <= a[0]:
        return 0.0
    overlap = min(a[2], b[2]) - max(a[0], b[0])
    base = min(a[2] - a[0], b[2] - b[0])
    return overlap / base if base > 0 else 0.0


# ============ 主分析器 ============
class LayoutAnalyzer:
    """布局分析器"""

    def __init__(self, pages: List[PageData], verbose: bool = False):
        self.pages = pages
        self.verbose = verbose
        self.page_count = len(pages)
        # 统一页面尺寸 (假设所有页相同)
        self.page_width = pages[0].width if pages else 612
        self.page_height = pages[0].height if pages else 792
        # 全局页眉页脚区域 (y 范围)
        self.header_y_max: Optional[float] = None
        self.footer_y_min: Optional[float] = None

    def log(self, msg: str):
        if self.verbose:
            print(f'[Analyzer] {msg}')

    def analyze(self) -> List[PageLayout]:
        """主分析入口"""
        # 1. 先检测全局页眉/页脚区域
        self._detect_header_footer_region()
        self.log(f'Header y_max={self.header_y_max}, Footer y_min={self.footer_y_min}')

        # 2. 逐页分析
        layouts = []
        for page in self.pages:
            layout = self._analyze_page(page)
            layouts.append(layout)
        return layouts

    # ---------- 页眉页脚区域检测 ----------
    def _detect_header_footer_region(self):
        """跨页检测页眉页脚的 y 范围
        策略: 收集每页顶部/底部的文本, 找出跨页重复的内容
        """
        if self.page_count < 2:
            # 单页, 用经验值
            self.header_y_max = self.page_height * 0.08
            self.footer_y_min = self.page_height * 0.92
            return

        # 收集每页顶部第一行和底部最后一行
        top_lines = []
        bottom_lines = []
        for page in self.pages:
            if not page.all_lines:
                continue
            # 顶部: y 最小的几行
            sorted_lines = sorted(page.all_lines, key=lambda l: l.bbox[1])
            for l in sorted_lines[:3]:
                if l.bbox[1] < self.page_height * 0.15:
                    top_lines.append((page.page_no, l))
            # 底部: y 最大的几行
            sorted_lines_d = sorted(page.all_lines, key=lambda l: -l.bbox[3])
            for l in sorted_lines_d[:3]:
                if l.bbox[3] > self.page_height * 0.85:
                    bottom_lines.append((page.page_no, l))

        # 找跨页重复的顶部文本
        top_texts = defaultdict(list)
        for pno, line in top_lines:
            txt = ''.join(s.text for s in line.spans).strip()
            if txt:
                # 归一化: 去掉页码数字
                import re
                norm = re.sub(r'\d+', '', txt)
                if norm:
                    top_texts[norm].append((pno, line))

        header_candidates = []
        for txt, occurs in top_texts.items():
            if len(occurs) >= max(2, self.page_count // 3):
                # 跨多页重复 = 页眉
                for pno, line in occurs:
                    header_candidates.append(line.bbox)
        if header_candidates:
            self.header_y_max = max(b[3] for b in header_candidates) + 5
        else:
            self.header_y_max = self.page_height * 0.08

        # 页脚同理
        bottom_texts = defaultdict(list)
        for pno, line in bottom_lines:
            txt = ''.join(s.text for s in line.spans).strip()
            if txt:
                import re
                norm = re.sub(r'\d+', '', txt)
                if norm:
                    bottom_texts[norm].append((pno, line))
        footer_candidates = []
        for txt, occurs in bottom_texts.items():
            if len(occurs) >= max(2, self.page_count // 3):
                for pno, line in occurs:
                    footer_candidates.append(line.bbox)
        if footer_candidates:
            self.footer_y_min = min(b[1] for b in footer_candidates) - 5
        else:
            self.footer_y_min = self.page_height * 0.92

    # ---------- 单页分析 ----------
    def _analyze_page(self, page: PageData) -> PageLayout:
        layout = PageLayout(
            page_no=page.page_no,
            width=page.width,
            height=page.height,
        )

        # 区分 body / header / footer
        body_lines: List[LineData] = []
        header_lines: List[LineData] = []
        footer_lines: List[LineData] = []
        for line in page.all_lines:
            y_mid = (line.bbox[1] + line.bbox[3]) / 2
            if self.header_y_max and y_mid < self.header_y_max:
                header_lines.append(line)
            elif self.footer_y_min and y_mid > self.footer_y_min:
                footer_lines.append(line)
            else:
                body_lines.append(line)

        # 页眉
        if header_lines:
            layout.header = self._build_header_footer(page, header_lines, 'header')
        # 页脚
        if footer_lines:
            layout.footer = self._build_header_footer(page, footer_lines, 'footer')

        # body 段落 (不在此排除表格区域, 由 converter 用 pdfplumber 精确表格 bbox 排除)
        # 保留所有 body 行作为段落候选, 后续由 pdfplumber 表格 bbox 过滤
        layout.tables = self._detect_tables(page)
        table_bboxes = []  # 不在 analyzer 层排除, 避免 false positive
        body_lines_outside_table = body_lines  # 保留全部
        layout.body_paragraphs = self._group_lines_to_paragraphs(body_lines_outside_table, page)

        # 图片
        for img in page.images:
            # 排除页眉页脚区域的图
            y_mid = (img.bbox[1] + img.bbox[3]) / 2
            if self.header_y_max and y_mid < self.header_y_max:
                continue
            if self.footer_y_min and y_mid > self.footer_y_min:
                continue
            layout.images.append({'bbox': img.bbox, 'xref': img.xref,
                                  'width': img.width, 'height': img.height})

        # 独立矢量图形 (非下划线/非表格边框/非页眉页脚线)
        used_drawings = self._collect_used_drawings(page, layout)
        for dr in page.drawings:
            if any(dr is ud for ud in used_drawings):
                continue
            # 忽略整页背景框
            if abs(dr.rect[2] - dr.rect[0] - page.width) < 5 and \
               abs(dr.rect[3] - dr.rect[1] - page.height) < 5:
                continue
            y_mid = (dr.rect[1] + dr.rect[3]) / 2
            if self.header_y_max and y_mid < self.header_y_max:
                continue
            if self.footer_y_min and y_mid > self.footer_y_min:
                continue
            layout.standalone_drawings.append(dr)

        return layout

    # ---------- 构建页眉/页脚 ----------
    def _build_header_footer(self, page: PageData, lines: List[LineData],
                            region: str) -> HeaderFooterData:
        """构建页眉或页脚
        自动检测对齐方式:
          - 页眉文字 x1 接近页面右边距 → RIGHT
          - 页码 x0 接近页面中心 → CENTER
        """
        hf = HeaderFooterData(region=region)
        if not lines:
            return hf
        all_bbox = [
            min(l.bbox[0] for l in lines),
            min(l.bbox[1] for l in lines),
            max(l.bbox[2] for l in lines),
            max(l.bbox[3] for l in lines),
        ]
        hf.bbox = tuple(all_bbox)
        # 段落化 (页眉通常1-2行)
        paras = self._group_lines_to_paragraphs(lines, page)
        hf.text_lines = paras
        # ★ 自动检测页眉/页脚对齐方式
        page_w = page.width
        content_right = page_w - 36  # 右边距约36pt
        content_center = page_w / 2
        for p in paras:
            for txt, fmt in p.runs:
                import re
                if re.fullmatch(r'\s*\d+\s*', txt):
                    hf.page_number_span = None  # 标记有页码, 后续用 field
                    # 页码对齐: 检查x位置
                    # 页码bbox已丢失(只存text), 用段落bbox判断
                    p_cx = (p.bbox[0] + p.bbox[2]) / 2
                    if abs(p_cx - content_center) < 30:
                        p.alignment = 'center'  # 居中页码
                    elif p.bbox[2] > content_right - 20:
                        p.alignment = 'right'
                    break
                else:
                    # 普通页眉文字: 检查是否右对齐
                    if p.bbox[2] > content_right - 20:
                        p.alignment = 'right'
        # 找页眉/页脚分隔线 (在该区域附近的水平直线)
        sep = self._find_separator_line(page, all_bbox, region)
        hf.separator_line = sep
        return hf

    def _find_separator_line(self, page: PageData, region_bbox: Tuple, region: str) -> Optional[Dict]:
        """在页眉/页脚区域附近找水平分隔线"""
        candidates = []
        ry0, ry1 = region_bbox[1], region_bbox[3]
        for dr in page.drawings:
            r = dr.rect
            w = r[2] - r[0]
            h = r[3] - r[1]
            # 水平线: 宽>>高
            if w > 20 and h < 3:
                y_mid = (r[1] + r[3]) / 2
                if region == 'header' and y_mid > ry1 - 3 and y_mid < ry1 + 15:
                    candidates.append((r, dr))
                elif region == 'footer' and y_mid > ry0 - 15 and y_mid < ry0 + 3:
                    candidates.append((r, dr))
            # 也可能是 line item
            for item in dr.items:
                if item[0] == 'l':
                    p1, p2 = item[1], item[2]
                    lw = abs(p2.x - p1.x)
                    lh = abs(p2.y - p1.y)
                    if lw > 20 and lh < 3:
                        y_mid = (p1.y + p2.y) / 2
                        if region == 'header' and y_mid > ry1 - 3 and y_mid < ry1 + 15:
                            candidates.append(((p1.x, p1.y, p2.x, p2.y), dr))
                        elif region == 'footer' and y_mid > ry0 - 15 and y_mid < ry0 + 3:
                            candidates.append(((p1.x, p1.y, p2.x, p2.y), dr))
        if not candidates:
            return None
        # 取最长
        best = max(candidates, key=lambda c: (c[0][2] - c[0][0]))
        rect, dr = best
        # 样式
        from formatter import detect_underline_style_from_drawing
        style = detect_underline_style_from_drawing({
            'dash': dr.dash,
        })
        # 颜色
        color = '000000'
        if dr.color:
            color = '%02X%02X%02X' % (
                int(dr.color[0] * 255), int(dr.color[1] * 255), int(dr.color[2] * 255))
        sz = max(4, int((dr.width or 1.0) * 8))  # pt → 1/8pt
        return {
            'y': rect[1],
            'style': style,
            'width': rect[2] - rect[0],
            'sz': str(sz),
            'color': color,
            'x0': rect[0],
            'x1': rect[2],
        }

    # ---------- 行→段落 ----------
    def _group_lines_to_paragraphs(self, lines: List[LineData],
                                   page: PageData) -> List[ParagraphData]:
        """将行聚合成段落
        策略:
          1. 按 y 排序
          2. 相邻行: x0 接近 / 字号接近 / 垂直间距 < 1.5倍行高 → 同段
          3. 行首缩进或新块 → 新段
          4. 字号变化(标题) → 新段
        """
        if not lines:
            return []
        # 按 y 排序
        lines = sorted(lines, key=lambda l: l.bbox[1])
        paragraphs: List[ParagraphData] = []
        current_lines: List[LineData] = []
        current_block_no: Optional[int] = None

        def flush():
            nonlocal current_lines
            if current_lines:
                p = self._build_paragraph(current_lines, page)
                if p:
                    paragraphs.append(p)
                current_lines = []

        for line in lines:
            if not line.spans:
                continue
            # 跳过纯空白行 (但保留作为段落分隔信号)
            line_text = ''.join(s.text for s in line.spans)
            if not line_text.strip():
                flush()
                continue
            # 判断是否同段
            same_block = (current_block_no is None or line.spans[0].block_no == current_block_no)
            if current_lines:
                last = current_lines[-1]
                last_size = max((s.size for s in last.spans), default=12)
                this_size = max((s.size for s in line.spans), default=12)
                last_y1 = last.bbox[3]
                this_y0 = line.bbox[1]
                gap = this_y0 - last_y1
                line_h = max(last_y1 - last.bbox[1], this_y0 - line.bbox[1], this_size)
                # 垂直间距过大 → 新段
                big_gap = gap > line_h * 1.6 if line_h > 0 else gap > 20
                # 字号显著变化 → 新段 (标题)
                size_change = abs(this_size - last_size) > 1.5
                # x0 左移且缩进变化 → 新段 (首行缩进)
                x0_change = abs(line.bbox[0] - last.bbox[0]) > 20
                if big_gap or size_change or not same_block:
                    flush()
                elif x0_change and abs(this_size - last_size) < 0.5:
                    # 缩进变化但字号同: 可能是列表项/换段
                    # 判断: 如果是明显的首行缩进 (中文2字符), 不切; 否则切
                    indent_diff = line.bbox[0] - last.bbox[0]
                    if indent_diff > last_size * 0.8:
                        # 明显首行缩进 → 可能是新段开始
                        flush()
            current_lines.append(line)
            current_block_no = line.spans[0].block_no
        flush()
        return paragraphs

    def _build_paragraph(self, lines: List[LineData], page: PageData) -> Optional[ParagraphData]:
        """从一组行构建段落"""
        if not lines:
            return None
        # 整段 bbox
        x0 = min(l.bbox[0] for l in lines)
        y0 = min(l.bbox[1] for l in lines)
        x1 = max(l.bbox[2] for l in lines)
        y1 = max(l.bbox[3] for l in lines)

        # 对齐: 用首行
        from formatter import detect_alignment
        # 估算边距 (用最宽行)
        max_w_line = max(lines, key=lambda l: l.bbox[2] - l.bbox[0])
        alignment = detect_alignment(max_w_line.bbox, page.width,
                                     left_margin=x0 if x0 > 30 else 36,
                                     right_margin=page.width - x1 if x1 < page.width - 30 else 36)

        # 首行缩进
        first = lines[0]
        rest = lines[1:] if len(lines) > 1 else []
        first_line_indent = 0.0
        if rest:
            avg_x0 = sum(l.bbox[0] for l in rest) / len(rest)
            first_line_indent = first.bbox[0] - avg_x0
            if first_line_indent < 0:
                first_line_indent = 0
        left_indent = 0.0
        if rest:
            left_indent = sum(l.bbox[0] for l in rest) / len(rest)
            # 减去页面左边距 (经验值)
            left_indent = max(0, left_indent - 36)

        # 行距: 固定14pt (匹配原PDF行高)
        # 原PDF: 10pt/14pt字号, 行高约14pt
        # 不再用动态计算(偏差大), 直接用固定值
        first_size = max((s.size for s in first.spans), default=12)
        line_spacing = None  # 不设置段落级行距, 用Normal样式的14pt

        # 段前/段后 (减小间距, 匹配原PDF紧凑排版)
        first_size = max((s.size for s in first.spans), default=12)
        space_before = 0   # 段前0 (原PDF段落间无额外间距)
        space_after = 0    # 段后0

        # 是否标题
        is_heading = False
        heading_level = 0
        first_bold = all(flags_to_bold(s.flags, s.font)[0] for s in first.spans) if first.spans else False
        first_text = ''.join(s.text for s in first.spans).strip()
        # 中文标题特征: 字号大 / 加粗 / 行短 / 编号开头
        import re
        if re.match(r'^[一二三四五六七八九十]+、', first_text) or \
           re.match(r'^第[一二三四五六七八九十]+[章节条]', first_text):
            is_heading = True
            heading_level = 1
        elif re.match(r'^[（(]\s*[一二三四五六七八九十]+\s*[)）]', first_text):
            is_heading = True
            heading_level = 2
        elif re.match(r'^\d+[\.、]', first_text) and first_size >= 14:
            is_heading = True
            heading_level = 3
        elif first_size >= 16 and (first_bold or len(first_text) < 30):
            is_heading = True
            heading_level = 1

        # 标题段前段后不设间距 (匹配原PDF紧凑排版)
        # 原PDF标题与正文间无额外间距, 通过字号区分
        # if is_heading:
        #     space_before = first_size * 0.5
        #     space_after = first_size * 0.3

        para = ParagraphData(
            alignment=alignment,
            line_spacing=line_spacing,
            first_line_indent=first_line_indent,
            left_indent=left_indent,
            space_before=space_before,
            space_after=space_after,
            bbox=(x0, y0, x1, y1),
            is_heading=is_heading,
            heading_level=heading_level,
            page_no=page.page_no,
        )

        # 构建 runs: 每个 span → 一个 run (合并相邻同格式 span)
        from formatter import (normalize_font_name, get_latin_font,
                                int_to_hex_color, flags_to_bold_italic)
        for line in lines:
            for span in line.spans:
                if not span.text:
                    continue
                cjk_font = normalize_font_name(span.font)
                latin_font = get_latin_font(cjk_font)
                bold, italic = flags_to_bold_italic(span.flags, span.font)
                color = int_to_hex_color(span.color)
                # 下划线检测
                underline = self._detect_span_underline(span, page)
                # 上标
                sup = bool(span.flags & 1)
                fmt = RunFormat(
                    font=cjk_font,
                    latin_font=latin_font,
                    size=span.size,
                    color=color,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    superscript=sup,
                )
                para.runs.append((span.text, fmt))
        return para

    # ---------- 下划线检测 ----------
    def _detect_span_underline(self, span: SpanData, page: PageData) -> Optional[str]:
        """检测 span 是否有下划线
        策略: 在 span bbox 下方 1-3pt 范围内找水平线
        """
        x0, y0, x1, y1 = span.bbox
        underline_y_min = y1 - 1
        underline_y_max = y1 + max(2, span.size * 0.25)
        for dr in page.drawings:
            for item in dr.items:
                if item[0] != 'l':
                    continue
                p1, p2 = item[1], item[2]
                # 水平线
                if abs(p1.y - p2.y) < 1.5:
                    ly = (p1.y + p2.y) / 2
                    lx0 = min(p1.x, p2.x)
                    lx1 = max(p1.x, p2.x)
                    if underline_y_min <= ly <= underline_y_max:
                        # 水平重叠
                        overlap_x0 = max(x0, lx0)
                        overlap_x1 = min(x1, lx1)
                        if overlap_x1 - overlap_x0 > (x1 - x0) * 0.5:
                            from formatter import detect_underline_style_from_drawing
                            return detect_underline_style_from_drawing({'dash': dr.dash})
        # PDF 渲染模式下划线 (flags bit? PyMuPDF 没直接给, 但有些字体用 underline)
        # 检查 span.flags: PyMuPDF 不直接给 underline flag, 但有些情况可用 font 特征
        return None

    # ---------- 表格检测 ----------
    def _detect_tables(self, page: PageData) -> List[TableRegion]:
        """检测表格区域
        策略: 用 pdfplumber (在 converter 层注入), 这里先做几何检测:
          - 找出交叉线密集的区域
        """
        # 这里简化: 找出包含多条水平线和垂直线的区域
        h_lines = []  # (y, x0, x1)
        v_lines = []  # (x, y0, y1)
        for dr in page.drawings:
            for item in dr.items:
                if item[0] == 'l':
                    p1, p2 = item[1], item[2]
                    if abs(p1.y - p2.y) < 1.5 and abs(p2.x - p1.x) > 20:
                        h_lines.append(((p1.y + p2.y) / 2, min(p1.x, p2.x), max(p1.x, p2.x)))
                    elif abs(p1.x - p2.x) < 1.5 and abs(p2.y - p1.y) > 20:
                        v_lines.append(((p1.x + p2.x) / 2, min(p1.y, p2.y), max(p1.y, p2.y)))
                elif item[0] == 're':
                    r = item[1]
                    h_lines.append((r.y0, r.x0, r.x1))
                    h_lines.append((r.y1, r.x0, r.x1))
                    v_lines.append((r.x0, r.y0, r.y1))
                    v_lines.append((r.x1, r.y0, r.y1))
        if len(h_lines) < 3 or len(v_lines) < 3:
            return []
        # 聚类水平线 y
        h_lines.sort()
        # 找密集的 y 区间
        # 简化: 整个线密集区作为一个表
        all_h_ys = [l[0] for l in h_lines]
        all_v_xs = [l[0] for l in v_lines]
        if not all_h_ys or not all_v_xs:
            return []
        y_min, y_max = min(all_h_ys), max(all_h_ys)
        x_min, x_max = min(all_v_xs), max(all_v_xs)
        bbox = (x_min, y_min, x_max, y_max)
        # 至少 3x3 才算表
        if (y_max - y_min) < 20 or (x_max - x_min) < 40:
            return []
        tr = TableRegion(
            bbox=bbox,
            page_no=page.page_no,
            rows=len(set(round(y, 1) for y in all_h_ys)),
            cols=len(set(round(x, 1) for x in all_v_xs)),
        )
        return [tr]

    def _collect_used_drawings(self, page: PageData, layout: PageLayout) -> List[DrawingData]:
        """收集已被使用的 drawings (下划线/表格边框/页眉页脚线)"""
        used = []
        # 表格边框
        for tr in layout.tables:
            for dr in page.drawings:
                if bbox_overlap(dr.rect, tr.bbox, tol=2):
                    used.append(dr)
        # 页眉页脚分隔线
        for hf in [layout.header, layout.footer]:
            if hf and hf.separator_line:
                # 找回原 drawing (近似匹配)
                for dr in page.drawings:
                    for item in dr.items:
                        if item[0] == 'l':
                            p1, p2 = item[1], item[2]
                            if abs((p1.y + p2.y) / 2 - hf.separator_line['y']) < 2:
                                used.append(dr)
                                break
        return used


def flags_to_bold(flags: int, font_name: str = '') -> Tuple[bool, bool]:
    """helper: 复用 formatter 的逻辑"""
    from formatter import flags_to_bold_italic
    return flags_to_bold_italic(flags, font_name)
