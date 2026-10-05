"""
docx_builder.py - DOCX 构建器
基于 python-docx + 直接 XML (lxml) 操作, 实现企业级格式复刻:
  - 页面设置 (尺寸/边距/方向)
  - 页眉页脚 (文字 + 页码字段 + 分隔线边框)
  - 段落 (对齐/缩进/行距/段前段后)
  - Run (字体/字号/颜色/加粗/斜体/下划线/上标)
  - 中文字体 (w:eastAsia) + 西文字体 (w:ascii/hAnsi)
  - 表格 (含合并/边框/单元格对齐)
  - 图片插入
  - 矢量图形 (作为图片或边框)
  - 矢量线段/分隔线 (段落底边框)
"""
from __future__ import annotations
import io
from typing import List, Dict, Optional, Tuple, Any
from docx import Document
from docx.shared import Pt, Cm, Mm, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from lxml import etree
from analyzer import (PageLayout, ParagraphData, RunFormat,
                      HeaderFooterData, TableRegion)
from table_extractor import ExtractedTable, TableCell
from formatter import normalize_font_name, get_latin_font


# DOCX 命名空间
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'


def w(tag: str) -> str:
    return f'{{{W_NS}}}{tag}'


class DocxBuilder:
    """DOCX 构建器"""

    def __init__(self, page_width: float = 612, page_height: float = 792,
                 verbose: bool = False):
        self.doc = Document()
        self.page_width = page_width  # pt
        self.page_height = page_height
        self.verbose = verbose
        self._image_counter = 0
        self._numbering_cache = {}  # (list_type, group_id) → numId
        self._numbering_initialized = False
        self._current_page_drawings = []
        self._setup_default_styles()

    def log(self, msg: str):
        if self.verbose:
            print(f'[DocxBuilder] {msg}')

    def _setup_default_styles(self):
        """设置默认样式 (Normal)
        默认字号10.5pt(五号), 固定行距14pt, 匹配中文公文标准
        """
        style = self.doc.styles['Normal']
        style.font.name = 'Times New Roman'
        # 中文字体
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = OxmlElement('w:rFonts')
            rpr.insert(0, rfonts)
        rfonts.set(qn('w:eastAsia'), '宋体')
        rfonts.set(qn('w:ascii'), 'Times New Roman')
        rfonts.set(qn('w:hAnsi'), 'Times New Roman')
        style.font.size = Pt(10.5)  # 五号字
        # 段落固定行距14pt (原PDF行高约14pt, 紧凑排版)
        from docx.enum.text import WD_LINE_SPACING
        pf = style.paragraph_format
        pf.line_spacing = Pt(13)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

    # ---------- 页面设置 ----------
    def setup_page(self, width_pt: float, height_pt: float,
                   margin_top: float = 72, margin_bottom: float = 72,
                   margin_left: float = 72, margin_right: float = 72,
                   header_dist: float = 36, footer_dist: float = 36):
        """设置页面 (单位 pt)"""
        section = self.doc.sections[0]
        section.page_width = Pt(width_pt)
        section.page_height = Pt(height_pt)
        section.top_margin = Pt(margin_top)
        section.bottom_margin = Pt(margin_bottom)
        section.left_margin = Pt(margin_left)
        section.right_margin = Pt(margin_right)
        section.header_distance = Pt(header_dist)
        section.footer_distance = Pt(footer_dist)

    # ---------- 页眉 ----------
    def build_header(self, hf: Optional[HeaderFooterData],
                     page_number_in_header: bool = False):
        """构建页眉"""
        if not hf:
            return
        section = self.doc.sections[0]
        header = section.header
        header.is_linked_to_previous = False
        # 清空默认段落
        for p in list(header.paragraphs):
            p._element.getparent().remove(p._element)
        # 添加内容段落
        if hf.text_lines:
            for para in hf.text_lines:
                self._add_paragraph_to_element(header._element, para,
                                                page_number=page_number_in_header)
        else:
            # 空段落 (保留间距)
            p = header.add_paragraph()
            # 加分隔线
        if hf.separator_line:
            sep = hf.separator_line
            # 添加到 header 的最后一个段落
            paras = header.paragraphs
            if paras:
                self._set_paragraph_bottom_border(paras[-1], sep['style'],
                                                  sep['sz'], sep['color'])

    def build_footer(self, hf: Optional[HeaderFooterData],
                     page_number_in_footer: bool = True):
        """构建页脚"""
        if not hf:
            return
        section = self.doc.sections[0]
        footer = section.footer
        footer.is_linked_to_previous = False
        for p in list(footer.paragraphs):
            p._element.getparent().remove(p._element)
        if hf.text_lines:
            for para in hf.text_lines:
                self._add_paragraph_to_element(footer._element, para,
                                                page_number=page_number_in_footer)
        else:
            p = footer.add_paragraph()
            if hf.separator_line:
                sep = hf.separator_line
                if footer.paragraphs:
                    self._set_paragraph_top_border(footer.paragraphs[-1], sep['style'],
                                                    sep['sz'], sep['color'])

    # ---------- 段落构建 ----------
    def add_paragraph(self, para: ParagraphData, hyperlinks: Optional[List[Dict]] = None):
        """添加段落到文档主体
        Args:
            hyperlinks: 该段落的超链接信息 [{text, uri, anchor, color}]
        Returns:
            创建的 Paragraph 对象 (供 caller 后续操作, 如设置 keepNext)
        """
        p = self.doc.add_paragraph()
        self._apply_paragraph_format(p, para)
        # ★ 列表项: 应用 numbering
        if para.is_list_item and para.list_group_id is not None:
            num_id = self._get_or_create_numbering(para.list_type or 'decimal',
                                                    para.list_group_id)
            self._apply_list_to_paragraph(p, num_id, para.list_level)
        # ★ 超链接处理: 将匹配的链接文本转为 hyperlink
        if hyperlinks:
            self._add_runs_with_hyperlinks(p, para, hyperlinks)
        else:
            for text, fmt in para.runs:
                self._add_run(p, text, fmt)
        # 段落级别边框 (段落分隔线)
        if para.bottom_border:
            self._set_paragraph_bottom_border(p, para.bottom_border.get('style', 'single'),
                                              para.bottom_border.get('sz', '6'),
                                              para.bottom_border.get('color', '000000'))
        if para.top_border:
            self._set_paragraph_top_border(p, para.top_border.get('style', 'single'),
                                            para.top_border.get('sz', '6'),
                                            para.top_border.get('color', '000000'))
        return p

    # ---------- 超链接 ----------
    def _add_runs_with_hyperlinks(self, p, para, hyperlinks: List[Dict]):
        """构建含超链接的 runs
        遍历 para.runs, 若某段文本匹配某个链接, 则用 hyperlink 替代
        """
        try:
            from link_handler import LinkHandler
            if not hasattr(self, '_link_handler'):
                self._link_handler = LinkHandler(self.doc)
            # 简单实现: 遍历 runs, 若 text 完整包含某 link 的 text, 则替换
            for text, fmt in para.runs:
                matched = False
                for hl in hyperlinks:
                    if hl.get('text') and hl['text'] in text:
                        # 切分: 前文 + link + 后文
                        idx = text.find(hl['text'])
                        before = text[:idx]
                        after = text[idx + len(hl['text']):]
                        if before:
                            self._add_run(p, before, fmt)
                        # 加超链接
                        if hl.get('uri'):
                            self._link_handler.add_hyperlink(p, hl['uri'], hl['text'])
                        elif hl.get('anchor'):
                            self._link_handler.add_internal_link(p, hl['anchor'], hl['text'])
                        if after:
                            self._add_run(p, after, fmt)
                        matched = True
                        break
                if not matched:
                    self._add_run(p, text, fmt)
        except Exception as e:
            self.log(f'hyperlink error: {e}')
            # fallback: 普通 runs
            for text, fmt in para.runs:
                self._add_run(p, text, fmt)

    # ---------- 列表/编号 ----------
    def _init_numbering_part(self):
        """初始化 numbering part (如果还没有)
        python-docx 默认有 numbering part, 但需要我们手动添加 abstractNum 和 num
        """
        if self._numbering_initialized:
            return
        try:
            # 获取或创建 numbering part
            from docx.parts.numbering import NumberingPart
            from docx.opc.constants import RELATIONSHIP_TYPE as RT
            try:
                numbering_part = self.doc.part.numbering_part
            except Exception:
                # 新建
                numbering_part = NumberingPart.new()
                self.doc.part.relate_to(numbering_part, RT.NUMBERING)
            self._numbering_part = numbering_part
            self._numbering_initialized = True
        except Exception as e:
            self.log(f'numbering init error: {e}')

    def _get_or_create_numbering(self, list_type: str, group_id: int) -> int:
        """获取或创建一个 abstractNum+num 对, 返回 numId
        不同 list_type 用不同 abstractNum, 同 type+不同 group 用同 abstractNum 不同 numId
        简化: 每个 group 一个独立 numId
        """
        cache_key = (list_type, group_id)
        if cache_key in self._numbering_cache:
            return self._numbering_cache[cache_key]
        self._init_numbering_part()
        if not self._numbering_initialized:
            return 0
        try:
            numbering = self._numbering_part.element
            # 找最大 abstractNumId 和 numId
            existing_abs = numbering.findall(qn('w:abstractNum'))
            max_abs_id = max([int(a.get(qn('w:abstractNumId'))) for a in existing_abs], default=-1)
            existing_nums = numbering.findall(qn('w:num'))
            max_num_id = max([int(n.get(qn('w:numId'))) for n in existing_nums], default=0)
            # 新建 abstractNum
            abs_id = max_abs_id + 1
            num_id = max_num_id + 1
            abstract_num = self._build_abstract_num(abs_id, list_type)
            # 插入到 num 元素之前
            first_num = numbering.find(qn('w:num'))
            if first_num is not None:
                first_num.addprevious(abstract_num)
            else:
                numbering.append(abstract_num)
            # 新建 num
            num_el = OxmlElement('w:num')
            num_el.set(qn('w:numId'), str(num_id))
            abs_ref = OxmlElement('w:abstractNumId')
            abs_ref.set(qn('w:val'), str(abs_id))
            num_el.append(abs_ref)
            numbering.append(num_el)
            self._numbering_cache[cache_key] = num_id
            return num_id
        except Exception as e:
            self.log(f'create numbering error: {e}')
            return 0

    def _build_abstract_num(self, abs_id: int, list_type: str):
        """构建 abstractNum XML, 9 级"""
        abstract = OxmlElement('w:abstractNum')
        abstract.set(qn('w:abstractNumId'), str(abs_id))
        # 多级类型映射
        # list_type → numFmt + lvlText
        type_map = {
            'bullet':          ('bullet',  '·'),
            'bullet_dash':     ('bullet',  '–'),
            'decimal':         ('decimal', '%1.'),
            'decimal_paren':   ('decimal', '%1)'),
            'decimal_fullparen': ('decimal', '(%1)'),
            'lowerLetter':     ('lowerLetter', '%1.'),
            'upperLetter':     ('upperLetter', '%1.'),
            'lowerLetter_paren': ('lowerLetter', '%1)'),
            'upperLetter_paren': ('upperLetter', '%1)'),
            'lowerRoman':      ('lowerRoman', '%1.'),
            'upperRoman':      ('upperRoman', '%1.'),
            'chinese':         ('chineseCounting', '%1、'),
            'chinese_paren':   ('chineseCounting', '(%1)'),
            'chinese_ordinal': ('chineseCounting', '第%1条'),
            'circled':         ('decimal', '%1.'),
        }
        fmt, lvl_text = type_map.get(list_type, ('decimal', '%1.'))
        # 9 级
        for level in range(9):
            lvl = OxmlElement('w:lvl')
            lvl.set(qn('w:ilvl'), str(level))
            # start
            start = OxmlElement('w:start')
            start.set(qn('w:val'), '1')
            lvl.append(start)
            # numFmt
            nf = OxmlElement('w:numFmt')
            nf.set(qn('w:val'), fmt)
            lvl.append(nf)
            # lvlText
            lt = OxmlElement('w:lvlText')
            lt.set(qn('w:val'), lvl_text)
            lvl.append(lt)
            # lvlJc
            jc = OxmlElement('w:lvlJc')
            jc.set(qn('w:val'), 'left')
            lvl.append(jc)
            # pPr 缩进
            pPr = OxmlElement('w:pPr')
            ind = OxmlElement('w:ind')
            indent_left = 360 + level * 360  # 每级 0.25 inch
            indent_hang = 360
            ind.set(qn('w:left'), str(indent_left))
            ind.set(qn('w:hanging'), str(indent_hang))
            pPr.append(ind)
            lvl.append(pPr)
            abstract.append(lvl)
        return abstract

    def _apply_list_to_paragraph(self, p, num_id: int, level: int):
        """给段落应用 numbering (numId + ilvl)"""
        if num_id <= 0:
            return
        pPr = p._element.get_or_add_pPr()
        # 移除已有 numPr
        existing = pPr.find(qn('w:numPr'))
        if existing is not None:
            pPr.remove(existing)
        numPr = OxmlElement('w:numPr')
        ilvl = OxmlElement('w:ilvl')
        ilvl.set(qn('w:val'), str(min(level, 8)))
        numPr.append(ilvl)
        numId_el = OxmlElement('w:numId')
        numId_el.set(qn('w:val'), str(num_id))
        numPr.append(numId_el)
        # 插入到 pPr 开头
        pPr.insert(0, numPr)

    def _add_paragraph_to_element(self, container, para: ParagraphData,
                                   page_number: bool = False):
        p_el = OxmlElement('w:p')
        container.append(p_el)
        # 包装成 paragraph 对象
        from docx.text.paragraph import Paragraph
        p = Paragraph(p_el, container.getparent() if hasattr(container, 'getparent') else None)
        self._apply_paragraph_format(p, para)
        for text, fmt in para.runs:
            # 如果是页码占位符
            if page_number and (text.strip().isdigit() or text.strip() == ''):
                # 插入页码字段
                self._add_page_number_field(p, fmt)
            else:
                self._add_run(p, text, fmt)
        return p

    def _apply_paragraph_format(self, p, para: ParagraphData):
        """应用段落格式"""
        pf = p.paragraph_format
        # 对齐
        align_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        p.alignment = align_map.get(para.alignment, WD_ALIGN_PARAGRAPH.LEFT)
        # 缩进
        if para.first_line_indent > 0:
            pf.first_line_indent = Pt(para.first_line_indent)
        if para.left_indent > 0:
            pf.left_indent = Pt(para.left_indent)
        # 行距: 支持固定行距(Pt值)和倍数行距
        if para.line_spacing:
            if isinstance(para.line_spacing, (int, float)) and para.line_spacing > 3:
                # 大于3的值视为固定行距(Pt), 如14.0
                pf.line_spacing = Pt(para.line_spacing)
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            else:
                # 倍数行距, 如1.0/1.5/2.0
                pf.line_spacing = para.line_spacing
                pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        # 段前段后
        if para.space_before:
            pf.space_before = Pt(para.space_before)
        if para.space_after:
            pf.space_after = Pt(para.space_after)

    def _add_run(self, p, text: str, fmt: RunFormat):
        """添加 run (含完整格式)"""
        if not text:
            return
        run = p.add_run(text)
        # 字号
        run.font.size = Pt(fmt.size)
        # 颜色
        if fmt.color and fmt.color != '000000':
            try:
                run.font.color.rgb = RGBColor.from_string(fmt.color)
            except Exception:
                pass
        # 加粗/斜体
        run.font.bold = fmt.bold if fmt.bold else None
        run.font.italic = fmt.italic if fmt.italic else None
        # 下划线
        if fmt.underline:
            from docx.enum.text import WD_UNDERLINE
            ul_map = {
                'single': WD_UNDERLINE.SINGLE,
                'double': WD_UNDERLINE.DOUBLE,
                'dash': WD_UNDERLINE.DASH,
                'dotted': WD_UNDERLINE.DOTTED,
                'dashLong': WD_UNDERLINE.DASH_LONG,
                'dotDash': WD_UNDERLINE.DOT_DASH,
                'wave': WD_UNDERLINE.WAVY,
                'heavyWave': WD_UNDERLINE.WAVY_HEAVY,
            }
            run.font.underline = ul_map.get(fmt.underline, WD_UNDERLINE.SINGLE)
        # 上标
        if fmt.superscript:
            run.font.superscript = True
        # 字体 (中西文分别设置)
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = OxmlElement('w:rFonts')
            rpr.insert(0, rfonts)
        rfonts.set(qn('w:eastAsia'), fmt.font)
        rfonts.set(qn('w:ascii'), fmt.latin_font)
        rfonts.set(qn('w:hAnsi'), fmt.latin_font)
        rfonts.set(qn('w:cs'), fmt.latin_font)
        # 中文字号也设置 (w:szCs 用于复杂脚本)
        sz_cs = rpr.find(qn('w:szCs'))
        if sz_cs is None:
            sz_cs = OxmlElement('w:szCs')
            rpr.append(sz_cs)
        sz_cs.set(qn('w:val'), str(int(fmt.size * 2)))
        return run

    def _add_page_number_field(self, p, fmt: RunFormat):
        """插入页码字段 (自动编号)"""
        run = p.add_run()
        self._add_run_fmt(run, fmt)
        # 插入字段
        fld_begin = OxmlElement('w:fldChar')
        fld_begin.set(qn('w:fldCharType'), 'begin')
        run._element.append(fld_begin)
        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = ' PAGE '
        run._element.append(instr)
        fld_end = OxmlElement('w:fldChar')
        fld_end.set(qn('w:fldCharType'), 'end')
        run._element.append(fld_end)

    def _add_run_fmt(self, run, fmt: RunFormat):
        """复用格式设置"""
        run.font.size = Pt(fmt.size)
        if fmt.color and fmt.color != '000000':
            try:
                run.font.color.rgb = RGBColor.from_string(fmt.color)
            except Exception:
                pass
        run.font.bold = fmt.bold if fmt.bold else None
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = OxmlElement('w:rFonts')
            rpr.insert(0, rfonts)
        rfonts.set(qn('w:eastAsia'), fmt.font)
        rfonts.set(qn('w:ascii'), fmt.latin_font)
        rfonts.set(qn('w:hAnsi'), fmt.latin_font)

    # ---------- 段落边框 (分隔线) ----------
    def _set_paragraph_bottom_border(self, p, style: str, sz: str, color: str):
        """设置段落底边框 (如下划线/分隔线)"""
        pPr = p._element.get_or_add_pPr()
        pbdr = pPr.find(qn('w:pBdr'))
        if pbdr is None:
            pbdr = OxmlElement('w:pBdr')
            pPr.append(pbdr)
        bottom = pbdr.find(qn('w:bottom'))
        if bottom is not None:
            pbdr.remove(bottom)
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), style if style else 'single')
        bottom.set(qn('w:sz'), sz if sz else '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), color if color else '000000')
        pbdr.append(bottom)

    def _set_paragraph_top_border(self, p, style: str, sz: str, color: str):
        pPr = p._element.get_or_add_pPr()
        pbdr = pPr.find(qn('w:pBdr'))
        if pbdr is None:
            pbdr = OxmlElement('w:pBdr')
            pPr.append(pbdr)
        top = pbdr.find(qn('w:top'))
        if top is not None:
            pbdr.remove(top)
        top = OxmlElement('w:top')
        top.set(qn('w:val'), style if style else 'single')
        top.set(qn('w:sz'), sz if sz else '6')
        top.set(qn('w:space'), '1')
        top.set(qn('w:color'), color if color else '000000')
        pbdr.append(top)

    # ---------- 表格 ----------
    def add_table(self, table: ExtractedTable):
        """添加表格"""
        if not table.cells:
            return
        n_rows = len(table.cells)
        n_cols = max(len(r) for r in table.cells) if table.cells else 0
        if n_rows == 0 or n_cols == 0:
            return
        docx_table = self.doc.add_table(rows=n_rows, cols=n_cols)
        docx_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        # 设置表格边框 (全框)
        self._set_table_borders(docx_table)
        # 填充
        for r_idx, row in enumerate(table.cells):
            for c_idx, cell in enumerate(row):
                if c_idx >= n_cols:
                    continue
                docx_cell = docx_table.cell(r_idx, c_idx)
                # 清空
                docx_cell.text = ''
                # 单元格内容
                if cell.spans:
                    from formatter import (int_to_hex_color, flags_to_bold_italic)
                    # 需要该页的 drawings 用于下划线检测 - 通过闭包注入
                    page_drawings = getattr(self, '_current_page_drawings', [])
                    for i, span in enumerate(cell.spans):
                        if i == 0:
                            run = docx_cell.paragraphs[0].add_run(span.text)
                            align_map = {
                                'left': WD_ALIGN_PARAGRAPH.LEFT,
                                'center': WD_ALIGN_PARAGRAPH.CENTER,
                                'right': WD_ALIGN_PARAGRAPH.RIGHT,
                            }
                            docx_cell.paragraphs[0].alignment = align_map.get(cell.align, WD_ALIGN_PARAGRAPH.LEFT)
                        else:
                            run = docx_cell.add_paragraph().add_run(span.text)
                        cjk_font = normalize_font_name(span.font)
                        latin_font = get_latin_font(cjk_font)
                        bold, italic = flags_to_bold_italic(span.flags, span.font)
                        color = int_to_hex_color(span.color)
                        # 下划线检测 (用该页 drawings)
                        underline = self._detect_cell_underline(span, page_drawings)
                        run.font.size = Pt(span.size)
                        if color != '000000':
                            try:
                                run.font.color.rgb = RGBColor.from_string(color)
                            except Exception:
                                pass
                        run.font.bold = bold if bold else None
                        run.font.italic = italic if italic else None
                        if underline:
                            from docx.enum.text import WD_UNDERLINE
                            ul_map = {'single': WD_UNDERLINE.SINGLE, 'double': WD_UNDERLINE.DOUBLE,
                                      'dash': WD_UNDERLINE.DASH, 'dotted': WD_UNDERLINE.DOTTED,
                                      'dashLong': WD_UNDERLINE.DASH_LONG,
                                      'dotDash': WD_UNDERLINE.DOT_DASH,
                                      'wave': WD_UNDERLINE.WAVY,
                                      'heavyWave': WD_UNDERLINE.WAVY_HEAVY}
                            run.font.underline = ul_map.get(underline, WD_UNDERLINE.SINGLE)
                        rpr = run._element.get_or_add_rPr()
                        rfonts = rpr.find(qn('w:rFonts'))
                        if rfonts is None:
                            rfonts = OxmlElement('w:rFonts')
                            rpr.insert(0, rfonts)
                        rfonts.set(qn('w:eastAsia'), cjk_font)
                        rfonts.set(qn('w:ascii'), latin_font)
                        rfonts.set(qn('w:hAnsi'), latin_font)
                # 合并
                if cell.colspan > 1 or cell.is_merged:
                    pass  # 简化处理, 后续可优化
        # 处理合并单元格
        self._merge_cells(docx_table, table)

    def _detect_cell_underline(self, span, page_drawings) -> Optional[str]:
        """检测表格单元格 span 的下划线 (与正文逻辑相同)"""
        if not page_drawings:
            return None
        x0, y0, x1, y1 = span.bbox
        underline_y_min = y1 - 1
        underline_y_max = y1 + max(2, span.size * 0.25)
        for dr in page_drawings:
            for item in dr.items:
                if item[0] != 'l':
                    continue
                p1, p2 = item[1], item[2]
                if abs(p1.y - p2.y) < 1.5:
                    ly = (p1.y + p2.y) / 2
                    lx0 = min(p1.x, p2.x)
                    lx1 = max(p1.x, p2.x)
                    if underline_y_min <= ly <= underline_y_max:
                        overlap_x0 = max(x0, lx0)
                        overlap_x1 = min(x1, lx1)
                        if overlap_x1 - overlap_x0 > (x1 - x0) * 0.5:
                            from formatter import detect_underline_style_from_drawing
                            return detect_underline_style_from_drawing({'dash': dr.dash})
        return None

    def _set_table_borders(self, table):
        tbl = table._tbl
        tblPr = tbl.tblPr
        borders = tblPr.find(qn('w:tblBorders'))
        if borders is not None:
            tblPr.remove(borders)
        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b = OxmlElement(f'w:{edge}')
            b.set(qn('w:val'), 'single')
            b.set(qn('w:sz'), '4')
            b.set(qn('w:space'), '0')
            b.set(qn('w:color'), '000000')
            borders.append(b)
        tblPr.append(borders)

    def _merge_cells(self, docx_table, et: ExtractedTable):
        """处理合并单元格"""
        for r_idx, row in enumerate(et.cells):
            c_idx = 0
            while c_idx < len(row):
                cell = row[c_idx]
                if cell.colspan > 1 and c_idx + cell.colspan - 1 < len(row):
                    try:
                        start = docx_table.cell(r_idx, c_idx)
                        end = docx_table.cell(r_idx, c_idx + cell.colspan - 1)
                        start.merge(end)
                    except Exception:
                        pass
                c_idx += 1

    # ---------- 图片 ----------
    def add_image(self, image_bytes: bytes, width_pt: Optional[float] = None,
                  height_pt: Optional[float] = None, ext: str = 'png'):
        """添加图片 (inline, 居中)"""
        try:
            stream = io.BytesIO(image_bytes)
            kwargs = {}
            if width_pt:
                kwargs['width'] = Pt(width_pt)
            if height_pt:
                kwargs['height'] = Pt(height_pt)
            self.doc.add_picture(stream, **kwargs)
            last = self.doc.paragraphs[-1]
            last.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._image_counter += 1
        except Exception as e:
            self.log(f'image insert error: {e}')

    def add_floating_image(self, image_bytes: bytes,
                           x_pt: float, y_pt: float,
                           width_pt: float, height_pt: float,
                           ext: str = 'png', behind_text: bool = False):
        """添加浮动图片 (绝对定位, 用于印章/印章等需覆盖文字的场景)
        x_pt, y_pt: 相对页面左上角的坐标 (pt)
        behind_text: True=衬于文字下方 (印章常见), False=浮于文字上方
        """
        try:
            from docx.parts.image import ImagePart
            from docx.opc.constants import RELATIONSHIP_TYPE as RT
            # 添加图片到文档 part, 拿到 rId
            # 注意: get_or_add_image 返回 (rId, image_part)
            rId, image_part = self.doc.part.get_or_add_image(io.BytesIO(image_bytes))
            # 创建段落 + run, 在 run 里放 anchored drawing
            p = self.doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            run = p.add_run()
            # EMU 转换: 1 pt = 12700 EMU
            cx = int(width_pt * 12700)
            cy = int(height_pt * 12700)
            pos_x = int(x_pt * 12700)
            pos_y = int(y_pt * 12700)
            self._image_counter += 1
            img_id = self._image_counter + 5000

            xml = f'''<wp:anchor xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}" xmlns:r="{R_NS}" distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="{'1' if behind_text else '0'}" locked="0" layoutInCell="1" allowOverlap="1">
  <wp:simplePos x="0" y="0"/>
  <wp:positionH relativeFrom="page"><wp:posOffset>{pos_x}</wp:posOffset></wp:positionH>
  <wp:positionV relativeFrom="page"><wp:posOffset>{pos_y}</wp:posOffset></wp:positionV>
  <wp:extent cx="{cx}" cy="{cy}"/>
  <wp:effectExtent l="0" t="0" r="0" b="0"/>
  <wp:wrapNone/>
  <wp:docPr id="{img_id}" name="Picture {img_id}"/>
  <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
  <a:graphic>
    <a:graphicData uri="{PIC_NS}">
      <pic:pic>
        <pic:nvPicPr>
          <pic:cNvPr id="{img_id}" name="image{self._image_counter}.png"/>
          <pic:cNvPicPr/>
        </pic:nvPicPr>
        <pic:blipFill>
          <a:blip r:embed="{rId}"/>
          <a:stretch><a:fillRect/></a:stretch>
        </pic:blipFill>
        <pic:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </pic:spPr>
      </pic:pic>
    </a:graphicData>
  </a:graphic>
</wp:anchor>'''
            from lxml import etree
            anchor_el = etree.fromstring(xml)
            drawing = OxmlElement('w:drawing')
            drawing.append(anchor_el)
            run._element.append(drawing)
        except Exception as e:
            self.log(f'floating image error: {e}')
            import traceback
            traceback.print_exc()

    # ---------- 矢量图形 ----------
    def add_drawing_as_image(self, image_bytes: bytes, width_pt: float, height_pt: float):
        """把矢量图形渲染成图片插入"""
        self.add_image(image_bytes, width_pt=width_pt, height_pt=height_pt)

    def add_horizontal_line(self, color: str = '000000', size: str = '6',
                             space: str = '1'):
        """添加一条水平分隔线"""
        p = self.doc.add_paragraph()
        pPr = p._element.get_or_add_pPr()
        pbdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), size)
        bottom.set(qn('w:space'), space)
        bottom.set(qn('w:color'), color)
        pbdr.append(bottom)
        pPr.append(pbdr)

    # ---------- 分页对齐优化 ----------
    def apply_paragraph_keep_with_next(self, p):
        """设置段落 keep_with_next (与下一段同页, 用于标题)"""
        try:
            pPr = p._element.get_or_add_pPr()
            existing = pPr.find(qn('w:keepNext'))
            if existing is not None:
                pPr.remove(existing)
            keepNext = OxmlElement('w:keepNext')
            keepNext.set(qn('w:val'), 'true')
            pPr.insert(0, keepNext)
        except Exception:
            pass

    def apply_paragraph_keep_lines(self, p):
        """设置段落 keep_lines (段内不分页)"""
        try:
            pPr = p._element.get_or_add_pPr()
            existing = pPr.find(qn('w:keepLines'))
            if existing is not None:
                pPr.remove(existing)
            keepLines = OxmlElement('w:keepLines')
            keepLines.set(qn('w:val'), 'true')
            pPr.insert(0, keepLines)
        except Exception:
            pass

    def prevent_table_row_split(self, table):
        """设置表格行不可分页"""
        try:
            for row in table.rows:
                trPr = row._tr.get_or_add_trPr()
                existing = trPr.find(qn('w:cantSplit'))
                if existing is not None:
                    trPr.remove(existing)
                cantSplit = OxmlElement('w:cantSplit')
                trPr.append(cantSplit)
        except Exception:
            pass

    def set_table_header_row(self, table, row_idx: int = 0):
        """设置表格行为表头 (跨页重复)"""
        try:
            if row_idx < len(table.rows):
                row = table.rows[row_idx]
                trPr = row._tr.get_or_add_trPr()
                existing = trPr.find(qn('w:tblHeader'))
                if existing is not None:
                    trPr.remove(existing)
                tblHeader = OxmlElement('w:tblHeader')
                tblHeader.set(qn('w:val'), 'true')
                trPr.append(tblHeader)
        except Exception:
            pass

    def fill_blank_lines(self, n_lines: int, size_pt: float = 1.0):
        """填充空行 (用于页面边界对齐)"""
        try:
            for _ in range(min(n_lines, 200)):
                p = self.doc.add_paragraph()
                pf = p.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.line_spacing = 1.0
                run = p.add_run('')
                run.font.size = Pt(size_pt)
        except Exception:
            pass

    # ---------- 分页符 ----------
    def add_page_break(self):
        """添加分页符"""
        from docx.enum.text import WD_BREAK
        p = self.doc.add_paragraph()
        run = p.add_run()
        run.add_break(WD_BREAK.PAGE)

    # ---------- 保存 ----------
    def save(self, output_path: str):
        self.doc.save(output_path)
        self.log(f'saved to {output_path}')

    # ---------- 数学公式 (OMML) ----------
    def add_formula(self, omml_xml: str, latex: str = '', display: bool = True):
        """插入 OMML 公式 (Word 原生公式格式)
        Args:
            omml_xml: 完整的 OMML XML 字符串 (含 m:oMathPara)
            latex: LaTeX 形式 (备用, 作为文档属性)
            display: True=独立公式行(居中), False=内联
        """
        try:
            from lxml import etree
            # 解析 OMML XML
            try:
                omml_el = etree.fromstring(omml_xml)
            except Exception:
                # XML 解析失败, 用纯文本 fallback
                p = self.doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(f'[公式] {latex}')
                run.font.italic = True
                return
            # 创建段落
            p = self.doc.add_paragraph()
            if display:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # 把 OMML 元素插入段落
            p._element.append(omml_el)
        except Exception as e:
            self.log(f'formula insert error: {e}')
            # fallback: 纯文本
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(latex or omml_xml)
            run.font.italic = True
