"""
page_aligner.py - 分页对齐优化器
解决 PDF→DOCX 转换后的分页错位问题
策略:
  1. 段落 keep_with_next: 让标题段与下一段保持在同一页
  2. 段落 keep_lines: 段内不分页
  3. 表格行 cant_split: 表格行不分页
  4. 精确边距: 根据原PDF页边距设置DOCX边距
  5. 内容溢出检测: 每页内容超过原PDF页面时, 用空段填充
"""
from __future__ import annotations
from typing import List, Optional, Tuple
from docx.shared import Pt, Cm, Emu
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class PageAligner:
    """分页对齐优化器"""

    def __init__(self, page_width_pt: float = 612, page_height_pt: float = 792,
                 margin_top: float = 72, margin_bottom: float = 72,
                 margin_left: float = 72, margin_right: float = 72):
        self.page_width = page_width_pt
        self.page_height = page_height_pt
        self.margin_top = margin_top_pt = margin_top
        self.margin_bottom = margin_bottom
        self.margin_left = margin_left
        self.margin_right = margin_right
        # 内容区高度 (pt)
        self.content_height = page_height_pt - margin_top - margin_bottom

    def apply_keep_with_next(self, paragraph):
        """设置段落 keep_with_next (与下一段同页)"""
        pPr = paragraph._element.get_or_add_pPr()
        # 移除已有的
        existing = pPr.find(qn('w:keepNext'))
        if existing is not None:
            pPr.remove(existing)
        keepNext = OxmlElement('w:keepNext')
        keepNext.set(qn('w:val'), 'true')
        # 插入到 pPr 开头
        pPr.insert(0, keepNext)

    def apply_keep_lines(self, paragraph):
        """设置段落 keep_lines (段内不分页)"""
        pPr = paragraph._element.get_or_add_pPr()
        existing = pPr.find(qn('w:keepLines'))
        if existing is not None:
            pPr.remove(existing)
        keepLines = OxmlElement('w:keepLines')
        keepLines.set(qn('w:val'), 'true')
        pPr.insert(0, keepLines)

    def apply_widow_control(self, paragraph, enable: bool = True):
        """设置孤行控制 (避免段落末尾单行孤悬)"""
        pPr = paragraph._element.get_or_add_pPr()
        existing = pPr.find(qn('w:widowControl'))
        if existing is not None:
            pPr.remove(existing)
        wc = OxmlElement('w:widowControl')
        wc.set(qn('w:val'), 'true' if enable else 'false')
        pPr.insert(0, wc)

    def prevent_table_row_split(self, table):
        """设置表格行不可分页 (cant_split)"""
        for row in table.rows:
            trPr = row._tr.get_or_add_trPr()
            existing = trPr.find(qn('w:cantSplit'))
            if existing is not None:
                trPr.remove(existing)
            cantSplit = OxmlElement('w:cantSplit')
            trPr.append(cantSplit)

    def set_row_as_header(self, table, row_idx: int = 0):
        """设置表格行为表头 (跨页重复)"""
        if row_idx < len(table.rows):
            row = table.rows[row_idx]
            trPr = row._tr.get_or_add_trPr()
            existing = trPr.find(qn('w:tblHeader'))
            if existing is not None:
                trPr.remove(existing)
            tblHeader = OxmlElement('w:tblHeader')
            tblHeader.set(qn('w:val'), 'true')
            trPr.append(tblHeader)

    def estimate_paragraph_height_pt(self, paragraph) -> float:
        """估算段落高度 (pt)
        公式: (行数 × 字号 × 行距) + 段前 + 段后
        """
        pf = paragraph.paragraph_format
        # 字号
        sizes = []
        for run in paragraph.runs:
            if run.font.size:
                sizes.append(run.font.size.pt)
        base_size = max(sizes) if sizes else 12
        # 行数估算: 字符数 / 每行字符数
        text = paragraph.text or ''
        # 中文每行约 (内容宽度 / 字号) 个字符
        content_width_pt = self.page_width - self.margin_left - self.margin_right
        chars_per_line = max(1, int(content_width_pt / base_size))
        char_count = len(text)
        line_count = max(1, (char_count + chars_per_line - 1) // chars_per_line)
        # 行距
        line_spacing = 1.0
        if pf.line_spacing:
            try:
                line_spacing = float(pf.line_spacing)
            except Exception:
                line_spacing = 1.0
        line_height = base_size * line_spacing
        # 段前段后
        space_before = pf.space_before.pt if pf.space_before else 0
        space_after = pf.space_after.pt if pf.space_after else 0
        # 首行缩进不影响高度
        total = line_count * line_height + space_before + space_after
        return total

    def fill_to_page_boundary(self, doc, current_height: float, target_height: float):
        """如果当前页内容高度不足目标, 用空段填充
        避免下页内容提前出现在本页
        """
        if current_height >= target_height:
            return current_height
        remaining = target_height - current_height
        # 加空段 (字号小, 行距固定)
        fill_size = 1  # 1pt 字号
        line_h = fill_size * 1.0
        n_lines = int(remaining // line_h)
        for _ in range(min(n_lines, 200)):  # 限制最多200行避免异常
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            run = p.add_run('')
            run.font.size = Pt(fill_size)
        return target_height

    def setup_section_page(self, section, page_width_pt: float, page_height_pt: float,
                            margin_top: float, margin_bottom: float,
                            margin_left: float, margin_right: float,
                            header_distance: float = 36, footer_distance: float = 36):
        """精确设置页面尺寸和边距"""
        section.page_width = Pt(page_width_pt)
        section.page_height = Pt(page_height_pt)
        section.top_margin = Pt(margin_top)
        section.bottom_margin = Pt(margin_bottom)
        section.left_margin = Pt(margin_left)
        section.right_margin = Pt(margin_right)
        section.header_distance = Pt(header_distance)
        section.footer_distance = Pt(footer_distance)
