"""
table_extractor.py - 表格提取器
用 pdfplumber 提取表格结构, 并匹配每个单元格内的 spans (含格式)
"""
from __future__ import annotations
import pdfplumber
from typing import List, Dict, Optional, Tuple, Any
from dataclasses import dataclass, field
from extractor import PageData, SpanData
from analyzer import TableRegion, bbox_overlap


@dataclass
class TableCell:
    """表格单元格"""
    text: str = ''
    spans: List[SpanData] = field(default_factory=list)
    bbox: Tuple[float, float, float, float] = (0, 0, 0, 0)
    rowspan: int = 1
    colspan: int = 1
    is_merged: bool = False  # 被合并的目标
    align: str = 'left'


@dataclass
class ExtractedTable:
    """提取的表格"""
    page_no: int
    bbox: Tuple[float, float, float, float]
    cells: List[List[TableCell]] = field(default_factory=list)  # [row][col]
    rows: int = 0
    cols: int = 0


class TableExtractor:
    """表格提取器"""

    def __init__(self, pdf_path: str, verbose: bool = False):
        self.pdf_path = pdf_path
        self.verbose = verbose
        self._plumber = None

    def _open(self):
        if self._plumber is None:
            self._plumber = pdfplumber.open(self.pdf_path)

    def close(self):
        if self._plumber:
            self._plumber.close()
            self._plumber = None

    def __enter__(self):
        self._open()
        return self

    def __exit__(self, *args):
        self.close()

    def extract_page_tables(self, page_no: int, page_data: PageData) -> List[ExtractedTable]:
        """提取单页所有表格"""
        self._open()
        if page_no >= len(self._plumber.pages):
            return []
        ppage = self._plumber.pages[page_no]
        tables = []
        try:
            # pdfplumber 的 extract_tables 默认策略
            found = ppage.find_tables()
            # 也试 lattice 策略 (有边框的表)
            # found 已用默认策略
        except Exception as e:
            if self.verbose:
                print(f'[TableExtractor] page {page_no} error: {e}')
            return []

        for ft in found:
            try:
                bbox = ft.bbox
                et = ExtractedTable(
                    page_no=page_no,
                    bbox=bbox,
                    rows=len(ft.rows),
                    cols=max(len(r.cells) for r in ft.rows) if ft.rows else 0,
                )
                # 提取每个单元格
                for r_idx, row in enumerate(ft.rows):
                    row_cells = []
                    for c_idx, cell in enumerate(row.cells):
                        if cell is None:
                            # 合并单元格
                            tc = TableCell(is_merged=True)
                            row_cells.append(tc)
                            continue
                        # 找到落在该 cell 内的 spans
                        cell_spans = [
                            s for s in page_data.all_spans
                            if self._span_in_cell(s, cell)
                        ]
                        # 排序: 先 y 后 x
                        cell_spans.sort(key=lambda s: (s.bbox[1], s.bbox[0]))
                        text = ''.join(s.text for s in cell_spans)
                        tc = TableCell(
                            text=text,
                            spans=cell_spans,
                            bbox=cell,
                            align=self._detect_cell_align(cell_spans, cell),
                        )
                        row_cells.append(tc)
                    et.cells.append(row_cells)
                # 处理合并单元格的 rowspan/colspan (简单估算)
                self._detect_merges(et)
                tables.append(et)
            except Exception as e:
                if self.verbose:
                    print(f'[TableExtractor] table extract error: {e}')
                continue
        return tables

    @staticmethod
    def _span_in_cell(span: SpanData, cell) -> bool:
        """判断 span 是否在 cell 内 (中心点)"""
        cx = (cell[0] + cell[2]) / 2
        cy = (cell[1] + cell[3]) / 2
        sx = (span.bbox[0] + span.bbox[2]) / 2
        sy = (span.bbox[1] + span.bbox[3]) / 2
        return cell[0] - 2 <= sx <= cell[2] + 2 and cell[1] - 2 <= sy <= cell[3] + 2

    @staticmethod
    def _detect_cell_align(spans: List[SpanData], cell) -> str:
        """单元格对齐"""
        if not spans:
            return 'left'
        # 水平: 看 span 中心 vs cell 中心
        cell_cx = (cell[0] + cell[2]) / 2
        span_cxs = [(s.bbox[0] + s.bbox[2]) / 2 for s in spans]
        avg_cx = sum(span_cxs) / len(span_cxs)
        cell_w = cell[2] - cell[0]
        if abs(avg_cx - cell_cx) < cell_w * 0.15:
            return 'center'
        # 右对齐: span 靠右
        if cell[2] - max(s.bbox[2] for s in spans) < cell_w * 0.15:
            return 'right'
        return 'left'

    def _detect_merges(self, et: ExtractedTable):
        """简单合并检测: 空单元格可能是被合并"""
        for r_idx, row in enumerate(et.cells):
            for c_idx, cell in enumerate(row):
                if cell.is_merged and c_idx > 0:
                    # 向左找最近的非合并
                    for k in range(c_idx - 1, -1, -1):
                        if not et.cells[r_idx][k].is_merged:
                            et.cells[r_idx][k].colspan += 1
                            break
