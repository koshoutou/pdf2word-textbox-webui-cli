"""
extractor.py - PDF 提取器
基于 PyMuPDF 提取每页的:
  - 文本 spans (含字体/字号/颜色/位置/flags)
  - 矢量图形 (drawings: lines/rects/curves)
  - 图片
  - 链接/注释
返回统一的数据结构 PageData
"""
from __future__ import annotations
import fitz  # PyMuPDF
import sys
import os
# 抑制 MuPDF 的非致命警告 (如 XObject 损坏 n3) - 通过 stderr 重定向
import io as _io
if not os.environ.get('PDF2DOCX_DEBUG'):
    # 重定向 MuPDF 错误输出到 /dev/null
    try:
        _devnull = open(os.devnull, 'w')
        fitz.TOOLS.mupdf_warnings(_devnull)
    except Exception:
        pass
os.environ.setdefault('PYMUPDF_MSG', '0')
try:
    fitz.TOOLS.mupdf_display_errors(False)
except Exception:
    pass
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Tuple


@dataclass
class SpanData:
    """单个文本 span"""
    text: str
    font: str               # 原始字体名
    size: float             # pt
    color: int              # RGB int
    flags: int              # PyMuPDF flags
    bbox: Tuple[float, float, float, float]  # (x0, y0, x1, y1)
    block_no: int
    line_no: int
    span_no: int
    origin: Tuple[float, float] = (0, 0)  # baseline origin
    ascender: float = 0.0
    descender: float = 0.0


@dataclass
class LineData:
    """一行文本 (由多个 span 组成)"""
    bbox: Tuple[float, float, float, float]
    spans: List[SpanData] = field(default_factory=list)
    direction: int = 0  # 0=LTR, 1=RTL
    wmode: int = 0      # 0=horizontal, 1=vertical


@dataclass
class BlockData:
    """一个文本块"""
    bbox: Tuple[float, float, float, float]
    lines: List[LineData] = field(default_factory=list)
    block_type: int = 0  # 0=text, 1=image


@dataclass
class DrawingData:
    """矢量图形项"""
    rect: Tuple[float, float, float, float]
    type: str             # 'line' / 'rect' / 'curve' / 'fill' / 'stroke' / 'quade'
    items: List[Any]      # fitz drawing items
    color: Optional[Tuple[float, float, float]] = None     # stroke color (0-1)
    fill: Optional[Tuple[float, float, float]] = None      # fill color
    width: float = 1.0
    dash: Optional[List[float]] = None
    line_count: int = 0   # 包含的线段数


@dataclass
class ImageData:
    """图片"""
    bbox: Tuple[float, float, float, float]
    xref: int
    width: int
    height: int
    ext: str = 'png'


@dataclass
class LinkData:
    """链接"""
    rect: Tuple[float, float, float, float]
    uri: str = ''
    kind: str = ''


@dataclass
class PageData:
    """单页数据"""
    page_no: int                          # 0-based
    width: float
    height: float
    rotation: int = 0
    blocks: List[BlockData] = field(default_factory=list)
    drawings: List[DrawingData] = field(default_factory=list)
    images: List[ImageData] = field(default_factory=list)
    links: List[LinkData] = field(default_factory=list)
    # 派生: 所有 spans 平铺列表 (方便分析)
    all_spans: List[SpanData] = field(default_factory=list)
    all_lines: List[LineData] = field(default_factory=list)


class PDFExtractor:
    """PDF 提取器"""

    def __init__(self, pdf_path: str, password: str = None):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        # 修复缺陷#9: 加密PDF检测和密码校验
        if self.doc.is_encrypted:
            if password:
                if not self.doc.authenticate(password):
                    raise ValueError(f"PDF已加密, 提供的密码不正确: {pdf_path}")
            else:
                raise ValueError(f"PDF已加密, 需要提供密码: {pdf_path}")
        self.page_count = len(self.doc)

    def close(self):
        self.doc.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

    def extract_page(self, page_no: int) -> PageData:
        """提取单页全部数据"""
        page = self.doc[page_no]
        rect = page.rect
        pd = PageData(
            page_no=page_no,
            width=rect.width,
            height=rect.height,
            rotation=page.rotation,
        )
        # 1. 文本 dict
        # 注意: 不要传 flags 参数, 默认 TEXTFLAGS_DICT 才包含 image block
        d = page.get_text("dict")
        for b_idx, block in enumerate(d.get('blocks', [])):
            bd = BlockData(
                bbox=tuple(block['bbox']),
                block_type=block.get('type', 0),
            )
            if block.get('type', 0) == 0:  # text
                for l_idx, line in enumerate(block.get('lines', [])):
                    ld = LineData(
                        bbox=tuple(line['bbox']),
                        direction=line.get('dir', [1, 0])[0],
                        wmode=line.get('wmode', 0),
                    )
                    for s_idx, span in enumerate(line.get('spans', [])):
                        sd = SpanData(
                            text=span.get('text', ''),
                            font=span.get('font', ''),
                            size=span.get('size', 12.0),
                            color=span.get('color', 0),
                            flags=span.get('flags', 0),
                            bbox=tuple(span['bbox']),
                            block_no=b_idx,
                            line_no=l_idx,
                            span_no=s_idx,
                            origin=tuple(span.get('origin', (0, 0))),
                        )
                        ld.spans.append(sd)
                        pd.all_spans.append(sd)
                    bd.lines.append(ld)
                    pd.all_lines.append(ld)
            pd.blocks.append(bd)

        # 2. 矢量图形
        try:
            drawings = page.get_drawings()
        except Exception:
            drawings = []
        for dr in drawings:
            dd = DrawingData(
                rect=tuple(dr.get('rect', (0, 0, 0, 0))),
                type=self._classify_drawing(dr),
                items=dr.get('items', []),
                color=dr.get('color'),
                fill=dr.get('fill'),
                width=dr.get('width', 1.0),
                dash=dr.get('dash'),
                line_count=sum(1 for it in dr.get('items', []) if it[0] == 'l'),
            )
            pd.drawings.append(dd)

        # 3. 图片
        try:
            images = page.get_images(full=True)
        except Exception:
            images = []
        seen_xrefs = set()
        for img in images:
            xref = img[0]
            try:
                pix = fitz.Pixmap(self.doc, xref)
                w, h = pix.width, pix.height
                ext = 'png'
                if pix.alpha:
                    ext = 'png'
                elif pix.colorspace and pix.colorspace.n == 1:
                    ext = 'png'
                pix = None
            except Exception:
                w, h, ext = 0, 0, 'png'
            # 找到图片在页面的位置
            try:
                rects = list(page.get_image_rects(xref))
            except Exception:
                rects = []
            if not rects:
                # 没有位置信息, 跳过 (后面用 image block 补)
                continue
            for inst in rects:
                pd.images.append(ImageData(
                    bbox=tuple(inst),
                    xref=xref,
                    width=w,
                    height=h,
                    ext=ext,
                ))
                seen_xrefs.add(xref)

        # 3b. 从文本 dict 的 image block 补充 (处理 XObject 损坏的情况)
        for block in d.get('blocks', []):
            if block.get('type', 0) == 1:
                bbox = tuple(block['bbox'])
                # 是否已覆盖
                covered = any(bbox_overlap(bbox, im.bbox, tol=2) for im in pd.images)
                if not covered:
                    pd.images.append(ImageData(
                        bbox=bbox,
                        xref=-1,  # 标记需用页面裁剪渲染
                        width=int(bbox[2] - bbox[0]),
                        height=int(bbox[3] - bbox[1]),
                        ext='png',
                    ))

        # 4. 链接
        try:
            for link in page.get_links():
                pd.links.append(LinkData(
                    rect=tuple(link.get('from', (0, 0, 0, 0))),
                    uri=link.get('uri', ''),
                    kind=str(link.get('kind', '')),
                ))
        except Exception:
            pass

        return pd

    @staticmethod
    def _classify_drawing(dr: dict) -> str:
        items = dr.get('items', [])
        if not items:
            return 'empty'
        types = set(it[0] for it in items)
        if types == {'l'}:
            return 'line'
        if types == {'re'}:
            return 'rect'
        if types <= {'c', 'qu'}:  # curves
            return 'curve'
        if 're' in types and 'l' in types:
            return 'rect+line'
        return 'mixed'

    def extract_pages(self, page_range: Optional[Tuple[int, int]] = None,
                      page_list: Optional[List[int]] = None) -> List[PageData]:
        """提取多页"""
        if page_list is not None:
            pages = page_list
        elif page_range is not None:
            start, end = page_range
            pages = list(range(start, end + 1))
        else:
            pages = list(range(self.page_count))
        result = []
        for pno in pages:
            if 0 <= pno < self.page_count:
                result.append(self.extract_page(pno))
        return result

    def get_image_pixmap(self, xref: int, dpi: int = 300) -> Optional[bytes]:
        """导出图片为 PNG bytes"""
        try:
            pix = fitz.Pixmap(self.doc, xref)
            if pix.alpha:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            return pix.tobytes('png')
        except Exception:
            return None

    def render_region_image(self, page_no: int, bbox: Tuple[float, float, float, float],
                            dpi: int = 200) -> Optional[bytes]:
        """裁剪渲染页面的指定区域为 PNG (用于 XObject 损坏的图片)"""
        try:
            page = self.doc[page_no]
            # 加一点 padding 避免边缘裁切
            pad = 1.0
            clip = fitz.Rect(bbox[0] - pad, bbox[1] - pad,
                             bbox[2] + pad, bbox[3] + pad)
            clip.intersect(page.rect)
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat, clip=clip)
            return pix.tobytes('png')
        except Exception as e:
            return None

    def render_page_image(self, page_no: int, dpi: int = 150) -> Optional[bytes]:
        """整页渲染为 PNG (fallback 用)"""
        try:
            page = self.doc[page_no]
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)
            return pix.tobytes('png')
        except Exception:
            return None


def bbox_overlap(a, b, tol=1.0) -> bool:
    """判断两个 bbox 是否重叠"""
    return not (a[2] < b[0] - tol or a[0] > b[2] + tol or
                a[3] < b[1] - tol or a[1] > b[3] + tol)
