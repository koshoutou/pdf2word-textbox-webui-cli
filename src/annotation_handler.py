"""
annotation_handler.py - PDF 注释还原器
将 PDF 注释 (annotations) 还原为 DOCX 对应元素:
  - Highlight (高亮): w:shd (段落底纹) 或 run highlight
  - Text/Popup (批注): w:comment + 批注引用
  - Underline (下划线注释): run underline
  - StrikeOut (删除线): run strike
  - Squiggly (波浪线): run wavy underline
  - Stamp (图章): 浮动图片
  - Redact (标记): 灰色背景
"""
from __future__ import annotations
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass, field
import fitz


@dataclass
class AnnotationInfo:
    """注释信息"""
    annot_type: str        # 'Highlight' / 'Text' / 'Underline' / 'StrikeOut' / 'Squiggly' / 'Stamp' / 'Redact'
    rect: Tuple[float, float, float, float]
    page_no: int
    content: str = ''      # 批注内容
    title: str = ''        # 作者
    subject: str = ''
    creation_date: str = ''
    color: Optional[Tuple[float, float, float]] = None  # RGB 0-1
    quad_points: List[Tuple[float, float]] = field(default_factory=list)  # 高亮四边形点
    icon: str = ''         # Text注释图标 (Comment/Note/Help等)


class AnnotationDetector:
    """注释检测器"""

    def __init__(self, pdf_path: str, verbose: bool = False):
        self.pdf_path = pdf_path
        self.verbose = verbose

    def detect_all_annotations(self) -> Dict[int, List[AnnotationInfo]]:
        """检测所有页的注释, 返回 {page_no: [AnnotationInfo]}"""
        result = {}
        try:
            doc = fitz.open(self.pdf_path)
            for pno in range(len(doc)):
                page = doc[pno]
                try:
                    annots = list(page.annots())
                except Exception:
                    annots = []
                if not annots:
                    continue
                page_annots = []
                for a in annots:
                    info = self._parse_annot(a, pno)
                    if info:
                        page_annots.append(info)
                if page_annots:
                    result[pno] = page_annots
                    if self.verbose:
                        print(f'[Annot] p{pno+1}: {len(page_annots)} annotations')
            doc.close()
        except Exception as e:
            if self.verbose:
                print(f'[Annot] error: {e}')
        return result

    def _parse_annot(self, annot, page_no: int) -> Optional[AnnotationInfo]:
        """解析单个注释"""
        try:
            type_info = annot.type  # (subtype, name)
            subtype = type_info[1] if isinstance(type_info, tuple) else str(type_info)
            info = annot.info
            rect = tuple(annot.rect)
            color = annot.colors.get('stroke') if hasattr(annot, 'colors') else None
            quad_points = []
            try:
                qp = annot.get_quad_points() if hasattr(annot, 'get_quad_points') else None
                if qp:
                    for i in range(0, len(qp), 2):
                        quad_points.append((qp[i], qp[i+1]))
            except Exception:
                pass
            return AnnotationInfo(
                annot_type=subtype,
                rect=rect,
                page_no=page_no,
                content=info.get('content', ''),
                title=info.get('title', ''),
                subject=info.get('subject', ''),
                creation_date=info.get('creationDate', ''),
                color=tuple(color) if color else None,
                quad_points=quad_points,
                icon=getattr(annot, 'icon_name', '') if hasattr(annot, 'icon_name') else '',
            )
        except Exception as e:
            if self.verbose:
                print(f'[Annot] parse error: {e}')
            return None


def color_to_hex(color: Optional[Tuple[float, float, float]]) -> Optional[str]:
    """RGB (0-1) → hex string"""
    if not color or len(color) < 3:
        return None
    r, g, b = color[0], color[1], color[2]
    return '%02X%02X%02X' % (int(r * 255), int(g * 255), int(b * 255))


def get_highlight_color(annot: AnnotationInfo) -> str:
    """获取高亮颜色 (默认黄色)"""
    hex_color = color_to_hex(annot.color)
    if hex_color:
        return hex_color
    return 'FFFF00'  # 默认黄色


def match_annot_to_spans(annot: AnnotationInfo, spans) -> List:
    """匹配注释与 span 列表 (返回重叠的 spans)"""
    matched = []
    ax0, ay0, ax1, ay1 = annot.rect
    for span in spans:
        sx0, sy0, sx1, sy1 = span.bbox
        # 重叠判断
        if ax0 < sx1 and ax1 > sx0 and ay0 < sy1 and ay1 > sy0:
            matched.append(span)
    return matched
