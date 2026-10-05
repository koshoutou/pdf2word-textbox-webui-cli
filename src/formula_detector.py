"""
formula_detector.py - 数学公式识别器
检测 PDF 中的数学公式区域, 并转换为 OMML (Office MathML) 或保留为图片

识别策略:
  1. 字体特征: Math/Symbol/CMS/Cambria Math 字体
  2. Unicode 数学符号: ∑∫√∞≤≥≠±×÷ αβγδεθλμπσφψω∂∇
  3. 上下标结构: span flags bit 0 (superscript) + 字号比例
  4. 公式块特征: 同行多字号混排 + 大量上下标
  5. 几何特征: 公式通常居中或独立成段

输出:
  - 公式区域的 bbox + 文本 + LaTeX 形式
  - 对应的 OMML XML (供 docx_builder 插入)
"""
from __future__ import annotations
import re
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass, field

# 数学字体关键词
MATH_FONT_KEYWORDS = ('math', 'symbol', 'cms', 'cambria', 'stix', 'latex', 'msam', 'msbm', 'eu')

# Unicode 数学符号范围
MATH_SYMBOLS = set('∑∫√∞≤≥≠±×÷αβγδεζηθικλμνξπρστυφχψω∂∇∇≃≅≈≡∝∈∉∀∃∧∨¬')
GREEK_UPPER = set('ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ')
MATH_OPERATORS = set('∧∨¬⊂⊃⊆⊇⊕⊗⊙∉∊∼∼≈≠≤≥≦≧≺≻⊕')
SUPERSCRIPT_CHARS = '⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ'
SUBSCRIPT_CHARS = '₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₙₐₑₒₓₔ'

# 数学函数名 (用于 LaTeX 映射)
FUNCTION_NAMES = {'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'log', 'ln', 'lg',
                  'exp', 'lim', 'max', 'min', 'sup', 'inf', 'det', 'dim', 'arg'}


@dataclass
class FormulaSpan:
    """公式内的 span"""
    text: str
    font: str
    size: float
    bbox: Tuple[float, float, float, float]
    flags: int
    is_superscript: bool = False
    is_subscript: bool = False


@dataclass
class FormulaRegion:
    """公式区域"""
    bbox: Tuple[float, float, float, float]
    spans: List[FormulaSpan] = field(default_factory=list)
    raw_text: str = ''
    latex: str = ''           # LaTeX 形式
    omml_xml: str = ''        # OMML XML
    page_no: int = 0
    formula_type: str = 'inline'  # 'inline' / 'display'


class FormulaDetector:
    """数学公式检测器"""

    def __init__(self, verbose: bool = False):
        self.verbose = verbose

    def detect_page_formulas(self, page_data) -> List[FormulaRegion]:
        """检测单页所有公式"""
        formulas = []
        # 1. 收集所有含数学特征的 span
        math_spans = []
        for span in page_data.all_spans:
            if self._is_math_span(span):
                math_spans.append(span)
        if not math_spans:
            return formulas
        # 2. 按 (block, line) 聚类
        clusters = self._cluster_spans(math_spans, page_data)
        for cluster in clusters:
            formula = self._build_formula(cluster, page_data.page_no)
            if formula:
                formulas.append(formula)
        return formulas

    def _is_math_span(self, span) -> bool:
        """判断 span 是否含数学特征"""
        # 字体特征
        font = (span.font or '').lower()
        for kw in MATH_FONT_KEYWORDS:
            if kw in font:
                return True
        # 文本特征
        text = span.text or ''
        for c in text:
            if c in MATH_SYMBOLS or c in GREEK_UPPER or c in MATH_OPERATORS:
                return True
            if c in SUPERSCRIPT_CHARS or c in SUBSCRIPT_CHARS:
                return True
        # 上下标 flag
        if span.flags & 1:  # superscript bit
            return True
        return False

    def _cluster_spans(self, spans, page_data) -> List[List]:
        """将数学 span 聚类成公式组 (同行/同block)"""
        if not spans:
            return []
        # 按 (block_no, line_no) 分组
        groups = {}
        for s in spans:
            key = (s.block_no, s.line_no)
            groups.setdefault(key, []).append(s)
        # 合并相邻行 (垂直距离接近)
        sorted_keys = sorted(groups.keys())
        clusters = []
        current = []
        last_key = None
        for key in sorted_keys:
            spans_in_line = groups[key]
            if not current:
                current = list(spans_in_line)
                last_key = key
            else:
                # 判断与上一行是否同公式 (垂直距离 < 行高*1.5 且水平重叠)
                last_spans = current
                last_y1 = max(s.bbox[3] for s in last_spans)
                this_y0 = min(s.bbox[1] for s in spans_in_line)
                line_h = max((s.bbox[3] - s.bbox[1]) for s in last_spans) if last_spans else 12
                gap = this_y0 - last_y1
                if gap < line_h * 1.2 and self._horizontal_overlap(last_spans, spans_in_line):
                    current.extend(spans_in_line)
                else:
                    clusters.append(current)
                    current = list(spans_in_line)
                    last_key = key
        if current:
            clusters.append(current)
        return clusters

    def _horizontal_overlap(self, spans_a, spans_b) -> bool:
        """两组 span 是否水平重叠"""
        a_x0 = min(s.bbox[0] for s in spans_a)
        a_x1 = max(s.bbox[2] for s in spans_a)
        b_x0 = min(s.bbox[0] for s in spans_b)
        b_x1 = max(s.bbox[2] for s in spans_b)
        return not (a_x1 < b_x0 or b_x1 < a_x0)

    def _build_formula(self, spans: List, page_no: int) -> Optional[FormulaRegion]:
        """从一组 span 构建公式"""
        if not spans:
            return None
        # bbox
        x0 = min(s.bbox[0] for s in spans)
        y0 = min(s.bbox[1] for s in spans)
        x1 = max(s.bbox[2] for s in spans)
        y1 = max(s.bbox[3] for s in spans)
        # 判断 inline vs display
        width = x1 - x0
        is_display = width > 200  # 独立公式行
        # 构建 FormulaSpan (标注上下标)
        fspans = []
        # 用字号中位数作为基准
        sizes = [s.size for s in spans]
        base_size = sorted(sizes)[len(sizes)//2]  # 中位数
        for s in sorted(spans, key=lambda x: (x.bbox[1], x.bbox[0])):
            is_sup = bool(s.flags & 1) or (s.size < base_size * 0.75 and s.size > base_size * 0.5)
            is_sub = False
            # 通过 y 偏移判断上下标
            if not is_sup:
                y_center = (s.bbox[1] + s.bbox[3]) / 2
                # 与同行其他 span 对比
                same_line = [x for x in spans if abs((x.bbox[1]+x.bbox[3])/2 - y_center) < 3]
                if same_line:
                    line_base_size = max(x.size for x in same_line)
                    if s.size < line_base_size * 0.75:
                        # 偏下 = 下标
                        if y_center > (max(x.bbox[1]+x.bbox[3] for x in same_line)/2 + 0):
                            # 简化: 比同行中位数偏下
                            med_y = sorted([(x.bbox[1]+x.bbox[3])/2 for x in same_line])[len(same_line)//2]
                            if y_center > med_y + 1:
                                is_sub = True
                            else:
                                is_sup = True
            fspans.append(FormulaSpan(
                text=s.text, font=s.font, size=s.size,
                bbox=s.bbox, flags=s.flags,
                is_superscript=is_sup, is_subscript=is_sub,
            ))
        # 生成 LaTeX
        latex = self._to_latex(fspans)
        # 生成 OMML
        omml = self._to_omml(fspans)
        raw_text = ''.join(s.text for s in fspans)
        return FormulaRegion(
            bbox=(x0, y0, x1, y1),
            spans=fspans,
            raw_text=raw_text,
            latex=latex,
            omml_xml=omml,
            page_no=page_no,
            formula_type='display' if is_display else 'inline',
        )

    # ---------- LaTeX 生成 ----------
    def _to_latex(self, fspans: List[FormulaSpan]) -> str:
        """生成 LaTeX 形式"""
        result = []
        for s in fspans:
            t = s.text
            # Unicode → LaTeX
            t = self._unicode_to_latex(t)
            if s.is_superscript:
                result.append(f'^{{{t}}}')
            elif s.is_subscript:
                result.append(f'_{{{t}}}')
            else:
                result.append(t)
        return ''.join(result)

    def _unicode_to_latex(self, text: str) -> str:
        """Unicode 数学符号 → LaTeX 命令"""
        mapping = {
            '∑': r'\sum', '∫': r'\int', '√': r'\sqrt',
            '∞': r'\infty', '≤': r'\leq', '≥': r'\geq',
            '≠': r'\neq', '±': r'\pm', '×': r'\times',
            '÷': r'\div', '∂': r'\partial', '∇': r'\nabla',
            'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma',
            'δ': r'\delta', 'ε': r'\epsilon', 'ζ': r'\zeta',
            'η': r'\eta', 'θ': r'\theta', 'λ': r'\lambda',
            'μ': r'\mu', 'ν': r'\nu', 'ξ': r'\xi',
            'π': r'\pi', 'ρ': r'\rho', 'σ': r'\sigma',
            'τ': r'\tau', 'φ': r'\phi', 'χ': r'\chi',
            'ψ': r'\psi', 'ω': r'\omega',
            'Θ': r'\Theta', 'Λ': r'\Lambda', 'Σ': r'\Sigma',
            'Φ': r'\Phi', 'Ψ': r'\Psi', 'Ω': r'\Omega',
            '∈': r'\in', '∉': r'\notin', '∀': r'\forall',
            '∃': r'\exists', '∧': r'\wedge', '∨': r'\vee', '¬': r'\neg',
            '⊂': r'\subset', '⊃': r'\supset', '⊆': r'\subseteq',
            '≈': r'\approx', '≡': r'\equiv', '∝': r'\propto',
            '·': r'\cdot', '…': r'\ldots',
            '⁰': '0', '¹': '1', '²': '2', '³': '3',
            '⁴': '4', '⁵': '5', '⁶': '6', '⁷': '7',
            '⁸': '8', '⁹': '9',
            '₀': '0', '₁': '1', '₂': '2', '₃': '3',
        }
        out = []
        for c in text:
            if c in mapping:
                out.append(mapping[c])
            else:
                out.append(c)
        return ''.join(out)

    # ---------- OMML (Office MathML) 生成 ----------
    def _to_omml(self, fspans: List[FormulaSpan]) -> str:
        """生成 OMML XML (Word 原生公式格式)
        简化版: 直接把 spans 转换为 m:r 元素, 上下标用 m:sSup/m:sSub
        """
        M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
        parts = ['<m:oMathPara xmlns:m="' + M_NS + '"><m:oMath>']
        i = 0
        while i < len(fspans):
            s = fspans[i]
            t = self._escape_xml(s.text)
            if s.is_superscript:
                # 找前一个非上下标 span 作为 base
                if parts and '<m:r>' in parts[-1]:
                    # 包裹前一个为 sSup
                    last = parts.pop()
                    # 提取文本
                    import re as _re
                    m = _re.search(r'<m:t>(.*?)</m:t>', last)
                    base_text = m.group(1) if m else ''
                    parts.append(f'<m:sSup><m:e><m:r><m:t>{base_text}</m:t></m:r></m:e><m:sup><m:r><m:t>{t}</m:t></m:r></m:sup></m:sSup>')
                else:
                    parts.append(f'<m:r><m:t>^{t}</m:t></m:r>')
            elif s.is_subscript:
                if parts and '<m:r>' in parts[-1]:
                    last = parts.pop()
                    import re as _re
                    m = _re.search(r'<m:t>(.*?)</m:t>', last)
                    base_text = m.group(1) if m else ''
                    parts.append(f'<m:sSub><m:e><m:r><m:t>{base_text}</m:t></m:r></m:e><m:sub><m:r><m:t>{t}</m:t></m:r></m:sub></m:sSub>')
                else:
                    parts.append(f'<m:r><m:t>_{t}</m:t></m:r>')
            else:
                parts.append(f'<m:r><m:t>{t}</m:t></m:r>')
            i += 1
        parts.append('</m:oMath></m:oMathPara>')
        return ''.join(parts)

    @staticmethod
    def _escape_xml(text: str) -> str:
        return (text.replace('&', '&amp;')
                   .replace('<', '&lt;')
                   .replace('>', '&gt;')
                   .replace('"', '&quot;'))

    # ---------- 公式区域过滤 ----------
    def is_span_in_formula(self, span, formulas: List[FormulaRegion]) -> bool:
        """判断 span 是否在公式区域内"""
        for f in formulas:
            if (f.bbox[0] - 2 <= span.bbox[0] and span.bbox[2] <= f.bbox[2] + 2 and
                f.bbox[1] - 2 <= span.bbox[1] and span.bbox[3] <= f.bbox[3] + 2):
                return True
        return False
