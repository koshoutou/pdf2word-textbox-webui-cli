"""
formatter.py - 格式映射器
负责 PDF 格式信息 → DOCX 格式信息的映射转换
包括: 字体名归一化、颜色转换、字号、加粗/斜体检测、下划线识别
"""
from __future__ import annotations
import re
from typing import Optional, Tuple

# ============ 字体映射表 ============
# PDF 中常见的字体名 → DOCX 的中文字体对应
# key 为小写匹配
FONT_NAME_MAP = {
    # 仿宋族
    'fangsong': '仿宋',
    'fangsong,bold': '仿宋',
    'fs': '仿宋',
    'stfangsong': '华文仿宋',
    # 黑体族
    'simhei': '黑体',
    'heiti': '黑体',
    'stheiti': '华文黑体',
    'microsoftyahei': '微软雅黑',
    'microsoftyaheibold': '微软雅黑',
    'msyh': '微软雅黑',
    'msyhbd': '微软雅黑',
    # 宋体族
    'simsun': '宋体',
    'simsun,bold': '宋体',
    'songti': '宋体',
    'nsimsun': '新宋体',
    'stsong': '华文宋体',
    'stzhongsong': '华文中宋',
    # 楷体族
    'kaiti': '楷体',
    'stkaiti': '华文楷体',
    'simkai': '楷体',
    # 西文
    'timesnewroman': 'Times New Roman',
    'times': 'Times New Roman',
    'arial': 'Arial',
    'helvetica': 'Arial',
    'calibri': 'Calibri',
    'couriernew': 'Courier New',
    'courier': 'Courier New',
}

# 西文字体判定 (latin)
LATIN_FONT_MAP = {
    '仿宋': 'Times New Roman',
    '宋体': 'Times New Roman',
    '黑体': 'Arial',
    '微软雅黑': 'Arial',
    '楷体': 'Times New Roman',
    '华文仿宋': 'Times New Roman',
    '华文宋体': 'Times New Roman',
    '华文黑体': 'Arial',
    '华文楷体': 'Times New Roman',
    '华文中宋': 'Times New Roman',
    '新宋体': 'Times New Roman',
}


def normalize_font_name(font_name: str) -> str:
    """归一化 PDF 字体名为标准中文字体名"""
    if not font_name:
        return '宋体'
    # 去掉子集前缀 (如 ABCDEF+SimSun)
    if '+' in font_name:
        font_name = font_name.split('+', 1)[1]
    # 去掉逗号后的 Bold 等后缀做主匹配
    base = font_name.split(',')[0]
    key = base.strip().lower().replace(' ', '')
    # 直接精确匹配
    if font_name.strip().lower() in FONT_NAME_MAP:
        return FONT_NAME_MAP[font_name.strip().lower()]
    if key in FONT_NAME_MAP:
        return FONT_NAME_MAP[key]
    # 模糊匹配
    for k, v in FONT_NAME_MAP.items():
        if k in key or key in k:
            return v
    # 含 heiti/hei 的中文黑体判断
    lower = font_name.lower()
    if 'hei' in lower:
        return '黑体'
    if 'song' in lower or 'simsun' in lower:
        return '宋体'
    if 'fang' in lower:
        return '仿宋'
    if 'kai' in lower:
        return '楷体'
    if 'yahei' in lower:
        return '微软雅黑'
    return '宋体'


def get_latin_font(cjk_font: str) -> str:
    """获取中文字体对应的西文字体"""
    return LATIN_FONT_MAP.get(cjk_font, 'Times New Roman')


# ============ 颜色转换 ============
def int_to_hex_color(color_int: int) -> str:
    """PyMuPDF 的颜色 (int RGB) → DOCX 的 hex 字符串"""
    if color_int is None:
        return '000000'
    # PyMuPDF 返回的是 sRGB int
    r = (color_int >> 16) & 0xFF
    g = (color_int >> 8) & 0xFF
    b = color_int & 0xFF
    return f'{r:02X}{g:02X}{b:02X}'


def hex_to_rgb_tuple(hex_str: str) -> Tuple[int, int, int]:
    """hex 颜色 → RGB tuple"""
    h = hex_str.lstrip('#')
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ============ 加粗/斜体检测 ============
# PyMuPDF span flags 位含义:
# bit 0: superscript (上标)
# bit 1: italic (斜体)
# bit 2: serifed (衬线)
# bit 3: monospaced (等宽)
# bit 4: bold (加粗)


def flags_to_bold_italic(flags: int, font_name: str = '') -> Tuple[bool, bool]:
    """从 flags 和字体名判断 bold/italic"""
    bold = bool(flags & 16)
    italic = bool(flags & 2)
    # 从字体名二次确认
    fn = font_name.lower()
    if 'bold' in fn or 'black' in fn or 'heavy' in fn:
        bold = True
    if 'italic' in fn or 'oblique' in fn:
        italic = True
    # 中文仿宋加粗有时是 FangSong,Bold
    if 'bold' in fn:
        bold = True
    return bold, italic


def detect_superscript_subscript(flags: int, size_ratio: Optional[float] = None) -> Optional[str]:
    """检测上标/下标"""
    if flags & 1:  # superscript bit
        return 'superscript'
    # 通过字号比例判断 (子字号明显小于父字号)
    if size_ratio is not None and 0.5 < size_ratio < 0.75:
        return 'superscript'
    return None


# ============ 字号映射 ============
# PDF 字号 (pt) → DOCX 字号 (半磅, Pt*2)
def pt_to_half_pt(pt: float) -> int:
    """磅 → 半磅 (python-docx 用 Pt, 内部存半磅)"""
    return int(round(pt * 2))


# ============ 字号到中文字号的映射 ============
# 中文排版常用字号
CHINESE_FONT_SIZE = {
    '初号': 42.0,
    '小初': 36.0,
    '一号': 26.0,
    '小一': 24.0,
    '二号': 22.0,
    '小二': 18.0,
    '三号': 16.0,
    '小三': 15.0,
    '四号': 14.0,
    '小四': 12.0,
    '五号': 10.5,
    '小五': 9.0,
    '六号': 7.5,
    '小六': 6.5,
    '七号': 5.5,
    '八号': 5.0,
}


def pt_to_chinese_size(pt: float) -> Optional[str]:
    """磅 → 中文字号名 (近似)"""
    for name, size in CHINESE_FONT_SIZE.items():
        if abs(pt - size) < 0.6:
            return name
    return None


# ============ 下划线类型 ============
UNDERLINE_PATTERNS = {
    'single': 'single',
    'double': 'double',
    'dash': 'dash',
    'dotted': 'dotted',
    'dashLong': 'dashLong',
    'dotDash': 'dotDash',
    'wave': 'wave',
    'heavyWave': 'heavyWave',
    'halfFrame': 'halfFrame',
}


def detect_underline_style_from_drawing(drawing_item: dict, line_width: float = 1.0) -> str:
    """根据矢量线段特征判断下划线样式
    支持识别:
      - single (单实线)
      - double (双线: 间距很小且并行)
      - dash (虚线: 长破折)
      - dotted (点线: 短点)
      - dashLong (长虚线)
      - dotDash (点划线)
      - wave (波浪线)
    """
    dash = drawing_item.get('dash') or []
    if dash:
        if len(dash) >= 2:
            on, off = dash[0], dash[1]
            # 短点 (on<1)
            if on < 1:
                return 'dotted'
            # 长虚线 (on>5)
            if on > 5:
                return 'dashLong'
            # 点划线 (dash 数组长度>=4 且交替)
            if len(dash) >= 4:
                return 'dotDash'
            # 普通虚线
            return 'dash'
        return 'dash'
    # 检查是否波浪线 (drawing item 类型为 c/qu)
    items = drawing_item.get('items') or []
    if items:
        types = set(it[0] for it in items)
        if 'c' in types or 'qu' in types:
            # 含曲线 → 可能是波浪线
            return 'wave'
    # 双线检测: 同一区域有两条平行线
    if drawing_item.get('_is_double'):
        return 'double'
    return 'single'


# ============ 对齐方式检测 ============
def detect_alignment(line_bbox: Tuple[float, float, float, float],
                     page_width: float,
                     left_margin: float,
                     right_margin: float) -> str:
    """根据行的 bbox 和页面边距判断对齐方式"""
    x0, y0, x1, y1 = line_bbox
    content_left = left_margin
    content_right = page_width - right_margin
    content_width = content_right - content_left
    line_width = x1 - x0
    # 居中
    if abs((x0 + x1) / 2 - (content_left + content_right) / 2) < content_width * 0.05 and \
       line_width < content_width * 0.85:
        return 'center'
    # 右对齐
    if abs(x1 - content_right) < content_width * 0.05 and x0 > content_left + content_width * 0.1:
        return 'right'
    # 两端对齐 (行宽接近内容宽度 且 非孤行)
    if line_width > content_width * 0.9 and abs(x0 - content_left) < 5:
        return 'justify'
    # 默认左对齐
    return 'left'
