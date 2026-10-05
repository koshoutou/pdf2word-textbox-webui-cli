"""
list_detector.py - 列表/项目符号检测器
检测段落开头是否为列表标记, 并归类:
  - 项目符号: ● ○ ■ ◆ □ ☆ ※ · • ‣ ⁃ ▪ ▫ - * +
  - 阿拉伯数字: 1. 1) 1、 (1) 1)
  - 中文数字: 一、 (一) 第一 ①
  - 字母: a. A. a) A) (a) (A)
  - 罗马数字: i. I. ii. II.
策略: 通过正则匹配段落首非空白文本
"""
from __future__ import annotations
import re
from typing import Optional, Tuple, List
from analyzer import ParagraphData


# 列表标记正则 (按优先级)
LIST_PATTERNS = [
    # 项目符号
    (re.compile(r'^[\s]*[●○■◆□☆※·•‣⁃▪▫][\s]*'), 'bullet', 0),
    (re.compile(r'^[\s]*[●○■◆□☆※·•‣⁃▪▫][\s]+'), 'bullet', 0),
    # 中文数字 一、二、
    (re.compile(r'^[\s]*([一二三四五六七八九十]+)[、．\.][\s]*'), 'chinese', 0),
    # (一)(二) 括号中文
    (re.compile(r'^[\s]*[（(]([一二三四五六七八九十]+)[)）][\s]*'), 'chinese_paren', 0),
    # 第X条/章/项
    (re.compile(r'^[\s]*第([一二三四五六七八九十百千]+|[0-9]+)[章节条款款项][\s]*'), 'chinese_ordinal', 0),
    # ① ② ③ 圈数字
    (re.compile(r'^[\s]*[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮][\s]*'), 'circled', 0),
    # 1. 2. (阿拉伯数字+点)
    (re.compile(r'^[\s]*(\d+)[\.．][\s]*'), 'decimal', 0),
    # 1) 2) (阿拉伯数字+右括号)
    (re.compile(r'^[\s]*(\d+)[\)）][\s]*'), 'decimal_paren', 0),
    # (1) (2) 全括号阿拉伯
    (re.compile(r'^[\s]*[（(](\d+)[)）][\s]*'), 'decimal_fullparen', 0),
    # a. b. / A. B. 字母
    (re.compile(r'^[\s]*([a-z])[\.．][\s]*'), 'lowerLetter', 0),
    (re.compile(r'^[\s]*([A-Z])[\.．][\s]*'), 'upperLetter', 0),
    # a) b) / A) B)
    (re.compile(r'^[\s]*([a-z])[\)）][\s]*'), 'lowerLetter_paren', 0),
    (re.compile(r'^[\s]*([A-Z])[\)）][\s]*'), 'upperLetter_paren', 0),
    # i. ii. iii. 罗马数字
    (re.compile(r'^[\s]*(i{1,3}|iv|v|vi{0,3}|ix|x)[\.．][\s]*'), 'lowerRoman', 0),
    (re.compile(r'^[\s]*(I{1,3}|IV|V|VI{0,3}|IX|X)[\.．][\s]*'), 'upperRoman', 0),
    # - * + 减号/星号/加号 项目符号
    (re.compile(r'^[\s]*[-*+][\s]+'), 'bullet_dash', 0),
]

# 嵌套层级估算: 缩进字符数 / 单字符宽度
NEST_INDENT_PT = 14.0  # 每级缩进约 14pt


def detect_list_marker(para: ParagraphData) -> Optional[Tuple[str, int, str]]:
    """检测段落是否为列表项
    Returns:
        (list_type, level, marker_text) 或 None
    注意: 标题段(已被analyzer标记is_heading)不作为列表项, 避免章节标题被误剥离
    """
    if not para.runs:
        return None
    # ★ 标题段不作为列表项 (避免"第一章"被剥离)
    if para.is_heading:
        return None
    first_text = para.runs[0][0]
    # 取前 30 字符用于匹配
    sample = first_text[:30]
    for pat, ltype, level in LIST_PATTERNS:
        m = pat.match(sample)
        if m:
            # 估算嵌套层级 (基于左缩进)
            actual_level = _estimate_level(para.left_indent)
            # 提取标记文本
            marker = m.group(0).strip()
            return (ltype, actual_level, marker)
    return None


def _estimate_level(left_indent: float) -> int:
    """根据左缩进估算列表层级"""
    if left_indent < NEST_INDENT_PT * 0.5:
        return 0
    return min(int(left_indent / NEST_INDENT_PT), 4)


def detect_lists_in_paragraphs(paragraphs: List[ParagraphData]):
    """对一组段落执行列表检测, 标记同组列表项 (共享 list_group_id)"""
    current_group_id = 0
    prev_list_type = None
    prev_list_level = -1
    group_counter = 0
    for p in paragraphs:
        result = detect_list_marker(p)
        if result:
            ltype, level, marker = result
            # 同类型同层级视为同组 (连续列表)
            if ltype == prev_list_type and level == prev_list_level:
                pass  # 同组
            else:
                group_counter += 1
                current_group_id = group_counter
            p.is_list_item = True
            p.list_type = ltype
            p.list_level = level
            p.list_marker_text = marker
            p.list_group_id = current_group_id
            # 从 runs 中移除标记文本 (避免重复显示)
            _strip_marker_from_runs(p, marker)
            prev_list_type = ltype
            prev_list_level = level
        else:
            # 非列表项: 中断连续性, 但不重置 group (允许列表中夹非列表段)
            if p.runs and ''.join(t for t, _ in p.runs).strip():
                # 检查是否是上段的延续 (短文本)
                total_text = ''.join(t for t, _ in p.runs)
                if len(total_text) > 50:
                    prev_list_type = None  # 长段中断列表


def _strip_marker_from_runs(para: ParagraphData, marker: str):
    """从段落 runs 中移除列表标记前缀 (避免在 DOCX 中重复显示)
    DOCX 列表会自动渲染标记
    """
    if not marker or not para.runs:
        return
    remaining = marker
    new_runs = []
    marker_stripped = False
    for text, fmt in para.runs:
        if not marker_stripped:
            # 在当前 run 中找 marker
            if remaining:
                stripped = _strip_prefix(text, remaining)
                if stripped is not None:
                    new_text, remaining = stripped
                    if new_text:
                        new_runs.append((new_text, fmt))
                    if not remaining:
                        marker_stripped = True
                    continue
                else:
                    # marker 跨 run, 保留原 run (罕见情况)
                    new_runs.append((text, fmt))
            else:
                new_runs.append((text, fmt))
        else:
            new_runs.append((text, fmt))
    para.runs = new_runs


def _strip_prefix(text: str, prefix: str) -> Optional[Tuple[str, str]]:
    """从 text 开头剥离 prefix, 返回 (剩余text, 剩余prefix)
    如果 text 不以 prefix 开头, 返回 None
    """
    if not prefix:
        return (text, '')
    if text.startswith(prefix):
        return (text[len(prefix):], '')
    # prefix 是 text 的前缀的一部分
    if len(text) < len(prefix) and prefix.startswith(text):
        return ('', prefix[len(text):])
    # 部分重叠
    for i in range(min(len(text), len(prefix)), 0, -1):
        if text[:i] == prefix[:i]:
            return (text[i:], prefix[i:])
    return None
