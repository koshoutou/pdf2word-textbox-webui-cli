"""
toc_builder.py - TOC 目录自动生成器
基于检测到的标题段 (is_heading) 生成 DOCX 目录:
  1. 收集所有标题段 (heading_level, text, page_no)
  2. 在文档开头插入 TOC 字段 (w:fldChar TOC)
  3. 为每个标题段添加书签 (w:bookmarkStart/End)
  4. TOC 字段在 Word 打开时自动渲染目录条目
支持多级目录 (1-9级)
"""
from __future__ import annotations
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass, field
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


@dataclass
class TOCEntry:
    """目录条目"""
    level: int            # 1-9
    title: str
    page_no: int = 0      # 目标页 (1-based), 0=未知
    bookmark_name: str = ''  # 书签名
    heading_text: str = ''    # 原标题文本


class TOCBuilder:
    """目录生成器"""

    def __init__(self, docx_doc):
        self.doc = docx_doc
        self._bookmark_counter = 0
        self._toc_entries: List[TOCEntry] = []
        # 标题文本 → bookmark_name 映射 (用于内部跳转)
        self._title_to_bookmark: Dict[str, str] = {}

    def collect_heading(self, level: int, title: str, page_no: int = 0) -> str:
        """收集一个标题, 返回书签名 (供后续插入 bookmark)
        Args:
            level: 标题级别 (1-9)
            title: 标题文本
            page_no: 1-based 页码 (0=未知)
        Returns:
            bookmark_name: 书签名 (用于内部跳转)
        """
        self._bookmark_counter += 1
        bm_name = f'_Toc{self._bookmark_counter:06d}'
        entry = TOCEntry(
            level=min(max(level, 1), 9),
            title=title.strip(),
            page_no=page_no,
            bookmark_name=bm_name,
            heading_text=title.strip(),
        )
        self._toc_entries.append(entry)
        self._title_to_bookmark[title.strip()] = bm_name
        return bm_name

    def get_bookmark_for_title(self, title: str) -> Optional[str]:
        """根据标题文本获取书签名"""
        return self._title_to_bookmark.get(title.strip())

    def insert_bookmark_at_paragraph(self, paragraph, bookmark_name: str):
        """在段落开头插入书签 (供 TOC 跳转)"""
        try:
            self._bookmark_counter += 1
            bm_id = self._bookmark_counter + 2000
            start = OxmlElement('w:bookmarkStart')
            start.set(qn('w:id'), str(bm_id))
            start.set(qn('w:name'), bookmark_name)
            end = OxmlElement('w:bookmarkEnd')
            end.set(qn('w:id'), str(bm_id))
            # 插入到段落第一个 run 之前, 没有则追加
            pPr = paragraph._element.find(qn('w:pPr'))
            if pPr is not None:
                pPr.addnext(start)
            else:
                paragraph._element.insert(0, start)
            paragraph._element.append(end)
        except Exception:
            pass

    def insert_toc_field(self, at_position: int = 0, max_level: int = 3,
                         title: str = '目  录'):
        """在文档指定位置插入 TOC 字段
        Args:
            at_position: 插入位置 (0=文档开头)
            max_level: 目录最大级别 (1-9)
            title: 目录标题文本
        Returns:
            插入的段落数 (标题 + TOC字段)
        """
        try:
            # 创建目录标题段落
            title_p = OxmlElement('w:p')
            # 标题样式
            pPr_title = OxmlElement('w:pPr')
            # 居中
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'center')
            pPr_title.append(jc)
            # 段后间距
            spacing = OxmlElement('w:spacing')
            spacing.set(qn('w:before'), '240')
            spacing.set(qn('w:after'), '240')
            pPr_title.append(spacing)
            title_p.append(pPr_title)
            # 标题 run
            title_run = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            # 加粗 + 大字号
            b = OxmlElement('w:b')
            rPr.append(b)
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '36')  # 18pt
            rPr.append(sz)
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), '36')
            rPr.append(szCs)
            title_run.append(rPr)
            t = OxmlElement('w:t')
            t.text = title
            t.set(qn('xml:space'), 'preserve')
            title_run.append(t)
            title_p.append(title_run)

            # 创建 TOC 字段段落
            toc_p = OxmlElement('w:p')
            pPr_toc = OxmlElement('w:pPr')
            toc_p.append(pPr_toc)
            # TOC 字段: begin + instrText + separate + end
            run1 = OxmlElement('w:r')
            fld_begin = OxmlElement('w:fldChar')
            fld_begin.set(qn('w:fldCharType'), 'begin')
            run1.append(fld_begin)
            toc_p.append(run1)
            # instrText
            run2 = OxmlElement('w:r')
            instr = OxmlElement('w:instrText')
            instr.set(qn('xml:space'), 'preserve')
            instr.text = f' TOC \\o "1-{max_level}" \\h \\z \\u '
            run2.append(instr)
            toc_p.append(run2)
            # separate
            run3 = OxmlElement('w:r')
            fld_sep = OxmlElement('w:fldChar')
            fld_sep.set(qn('w:fldCharType'), 'separate')
            run3.append(fld_sep)
            toc_p.append(run3)
            # 占位文本 (Word打开时会被实际目录替换)
            run4 = OxmlElement('w:r')
            t_placeholder = OxmlElement('w:t')
            t_placeholder.text = '右键此处选择"更新域"以生成目录...'
            t_placeholder.set(qn('xml:space'), 'preserve')
            run4.append(t_placeholder)
            toc_p.append(run4)
            # end
            run5 = OxmlElement('w:r')
            fld_end = OxmlElement('w:fldChar')
            fld_end.set(qn('w:fldCharType'), 'end')
            run5.append(fld_end)
            toc_p.append(run5)

            # 分页符 (目录后换页) - 仅当有标题时才加
            # 注意: 如果文档已有内容, TOC后分页符可能导致空页
            # 改为仅在TOC后加分页符, 不在TOC前加
            page_break_p = OxmlElement('w:p')
            r_pb = OxmlElement('w:r')
            br = OxmlElement('w:br')
            br.set(qn('w:type'), 'page')
            r_pb.append(br)
            page_break_p.append(r_pb)

            # 插入到文档开头 (body 的第一个元素之前)
            body = self.doc.element.body
            # 找到第一个 sectPr 之前的位置
            first_elem = body[0] if len(body) > 0 else None
            # 插入顺序: 标题 → TOC字段 → 分页符
            body.insert(0, page_break_p)
            body.insert(0, toc_p)
            body.insert(0, title_p)
            return 3  # 插入3个段落
        except Exception as e:
            return 0

    def apply_outline_level(self, paragraph, level: int):
        """给段落设置大纲级别 (Word导航窗格识别)
        让标题出现在Word的"导航窗格"中
        """
        try:
            pPr = paragraph._element.get_or_add_pPr()
            existing = pPr.find(qn('w:outlineLvl'))
            if existing is not None:
                pPr.remove(existing)
            outline = OxmlElement('w:outlineLvl')
            outline.set(qn('w:val'), str(min(max(level - 1, 0), 8)))
            pPr.append(outline)
        except Exception:
            pass

    def get_stats(self) -> Dict:
        """获取统计信息"""
        return {
            'total_entries': len(self._toc_entries),
            'by_level': {lvl: sum(1 for e in self._toc_entries if e.level == lvl)
                          for lvl in range(1, 10)},
        }
