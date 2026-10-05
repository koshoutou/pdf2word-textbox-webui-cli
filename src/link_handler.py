"""
link_handler.py - 超链接/书签还原器
将 PDF 中的链接还原为 DOCX 超链接 (w:hyperlink) 和书签 (w:bookmarkStart/End)
- 网页链接 (URI): w:hyperlink r:id=...
- 内部跳转 (页面跳转): w:hyperlink w:anchor=bookmark_name
- 书签: w:bookmarkStart/w:bookmarkEnd (供内部跳转目标)
"""
from __future__ import annotations
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass, field
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


@dataclass
class LinkInfo:
    """链接信息"""
    bbox: Tuple[float, float, float, float]
    uri: str = ''
    target_page: int = -1    # 内部跳转目标页 (0-based)
    kind: str = 'uri'         # 'uri' / 'goto' / 'named'
    named_dest: str = ''
    text: str = ''            # 链接文字


class LinkHandler:
    """超链接处理器"""

    def __init__(self, docx_doc):
        self.doc = docx_doc
        self._bookmark_counter = 0
        self._rel_counter = 0
        # bookmark_name → page 映射 (供后续 goto 链接引用)
        self.page_bookmarks: Dict[int, str] = {}
        # rId 缓存, 避免重复创建
        self._url_rel_cache: Dict[str, str] = {}

    def _get_url_rid(self, url: str) -> str:
        """获取或创建 URL 的 rId (复用同URL的rId)
        注意: 必须用 is_external=True, 否则python-docx会把它当内部Part关系, 保存时报错
        """
        if url in self._url_rel_cache:
            return self._url_rel_cache[url]
        try:
            from docx.opc.constants import RELATIONSHIP_TYPE as RT
            # is_external=True 才能正确处理外部URL
            r_id = self.doc.part.relate_to(url, RT.HYPERLINK, is_external=True)
            self._url_rel_cache[url] = r_id
            return r_id
        except Exception as e:
            return ''

    def ensure_bookmark_for_page(self, page_no: int) -> str:
        """确保某页有一个书签, 返回书签名 (供内部跳转)
        注意: 实际插入书签需要在段落构建时调用 insert_bookmark
        """
        if page_no not in self.page_bookmarks:
            self._bookmark_counter += 1
            self.page_bookmarks[page_no] = f'_page_{page_no + 1}'
        return self.page_bookmarks[page_no]

    def insert_bookmark(self, paragraph, bookmark_name: str):
        """在段落开头插入书签"""
        self._bookmark_counter += 1
        bm_id = self._bookmark_counter + 1000
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

    def add_hyperlink(self, paragraph, url: str, text: str,
                      color: str = '0563C1', underline: bool = True):
        """给段落添加超链接 (外部URL)
        使用 LinkHandler 缓存的 rId, 避免 paragraph.part 不支持的问题
        """
        r_id = self._get_url_rid(url)
        if not r_id:
            # fallback: 纯文本 + 蓝色下划线
            try:
                run = paragraph.add_run(text)
                from docx.shared import RGBColor
                from docx.enum.text import WD_UNDERLINE
                run.font.color.rgb = RGBColor.from_string(color)
                run.font.underline = WD_UNDERLINE.SINGLE
            except Exception:
                paragraph.add_run(text)
            return None
        try:
            # 创建 w:hyperlink 元素
            hyperlink = OxmlElement('w:hyperlink')
            hyperlink.set(qn('r:id'), r_id)
            # 创建 run
            run = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            # 颜色
            if color:
                c = OxmlElement('w:color')
                c.set(qn('w:val'), color)
                rPr.append(c)
            # 下划线
            if underline:
                u = OxmlElement('w:u')
                u.set(qn('w:val'), 'single')
                rPr.append(u)
            run.append(rPr)
            # 文本
            t = OxmlElement('w:t')
            t.text = text
            t.set(qn('xml:space'), 'preserve')
            run.append(t)
            hyperlink.append(run)
            paragraph._element.append(hyperlink)
            return hyperlink
        except Exception as e:
            # fallback
            try:
                run = paragraph.add_run(text)
                from docx.shared import RGBColor
                from docx.enum.text import WD_UNDERLINE
                run.font.color.rgb = RGBColor.from_string(color)
                run.font.underline = WD_UNDERLINE.SINGLE
            except Exception:
                paragraph.add_run(text)
            return None

    def add_internal_link(self, paragraph, anchor_name: str, text: str,
                           color: str = '0563C1', underline: bool = True):
        """给段落添加内部跳转链接 (anchor)"""
        try:
            hyperlink = OxmlElement('w:hyperlink')
            hyperlink.set(qn('w:anchor'), anchor_name)
            run = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            if color:
                c = OxmlElement('w:color')
                c.set(qn('w:val'), color)
                rPr.append(c)
            if underline:
                u = OxmlElement('w:u')
                u.set(qn('w:val'), 'single')
                rPr.append(u)
            run.append(rPr)
            t = OxmlElement('w:t')
            t.text = text
            t.set(qn('xml:space'), 'preserve')
            run.append(t)
            hyperlink.append(run)
            paragraph._element.append(hyperlink)
            return hyperlink
        except Exception:
            run = paragraph.add_run(text)
            return run._element


def find_links_in_spans(page_data, link_list) -> List[Dict]:
    """匹配页面的链接与对应的 span 文本
    返回: [{'bbox', 'uri', 'target_page', 'text', 'kind'}]
    """
    result = []
    for link in link_list:
        lx0, ly0, lx1, ly1 = link.rect
        # 找到与链接 bbox 重叠的 spans
        link_text = ''
        for span in page_data.all_spans:
            sx0, sy0, sx1, sy1 = span.bbox
            cx = (sx0 + sx1) / 2
            cy = (sy0 + sy1) / 2
            if lx0 - 2 <= cx <= lx1 + 2 and ly0 - 2 <= cy <= ly1 + 2:
                link_text += span.text
        if link_text.strip() or link.kind in ('goto', 'named'):
            result.append({
                'bbox': link.rect,
                'uri': link.uri,
                'target_page': -1,
                'text': link_text,
                'kind': 'uri' if link.uri else 'goto',
                'named_dest': '',
            })
    return result


# URL 正则 (识别纯文本中的 URL)
import re as _re
URL_PATTERN = _re.compile(r'https?://[^\s\u4e00-\u9fff，。；：）】>）]+')
EMAIL_PATTERN = _re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')


def detect_urls_in_paragraph(para) -> List[Dict]:
    """检测段落文本中的URL和Email, 返回链接信息列表"""
    result = []
    for text, fmt in para.runs:
        # URL
        for m in URL_PATTERN.finditer(text):
            url = m.group(0).rstrip('.,;:)》】>')
            result.append({
                'text': url,
                'uri': url,
                'kind': 'uri',
            })
        # Email
        for m in EMAIL_PATTERN.finditer(text):
            email = m.group(0)
            result.append({
                'text': email,
                'uri': f'mailto:{email}',
                'kind': 'uri',
            })
    return result
