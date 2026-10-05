"""
metadata_handler.py - PDF 元数据还原器
将 PDF 元数据 (metadata) 还原到 DOCX 文档属性:
  - title (标题)
  - author (作者)
  - subject (主题)
  - keywords (关键词)
  - creator (创建程序)
  - producer (生成器)
  - creationDate (创建日期)
  - modDate (修改日期)
  - category (类别)
  - comments (注释)
DOCX 核心属性 (core properties): title/author/subject/keywords/comments/category/created/modified
"""
from __future__ import annotations
from typing import Dict, Optional, Tuple
from datetime import datetime
import re


def parse_pdf_date(date_str: str) -> Optional[datetime]:
    """解析 PDF 日期格式 D:YYYYMMDDHHmmSS+TZ 为 datetime
    格式: D:20260904165307+08'00'
    """
    if not date_str:
        return None
    # 去掉 D: 前缀
    s = date_str.strip()
    if s.startswith('D:'):
        s = s[2:]
    # 提取 YYYYMMDDHHmmSS
    m = re.match(r'(\d{4})(\d{2})(\d{2})(\d{2})?(\d{2})?(\d{2})?', s)
    if not m:
        return None
    try:
        year = int(m.group(1))
        month = int(m.group(2))
        day = int(m.group(3))
        hour = int(m.group(4) or 0)
        minute = int(m.group(5) or 0)
        second = int(m.group(6) or 0)
        return datetime(year, month, day, hour, minute, second)
    except Exception:
        return None


def extract_pdf_metadata(pdf_path: str) -> Dict:
    """提取 PDF 元数据"""
    import fitz
    result = {}
    try:
        doc = fitz.open(pdf_path)
        meta = doc.metadata
        result['title'] = meta.get('title', '') or ''
        result['author'] = meta.get('author', '') or ''
        result['subject'] = meta.get('subject', '') or ''
        result['keywords'] = meta.get('keywords', '') or ''
        result['creator'] = meta.get('creator', '') or ''
        result['producer'] = meta.get('producer', '') or ''
        # 日期解析
        creation_date = parse_pdf_date(meta.get('creationDate', ''))
        mod_date = parse_pdf_date(meta.get('modDate', ''))
        if creation_date:
            result['created'] = creation_date
        if mod_date:
            result['modified'] = mod_date
        # 格式
        result['format'] = meta.get('format', '')
        # 页数
        result['page_count'] = len(doc)
        # 是否加密
        result['encryption'] = meta.get('encryption', None)
        doc.close()
    except Exception:
        pass
    return result


def apply_metadata_to_docx(doc, metadata: Dict):
    """将元数据应用到 DOCX 文档属性 (core properties)
    Args:
        doc: python-docx Document 对象
        metadata: extract_pdf_metadata 返回的 dict
    """
    try:
        cp = doc.core_properties
        if metadata.get('title'):
            cp.title = metadata['title']
        if metadata.get('author'):
            cp.author = metadata['author']
        if metadata.get('subject'):
            cp.subject = metadata['subject']
        if metadata.get('keywords'):
            cp.keywords = metadata['keywords']
        # comments
        # category
        if metadata.get('producer'):
            # 把 producer 放到 comments (DOCX 没有原生 producer 字段)
            existing_comments = cp.comments or ''
            producer_info = f'PDF Producer: {metadata["producer"]}'
            if producer_info not in existing_comments:
                cp.comments = (existing_comments + '\n' + producer_info).strip() if existing_comments else producer_info
        if metadata.get('creator'):
            # 把 creator 放到 category
            cp.category = metadata['creator']
        if metadata.get('created'):
            cp.created = metadata['created']
        if metadata.get('modified'):
            cp.modified = metadata['modified']
        # language
        # content_status
        # version
        # last_modified_by
        if metadata.get('author'):
            cp.last_modified_by = metadata['author']
    except Exception as e:
        pass


def get_metadata_summary(metadata: Dict) -> Dict:
    """获取元数据摘要 (用于报告)"""
    return {
        'has_title': bool(metadata.get('title')),
        'has_author': bool(metadata.get('author')),
        'has_subject': bool(metadata.get('subject')),
        'has_keywords': bool(metadata.get('keywords')),
        'has_dates': bool(metadata.get('created') or metadata.get('modified')),
        'page_count': metadata.get('page_count', 0),
        'producer': metadata.get('producer', '')[:50],
    }
