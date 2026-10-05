"""
reporter.py - 转换质量报告生成器
对比原 PDF 与转换后的 DOCX, 输出质量报告:
  - 文本完整性 (字符数对比)
  - 表格数对比
  - 图片数对比
  - 字体覆盖统计
  - 格式覆盖率 (加粗/斜体/下划线/颜色)
  - 逐页段落统计
"""
from __future__ import annotations
import fitz
from docx import Document
from typing import Dict, List, Tuple
from collections import Counter
import os


class ConversionReporter:
    """转换质量报告"""

    def __init__(self, pdf_path: str, docx_path: str, verbose: bool = False,
                 page_range: Tuple[int, int] = None, page_list: List[int] = None):
        """
        Args:
            page_range: (start, end) 1-based 闭区间, 仅统计这些页
            page_list: 显式页码列表 (1-based), 优先于 page_range
        """
        self.pdf_path = pdf_path
        self.docx_path = docx_path
        self.verbose = verbose
        self.page_range = page_range
        self.page_list = page_list
        # 解析实际要统计的页 (0-based)
        self._target_pages = self._resolve_target_pages()

    def _resolve_target_pages(self) -> List[int]:
        """解析要统计的 PDF 页 (0-based)"""
        # 先获取总页数
        try:
            doc = fitz.open(self.pdf_path)
            total = len(doc)
            doc.close()
        except Exception:
            total = 0
        if self.page_list:
            return [p - 1 for p in self.page_list if 1 <= p <= total]
        if self.page_range:
            s, e = self.page_range
            s = max(1, s)
            e = min(total, e) if total else e
            return list(range(s - 1, e))
        return list(range(total)) if total else []

    def generate(self) -> Dict:
        """生成质量报告"""
        report = {
            'pdf_path': self.pdf_path,
            'docx_path': self.docx_path,
            'pdf_stats': {},
            'docx_stats': {},
            'comparison': {},
            'quality_score': 0,
            'issues': [],
        }
        # PDF 统计
        pdf_stats = self._analyze_pdf()
        report['pdf_stats'] = pdf_stats
        # DOCX 统计
        docx_stats = self._analyze_docx()
        report['docx_stats'] = docx_stats
        # 对比
        comp = report['comparison']
        comp['text_coverage'] = (docx_stats['char_count'] / pdf_stats['char_count']
                                  if pdf_stats['char_count'] > 0 else 0)
        comp['tables_coverage'] = (docx_stats['table_count'] / pdf_stats['table_count']
                                    if pdf_stats['table_count'] > 0 else 1.0)
        comp['images_coverage'] = (docx_stats['image_count'] / pdf_stats['image_count']
                                    if pdf_stats['image_count'] > 0 else 1.0)
        # 质量评分 (加权)
        score = 0
        score += min(comp['text_coverage'], 1.0) * 50      # 文本 50%
        score += min(comp['tables_coverage'], 1.0) * 25     # 表格 25%
        score += min(comp['images_coverage'], 1.0) * 15     # 图片 15%
        score += min(docx_stats['format_coverage'], 1.0) * 10  # 格式 10%
        report['quality_score'] = round(score, 1)
        # 问题检测
        if comp['text_coverage'] < 0.95:
            report['issues'].append(f"文本覆盖率偏低: {comp['text_coverage']:.1%}")
        if comp['tables_coverage'] < 0.9 and pdf_stats['table_count'] > 0:
            report['issues'].append(f"表格丢失: PDF {pdf_stats['table_count']} → DOCX {docx_stats['table_count']}")
        if comp['images_coverage'] < 0.9 and pdf_stats['image_count'] > 0:
            report['issues'].append(f"图片丢失: PDF {pdf_stats['image_count']} → DOCX {docx_stats['image_count']}")
        return report

    def _analyze_pdf(self) -> Dict:
        """统计 PDF (仅目标页)"""
        stats = {
            'page_count': 0, 'total_page_count': 0, 'char_count': 0,
            'table_count': 0, 'image_count': 0, 'fonts': set(), 'colors': set(),
            'bold_spans': 0, 'total_spans': 0,
            'pages_analyzed': 0,
        }
        import pdfplumber
        target = self._target_pages
        stats['pages_analyzed'] = len(target)
        doc = fitz.open(self.pdf_path)
        stats['total_page_count'] = len(doc)
        stats['page_count'] = len(target)
        for pno in target:
            if pno >= len(doc):
                continue
            page = doc[pno]
            d = page.get_text('dict')
            for b in d.get('blocks', []):
                if b.get('type') == 0:
                    for l in b.get('lines', []):
                        for s in l.get('spans', []):
                            stats['char_count'] += len(s.get('text', ''))
                            stats['fonts'].add(s.get('font', ''))
                            stats['colors'].add(s.get('color', 0))
                            stats['total_spans'] += 1
                            if s.get('flags', 0) & 16:
                                stats['bold_spans'] += 1
                elif b.get('type') == 1:
                    stats['image_count'] += 1
            stats['image_count'] += len(page.get_images(full=True))
        doc.close()
        # 表格统计 (仅目标页)
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for pno in target:
                    if pno < len(pdf.pages):
                        stats['table_count'] += len(pdf.pages[pno].find_tables())
        except Exception:
            pass
        stats['fonts'] = len(stats['fonts'])
        stats['colors'] = len(stats['colors'])
        return stats

    def _analyze_docx(self) -> Dict:
        """统计 DOCX"""
        stats = {
            'paragraph_count': 0, 'char_count': 0, 'table_count': 0,
            'image_count': 0, 'bold_runs': 0, 'italic_runs': 0,
            'underline_runs': 0, 'colored_runs': 0, 'total_runs': 0,
            'fonts': set(), 'format_coverage': 0,
        }
        doc = Document(self.docx_path)
        stats['paragraph_count'] = len(doc.paragraphs)
        stats['table_count'] = len(doc.tables)
        # 图片
        import zipfile
        z = zipfile.ZipFile(self.docx_path)
        stats['image_count'] = len([n for n in z.namelist() if n.startswith('word/media/')])
        z.close()
        # runs 统计
        for p in doc.paragraphs:
            for r in p.runs:
                stats['total_runs'] += 1
                stats['char_count'] += len(r.text or '')
                if r.font.bold:
                    stats['bold_runs'] += 1
                if r.font.italic:
                    stats['italic_runs'] += 1
                if r.font.underline:
                    stats['underline_runs'] += 1
                try:
                    if r.font.color and r.font.color.rgb:
                        stats['colored_runs'] += 1
                except Exception:
                    pass
                rpr = r._element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
                if rpr is not None:
                    rf = rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
                    if rf is not None:
                        ea = rf.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
                        if ea:
                            stats['fonts'].add(ea)
        # 表格内 runs
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            stats['total_runs'] += 1
                            stats['char_count'] += len(r.text or '')
        stats['fonts'] = len(stats['fonts'])
        # 格式覆盖率
        if stats['total_runs'] > 0:
            covered = stats['bold_runs'] + stats['italic_runs'] + stats['underline_runs'] + stats['colored_runs']
            stats['format_coverage'] = min(covered / stats['total_runs'], 1.0)
        return stats

    def print_report(self, report: Dict):
        """打印报告"""
        print('=' * 60)
        print('  PDF→DOCX 转换质量报告')
        print('=' * 60)
        print(f'PDF:  {report["pdf_path"]}')
        print(f'DOCX: {report["docx_path"]}')
        print()
        print('--- PDF 统计 ---')
        ps = report['pdf_stats']
        if ps.get('pages_analyzed', 0) != ps.get('total_page_count', 0):
            print(f'  统计页数: {ps["page_count"]} / 总 {ps["total_page_count"]} 页 (分页转换)')
        else:
            print(f'  页数: {ps["page_count"]}')
        print(f'  字符数: {ps["char_count"]}')
        print(f'  表格数: {ps["table_count"]}')
        print(f'  图片数: {ps["image_count"]}')
        print(f'  字体种类: {ps["fonts"]}')
        print(f'  颜色数: {ps["colors"]}')
        print()
        print('--- DOCX 统计 ---')
        ds = report['docx_stats']
        print(f'  段落数: {ds["paragraph_count"]}')
        print(f'  字符数: {ds["char_count"]}')
        print(f'  表格数: {ds["table_count"]}')
        print(f'  图片数: {ds["image_count"]}')
        print(f'  字体种类: {ds["fonts"]}')
        print(f'  加粗 run: {ds["bold_runs"]}, 斜体: {ds["italic_runs"]}, 下划线: {ds["underline_runs"]}')
        print()
        print('--- 对比 ---')
        comp = report['comparison']
        print(f'  文本覆盖率: {comp["text_coverage"]:.1%}')
        print(f'  表格覆盖率: {comp["tables_coverage"]:.1%}')
        print(f'  图片覆盖率: {comp["images_coverage"]:.1%}')
        print(f'  格式覆盖率: {comp.get("format_coverage", 0):.1%}' if 'format_coverage' in comp else '')
        print()
        print(f'  ★ 综合质量评分: {report["quality_score"]}/100')
        if report['issues']:
            print('\n--- 问题 ---')
            for iss in report['issues']:
                print(f'  ⚠ {iss}')
        else:
            print('\n  ✓ 无明显问题')
        print('=' * 60)
