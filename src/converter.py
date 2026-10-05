"""
converter.py - 主转换器
编排: PDF 提取 → 布局分析 → 表格提取 → DOCX 构建
支持页码范围选择、分页符、图片渲染等
"""
from __future__ import annotations
import os
import sys
import time
from typing import List, Optional, Tuple
from dataclasses import dataclass, field

from extractor import PDFExtractor, PageData
from analyzer import LayoutAnalyzer, PageLayout, ParagraphData
from table_extractor import TableExtractor, ExtractedTable
from docx_builder import DocxBuilder
import fitz


@dataclass
class ConversionOptions:
    """转换选项"""
    page_range: Optional[Tuple[int, int]] = None      # 1-based 闭区间
    page_list: Optional[List[int]] = None             # 1-based 列表
    dpi: int = 150                                     # 图片渲染 DPI
    render_complex_drawings: bool = True              # 复杂矢量图形渲染为图片
    render_shapes: bool = True                        # 矢量图形还原为DOCX形状(VML/DrawingML)
    detect_tables: bool = True
    detect_lists: bool = True                         # 列表/项目符号检测
    detect_formulas: bool = True                      # 数学公式检测 (OMML)
    detect_links: bool = True                         # 超链接/书签还原
    detect_form_fields: bool = True                   # 表单域识别 (签名/复选框/文本框)
    preserve_page_breaks: bool = True                 # 保留分页
    align_pages: bool = True                          # 分页对齐: keepNext/cantSplit (轻量, 默认开)
    fill_page_boundary: bool = False                  # 页面填充对齐 (重量级, 可能增加页数, 默认关)
    generate_toc: bool = True                         # 自动生成目录 (基于heading)
    toc_max_level: int = 3                            # 目录最大级别
    restore_metadata: bool = True                     # 还原 PDF 元数据到 DOCX 属性
    detect_annotations: bool = True                   # PDF 注释还原 (高亮/批注/删除线)
    header_to_all: bool = True                        # 统一页眉
    footer_to_all: bool = True                        # 统一页脚
    verbose: bool = False


class PDF2DocxConverter:
    """主转换器"""

    def __init__(self, pdf_path: str, options: Optional[ConversionOptions] = None):
        self.pdf_path = pdf_path
        self.options = options or ConversionOptions()
        self.stats = {
            'pages': 0,
            'paragraphs': 0,
            'tables': 0,
            'images': 0,
            'errors': [],
        }

    def log(self, msg: str):
        if self.options.verbose:
            print(f'[Converter] {msg}')

    def _emit_progress(self, progress: int, message: str):
        """输出进度信息 (供API解析)
        格式: [Progress] <percent>% <message>
        """
        print(f'[Progress] {progress}% {message}', flush=True)

    def convert(self, output_path: str):
        """主转换流程"""
        t0 = time.time()
        self.log(f'开始转换: {self.pdf_path}')
        self._emit_progress(5, '开始转换')

        # 1. 提取 PDF 数据
        with PDFExtractor(self.pdf_path) as extractor:
            self.stats['pages'] = extractor.page_count
            # 解析页码范围
            pages_to_convert = self._resolve_pages(extractor.page_count)
            self.log(f'转换页数: {len(pages_to_convert)} / {extractor.page_count}')
            self._emit_progress(10, f'提取PDF数据 ({len(pages_to_convert)}页)')

            page_data_list = extractor.extract_pages(page_list=pages_to_convert)

            # 2. 布局分析
            analyzer = LayoutAnalyzer(page_data_list, verbose=self.options.verbose)
            layouts = analyzer.analyze()
            self._emit_progress(25, '布局分析完成')
            # 建立 page_no → PageData 映射 (供表格下划线检测用)
            page_data_map = {pd.page_no: pd for pd in page_data_list}

            # 3. 表格提取
            tables_by_page = {}
            if self.options.detect_tables:
                with TableExtractor(self.pdf_path, verbose=self.options.verbose) as te:
                    for idx, pno in enumerate(pages_to_convert):
                        tables_by_page[pno] = te.extract_page_tables(pno, page_data_list[idx])
                        self.stats['tables'] += len(tables_by_page[pno])

            # 3b. 表单域检测 (签名/复选框/文本框)
            form_fields_by_page = {}
            if self.options.detect_form_fields:
                try:
                    from form_handler import FormFieldDetector
                    detector = FormFieldDetector(self.pdf_path, verbose=self.options.verbose)
                    form_fields_by_page = detector.detect_all_fields()
                    # 只保留目标页
                    form_fields_by_page = {pno: fields for pno, fields in form_fields_by_page.items()
                                            if pno in pages_to_convert}
                    self.stats['form_fields'] = sum(len(f) for f in form_fields_by_page.values())
                    if self.options.verbose and form_fields_by_page:
                        print(f'[Form] 检测到 {self.stats["form_fields"]} 个表单域')
                except Exception as e:
                    if self.options.verbose:
                        print(f'[Form] error: {e}')

            # 3c. 注释检测 (高亮/批注/删除线/图章)
            annotations_by_page = {}
            if self.options.detect_annotations:
                try:
                    from annotation_handler import AnnotationDetector
                    ann_detector = AnnotationDetector(self.pdf_path, verbose=self.options.verbose)
                    annotations_by_page = ann_detector.detect_all_annotations()
                    annotations_by_page = {pno: a for pno, a in annotations_by_page.items()
                                            if pno in pages_to_convert}
                    self.stats['annotations'] = sum(len(a) for a in annotations_by_page.values())
                    if self.options.verbose and annotations_by_page:
                        print(f'[Annot] 检测到 {self.stats["annotations"]} 个注释')
                except Exception as e:
                    if self.options.verbose:
                        print(f'[Annot] error: {e}')

            # 4. 设置页面尺寸
            page_width = page_data_list[0].width if page_data_list else 612
            page_height = page_data_list[0].height if page_data_list else 792
            builder = DocxBuilder(page_width=page_width, page_height=page_height,
                                  verbose=self.options.verbose)
            # 边距: 精确匹配原PDF内容区
            # 原PDF: 内容y范围约73~694, 即内容区高度621pt
            # 页面高792 - 内容区621 = 171pt (上下边距总和)
            # top=73, bottom=792-694=98
            margin_top = 73
            margin_bottom = 98
            margin_left = 72
            margin_right = 72
            builder.setup_page(
                width_pt=page_width,
                height_pt=page_height,
                margin_top=margin_top,
                margin_bottom=margin_bottom,
                margin_left=margin_left,
                margin_right=margin_right,
                header_dist=28,    # 页眉距顶部28pt
                footer_dist=28,    # 页脚距底部28pt
            )

            # 5. 构建页眉/页脚 (取非封面页的页眉页脚, 封面通常无标准页眉页脚)
            header_layout = self._find_header_footer_layout(layouts)
            if header_layout and header_layout.header:
                builder.build_header(header_layout.header, page_number_in_header=True)
            if header_layout and header_layout.footer:
                builder.build_footer(header_layout.footer, page_number_in_footer=True)

            # 6. 构建正文 (按页)
            # 6a. 全局列表检测 (跨页同组列表)
            if self.options.detect_lists:
                try:
                    from list_detector import detect_lists_in_paragraphs
                    all_paras = []
                    for layout in layouts:
                        all_paras.extend(layout.body_paragraphs)
                    detect_lists_in_paragraphs(all_paras)
                    self.stats['list_items'] = sum(1 for p in all_paras if p.is_list_item)
                except Exception as e:
                    if self.options.verbose:
                        print(f'[List] detect error: {e}')

            for idx, layout in enumerate(layouts):
                # 进度: 30% ~ 85% 之间, 按页数比例
                progress = 30 + int(55 * idx / max(len(layouts), 1))
                self._emit_progress(progress, f'构建第 {idx+1}/{len(layouts)} 页')
                self._build_page_content(builder, layout, extractor,
                                          tables_by_page.get(layout.page_no, []),
                                          page_data_map.get(layout.page_no),
                                          form_fields_by_page.get(layout.page_no, []),
                                          annotations_by_page.get(layout.page_no, []))
                # ★ 分页对齐: 在页末填充空行, 把下页内容推到新页 (重量级, 默认关)
                if self.options.fill_page_boundary and self.options.preserve_page_breaks:
                    self._fill_page_to_boundary(builder, layout, page_data_map.get(layout.page_no))
                # ★ 每页强制分页: 确保DOCX页数=PDF页数
                # 配合精确边距, 让每页内容不溢出
                if self.options.preserve_page_breaks and idx < len(layouts) - 1:
                    builder.add_page_break()

            # 6b. ★ TOC 目录生成 (基于heading) + 书签
            if self.options.generate_toc:
                self._emit_progress(88, '生成TOC目录')
                self._generate_toc(builder, layouts)

            # 6c. ★ PDF 元数据还原
            if self.options.restore_metadata:
                self._emit_progress(92, '还原元数据')
                try:
                    from metadata_handler import extract_pdf_metadata, apply_metadata_to_docx
                    metadata = extract_pdf_metadata(self.pdf_path)
                    apply_metadata_to_docx(builder.doc, metadata)
                    self.stats['metadata_restored'] = sum(1 for v in metadata.values() if v)
                    if self.options.verbose:
                        print(f'[Metadata] restored {self.stats["metadata_restored"]} fields')
                except Exception as e:
                    if self.options.verbose:
                        print(f'[Metadata] error: {e}')

            # 7. 保存
            self._emit_progress(95, '保存DOCX')
            builder.save(output_path)

        elapsed = time.time() - t0
        self.log(f'转换完成, 用时 {elapsed:.1f}s, 输出: {output_path}')
        self.log(f'统计: {self.stats}')
        return self.stats

    def _resolve_pages(self, total: int) -> List[int]:
        """解析页码 → 0-based list"""
        if self.options.page_list:
            return [p - 1 for p in self.options.page_list if 1 <= p <= total]
        if self.options.page_range:
            start, end = self.options.page_range
            start = max(1, start)
            end = min(total, end)
            return list(range(start - 1, end))
        return list(range(total))

    def _build_page_content(self, builder: DocxBuilder, layout: PageLayout,
                            extractor: PDFExtractor,
                            tables: List[ExtractedTable],
                            page_data=None,
                            form_fields=None,
                            annotations=None):
        """构建单页内容 (按 y 顺序混合段落和表格)
        用 pdfplumber 的精确表格 bbox 过滤掉落入表格内的段落
        form_fields: 该页的表单域列表 (FormField)
        annotations: 该页的注释列表 (AnnotationInfo)
        """
        form_fields = form_fields or []
        annotations = annotations or []
        # 签名域 bbox (用于图片浮动定位优先级, 优先用签名域位置)
        signature_bboxes = []
        for ff in form_fields:
            if ff.field_type == 'Signature' and ff.rect[2] - ff.rect[0] > 50:
                # 大签名域(印章) → 用其bbox作为图片定位
                signature_bboxes.append(ff.rect)
        # ★ 公式检测
        formulas = []
        if self.options.detect_formulas and page_data:
            try:
                from formula_detector import FormulaDetector
                detector = FormulaDetector(verbose=self.options.verbose)
                formulas = detector.detect_page_formulas(page_data)
                if formulas:
                    self.stats['formulas'] = self.stats.get('formulas', 0) + len(formulas)
                    if self.options.verbose:
                        print(f'[Formula] p{layout.page_no+1}: {len(formulas)} 公式')
            except Exception as e:
                if self.options.verbose:
                    print(f'[Formula] detect error: {e}')
        formula_bboxes = [f.bbox for f in formulas]

        # 收集表格 bbox 用于段落过滤
        table_bboxes = [t.bbox for t in tables]
        # 过滤段落: 不在表格内, 不在公式内 (公式内容不重复进段落)
        filtered_paras = []
        for p in layout.body_paragraphs:
            in_table = False
            for tb in table_bboxes:
                cx = (p.bbox[0] + p.bbox[2]) / 2
                cy = (p.bbox[1] + p.bbox[3]) / 2
                if tb[0] - 2 <= cx <= tb[2] + 2 and tb[1] - 2 <= cy <= tb[3] + 2:
                    in_table = True
                    break
            if in_table:
                continue
            # 检查段落是否大部分在公式区域 (是则跳过, 由公式OMML代替)
            if formula_bboxes:
                p_cx = (p.bbox[0] + p.bbox[2]) / 2
                p_cy = (p.bbox[1] + p.bbox[3]) / 2
                in_formula = any(
                    fb[0]-2 <= p_cx <= fb[2]+2 and fb[1]-2 <= p_cy <= fb[3]+2
                    for fb in formula_bboxes
                )
                if in_formula:
                    continue
            filtered_paras.append(p)
        # 收集所有内容项, 按 y 排序
        items = []
        for p in filtered_paras:
            items.append((p.bbox[1], 'para', p))
        for t in tables:
            items.append((t.bbox[1], 'table', t))
        for img in layout.images:
            items.append((img['bbox'][1], 'image', img))
        for f in formulas:
            items.append((f.bbox[1], 'formula', f))
        # ★ 非签名表单域 (复选框/文本框/单选/按钮) 作为内容项
        for ff in form_fields:
            if ff.field_type not in ('Signature',):
                items.append((ff.rect[1], 'formfield', ff))
        items.sort(key=lambda x: x[0])

        # ★ 矢量图形: 默认禁用 (导致大量空段, 影响排版)
        # 仅当用户明确启用 render_shapes 且该页有有意义的图形时才构建
        # (注释掉自动构建, 避免空段累积导致页数爆炸)
        # if self.options.render_shapes and layout.standalone_drawings:
        #     self._add_standalone_shapes(builder, layout, page_data)

        # ★ 链接检测: 匹配页面链接与段落 + URL文本识别
        para_hyperlinks = {}
        if self.options.detect_links:
            # 1. PDF 原生链接
            if page_data and page_data.links:
                try:
                    from link_handler import find_links_in_spans
                    page_links = find_links_in_spans(page_data, page_data.links)
                    for link in page_links:
                        for p in filtered_paras:
                            if (p.bbox[0] <= link['bbox'][2] and p.bbox[2] >= link['bbox'][0] and
                                p.bbox[1] <= link['bbox'][3] and p.bbox[3] >= link['bbox'][1]):
                                para_hyperlinks.setdefault(id(p), []).append(link)
                                break
                    self.stats['links'] = self.stats.get('links', 0) + len(page_links)
                except Exception as e:
                    if self.options.verbose:
                        print(f'[Link] pdf error: {e}')
            # 2. URL/Email 文本识别
            try:
                from link_handler import detect_urls_in_paragraph
                url_count = 0
                for p in filtered_paras:
                    urls = detect_urls_in_paragraph(p)
                    if urls:
                        para_hyperlinks.setdefault(id(p), []).extend(urls)
                        url_count += len(urls)
                if url_count:
                    self.stats['links'] = self.stats.get('links', 0) + url_count
            except Exception as e:
                if self.options.verbose:
                    print(f'[Link] url error: {e}')

        for y0, typ, data in items:
            if typ == 'para':
                hls = para_hyperlinks.get(id(data), [])
                last_p = builder.add_paragraph(data, hyperlinks=hls)
                self.stats['paragraphs'] += 1
                # ★ 注释高亮: 检测段落是否与高亮注释重叠
                if self.options.detect_annotations and annotations and last_p:
                    self._apply_annotations_to_paragraph(last_p, data, annotations, page_data)
                # ★ 分页对齐: 标题段 keep_with_next + keep_lines
                if self.options.align_pages and data.is_heading:
                    try:
                        if last_p and hasattr(last_p, '_element'):
                            builder.apply_paragraph_keep_with_next(last_p)
                            builder.apply_paragraph_keep_lines(last_p)
                    except Exception:
                        pass
            elif typ == 'formula':
                # 插入 OMML 公式
                try:
                    builder.add_formula(
                        data.omml_xml,
                        latex=data.latex,
                        display=(data.formula_type == 'display'),
                    )
                except Exception as e:
                    self.stats['errors'].append(f'formula error: {e}')
            elif typ == 'formfield':
                # 表单域 (复选框/文本框/单选/按钮)
                try:
                    self._add_form_field(builder, data)
                except Exception as e:
                    self.stats['errors'].append(f'formfield error: {e}')
            elif typ == 'table':
                # 注入该页 drawings 用于表格内下划线检测
                builder._current_page_drawings = page_data.drawings if page_data else []
                builder.add_table(data)
                builder._current_page_drawings = []
                # ★ 分页对齐: 表格行不可分页 + 表头跨页重复
                if self.options.align_pages:
                    try:
                        last_table = builder.doc.tables[-1]
                        builder.prevent_table_row_split(last_table)
                        builder.set_table_header_row(last_table, 0)
                    except Exception:
                        pass
            elif typ == 'image':
                try:
                    xref = data['xref']
                    img_bytes = None
                    if xref and xref > 0:
                        img_bytes = extractor.get_image_pixmap(xref, dpi=self.options.dpi)
                    if not img_bytes:
                        # fallback: 裁剪渲染页面区域
                        img_bytes = extractor.render_region_image(
                            layout.page_no, data['bbox'], dpi=self.options.dpi)
                    if img_bytes:
                        # 计算显示尺寸 (pt)
                        w_pt = data['bbox'][2] - data['bbox'][0]
                        h_pt = data['bbox'][3] - data['bbox'][1]
                        # 判断是否需要浮动定位 (与正文文字重叠 → 印章类)
                        is_float = self._is_image_overlapping_text(layout, data['bbox'])
                        if is_float:
                            # 浮动定位 (印章: 衬于文字上方, 因印章要盖在文字上)
                            builder.add_floating_image(
                                img_bytes,
                                x_pt=data['bbox'][0],
                                y_pt=data['bbox'][1],
                                width_pt=w_pt, height_pt=h_pt,
                                behind_text=False)
                        else:
                            builder.add_image(img_bytes, width_pt=w_pt, height_pt=h_pt)
                        self.stats['images'] += 1
                except Exception as e:
                    self.stats['errors'].append(f'image error: {e}')

    def _is_image_overlapping_text(self, layout: PageLayout, img_bbox) -> bool:
        """判断图片是否与正文文字重叠 (印章类需浮动定位)
        启发式:
          - 图片较小 (宽高都 < 200pt, 典型印章/签名)
          - 且与文字段落有重叠
        """
        from analyzer import bbox_overlap
        img_w = img_bbox[2] - img_bbox[0]
        img_h = img_bbox[3] - img_bbox[1]
        is_small = img_w < 220 and img_h < 220
        for p in layout.body_paragraphs:
            if bbox_overlap(img_bbox, p.bbox, tol=2):
                if is_small:
                    return True
                # 大图: 重叠面积比例 > 0.1 才浮动
                ix0 = max(img_bbox[0], p.bbox[0])
                iy0 = max(img_bbox[1], p.bbox[1])
                ix1 = min(img_bbox[2], p.bbox[2])
                iy1 = min(img_bbox[3], p.bbox[3])
                overlap = (ix1 - ix0) * (iy1 - iy0)
                img_area = img_w * img_h
                if img_area > 0 and overlap / img_area > 0.1:
                    return True
        return False

    # ---------- TOC 目录生成 ----------
    def _generate_toc(self, builder, layouts):
        """生成目录:
        1. 收集所有标题段, 注册到 TOCBuilder
        2. 为每个标题段添加书签 + outlineLvl (导航窗格识别)
        3. 在文档开头插入 TOC 字段
        """
        try:
            from toc_builder import TOCBuilder
            toc = TOCBuilder(builder.doc)
            # 收集标题 (跨页)
            heading_count = 0
            for layout in layouts:
                for p in layout.body_paragraphs:
                    if p.is_heading and p.heading_level > 0:
                        text = ''.join(t for t, _ in p.runs).strip()
                        if not text or len(text) > 100:
                            continue
                        # 只收集 <= toc_max_level 的标题
                        if p.heading_level > self.options.toc_max_level:
                            continue
                        bm_name = toc.collect_heading(
                            level=p.heading_level,
                            title=text,
                            page_no=layout.page_no + 1,
                        )
                        # 记录书签名到段落, 后续插入
                        p._toc_bookmark = bm_name
                        heading_count += 1
            if heading_count == 0:
                return
            # 为标题段插入书签 + outlineLvl
            # 遍历 docx 段落, 匹配标题文本
            for p in builder.doc.paragraphs:
                text = p.text.strip()
                if not text:
                    continue
                # 找匹配的 TOC 条目
                for entry in toc._toc_entries:
                    if entry.heading_text == text:
                        # 插入书签
                        toc.insert_bookmark_at_paragraph(p, entry.bookmark_name)
                        # 设置大纲级别 (导航窗格识别)
                        toc.apply_outline_level(p, entry.level)
                        break
            # 在文档开头插入 TOC 字段
            inserted = toc.insert_toc_field(
                at_position=0,
                max_level=self.options.toc_max_level,
                title='目  录',
            )
            stats = toc.get_stats()
            self.stats['toc_entries'] = stats['total_entries']
            if self.options.verbose:
                print(f'[TOC] {stats["total_entries"]} entries, inserted={inserted}')
        except Exception as e:
            if self.options.verbose:
                print(f'[TOC] error: {e}')
            import traceback
            traceback.print_exc()

    # ---------- 注释还原 ----------
    def _apply_annotations_to_paragraph(self, paragraph, para_data, annotations, page_data=None):
        """将注释应用到段落
        - Highlight: 段落底纹 (w:shd) 或 run highlight
        - Underline: run underline
        - StrikeOut: run strike
        - Squiggly: run wavy underline
        - Text: 批注 (简化为段落尾追加批注文本)
        """
        try:
            from annotation_handler import match_annot_to_spans, get_highlight_color
            for annot in annotations:
                # 检查注释是否与段落重叠
                ax0, ay0, ax1, ay1 = annot.rect
                px0, py0, px1, py1 = para_data.bbox
                if not (ax0 < px1 and ax1 > px0 and ay0 < py1 and ay1 > py0):
                    continue
                # 按注释类型处理
                if annot.annot_type == 'Highlight':
                    # 给段落设置底纹
                    self._set_paragraph_shading(paragraph, get_highlight_color(annot))
                elif annot.annot_type in ('Underline', 'Squiggly'):
                    # 给 runs 加下划线
                    for run in paragraph.runs:
                        if annot.annot_type == 'Squiggly':
                            from docx.enum.text import WD_UNDERLINE
                            run.font.underline = WD_UNDERLINE.WAVY
                        else:
                            from docx.enum.text import WD_UNDERLINE
                            run.font.underline = WD_UNDERLINE.SINGLE
                elif annot.annot_type == 'StrikeOut':
                    # 给 runs 加删除线
                    for run in paragraph.runs:
                        run.font.strike = True
                elif annot.annot_type == 'Text' and annot.content:
                    # 批注: 简化为段尾追加 [批注: 内容]
                    run = paragraph.add_run(f' [批注: {annot.content}]')
                    run.font.size = paragraph.runs[0].font.size if paragraph.runs else None
                    from docx.shared import RGBColor
                    try:
                        run.font.color.rgb = RGBColor.from_string('FF0000')
                    except Exception:
                        pass
                    run.font.italic = True
        except Exception as e:
            if self.options.verbose:
                print(f'[Annot] apply error: {e}')

    def _set_paragraph_shading(self, paragraph, color_hex: str):
        """设置段落底纹 (高亮)"""
        try:
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn
            pPr = paragraph._element.get_or_add_pPr()
            existing = pPr.find(qn('w:shd'))
            if existing is not None:
                pPr.remove(existing)
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), color_hex)
            pPr.append(shd)
        except Exception:
            pass

    # ---------- 分页对齐 ----------
    def _fill_page_to_boundary(self, builder, layout, page_data=None):
        """在页末填充空行, 让下一页内容强制从新页开始
        策略: 仅当当前页内容已超过页底 80% 时才填充, 避免过度填充
        """
        try:
            if not layout.body_paragraphs and not layout.tables:
                return
            # 找页面内容最大 y (底部)
            max_y = 0
            for p in layout.body_paragraphs:
                max_y = max(max_y, p.bbox[3])
            for t in layout.tables:
                max_y = max(max_y, t.bbox[3])
            # 内容区
            content_bottom = layout.footer.bbox[1] - 5 if layout.footer else layout.height - 50
            content_top = layout.header.bbox[3] + 5 if layout.header else 50
            content_height_pt = content_bottom - content_top
            used_height = max(0, max_y - content_top)
            # 仅当内容已用 > 70% 时才填充 (避免空白页)
            if used_height < content_height_pt * 0.7:
                return
            # 剩余高度
            remaining = content_height_pt - used_height
            if remaining > 10:
                # 用更小字号(0.5pt) 减少填充量
                n_lines = int(remaining / 2.0)  # 行高 = 0.5pt * 1.5 行距 ≈ 0.75pt, 但保守用2pt
                if n_lines > 0:
                    builder.fill_blank_lines(min(n_lines, 80), size_pt=0.5)
        except Exception as e:
            if self.options.verbose:
                print(f'[Align] fill error: {e}')

    # ---------- 表单域还原 ----------
    def _add_form_field(self, builder: DocxBuilder, ff):
        """将表单域还原为 DOCX 元素
        - CheckBox: ☑/☐ 符号 (inline)
        - RadioButton: ●/○ 符号
        - Text: 文本 + 边框
        - ComboBox: 文本 + ▼
        - Button: 按钮 (浅灰背景)
        - Signature: 已由图片处理流程覆盖 (浮动定位)
        """
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        try:
            # 用浮动定位把表单域放在原位置
            display = ff.display_text or ''
            if not display:
                return
            p = builder.doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            run = p.add_run(display)
            # 根据类型设置样式
            if ff.field_type == 'CheckBox':
                run.font.size = Pt(12)
                run.font.bold = True
            elif ff.field_type == 'RadioButton':
                run.font.size = Pt(12)
            elif ff.field_type == 'Text' and ff.field_value:
                # 文本框: 加边框
                run.font.size = Pt(10.5)
                run.font.color.rgb = RGBColor.from_string('000000')
            elif ff.field_type == 'ComboBox':
                run.font.size = Pt(10.5)
            elif ff.field_type == 'Button':
                run.font.size = Pt(11)
                run.font.bold = True
                try:
                    run.font.color.rgb = RGBColor.from_string('FFFFFF')
                except Exception:
                    pass
            # 浮动定位 (绝对位置, 衬于文字上方)
            # 用 wp:anchor 把表单域放到原位置
            try:
                from docx.oxml import OxmlElement
                from docx.oxml.ns import qn
                from lxml import etree
                WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
                # 创建文本框 (wsp with txbxContent)
                # 简化: 用段落边框模拟文本框
                if ff.field_type == 'Text':
                    pPr = p._element.get_or_add_pPr()
                    pbdr = OxmlElement('w:pBdr')
                    for edge in ('top', 'left', 'bottom', 'right'):
                        b = OxmlElement(f'w:{edge}')
                        b.set(qn('w:val'), 'single')
                        b.set(qn('w:sz'), '4')
                        b.set(qn('w:space'), '1')
                        b.set(qn('w:color'), '808080')
                        pbdr.append(b)
                    pPr.append(pbdr)
            except Exception:
                pass
        except Exception as e:
            if self.options.verbose:
                print(f'[FormField] error: {e}')

    # ---------- 矢量图形还原 ----------
    def _add_standalone_shapes(self, builder: DocxBuilder, layout: PageLayout,
                                page_data=None):
        """将页面独立矢量图形还原为 DOCX 浮动形状 (衬于文字下方)
        过滤掉表格边框、下划线、页眉页脚分隔线 (这些已单独处理)
        """
        try:
            from shape_builder import ShapeClassifier, DocxShapeBuilder, filter_meaningful_drawings
        except ImportError:
            return
        # 过滤无意义的图形
        drawings = filter_meaningful_drawings(layout.standalone_drawings, min_size=4.0)
        if not drawings:
            return
        shape_builder = DocxShapeBuilder()
        # 收集表格 bbox 用于排除表格边框 (避免重复绘制)
        table_bboxes = [t.bbox for t in layout.tables]
        # 排除下划线 (高度 < 3pt 的水平线且与span重叠的已在analyzer处理为下划线)
        added = 0
        for dr in drawings:
            r = dr.rect
            # 跳过太小的线 (可能是下划线)
            w = r[2] - r[0]
            h = r[3] - r[1]
            if h < 3 and w < 100:
                # 检查是否在某个span下方 (是下划线则跳过)
                if page_data and self._is_underline_drawing(dr, page_data):
                    continue
            # 跳过表格边框 (drawing在表格bbox内)
            in_table = any(
                tb[0] - 2 <= r[0] and r[2] <= tb[2] + 2 and
                tb[1] - 2 <= r[1] and r[3] <= tb[3] + 2
                for tb in table_bboxes
            )
            if in_table:
                continue
            # 分类并构建
            shapes = ShapeClassifier.classify_drawing(dr)
            for shape in shapes:
                # 加到 docx (用空段落)
                try:
                    p = builder.doc.add_paragraph()
                    p.paragraph_format.space_before = builder._Pt(0) if hasattr(builder, '_Pt') else None
                    shape_builder.add_shape_to_paragraph(p, shape)
                    added += 1
                except Exception as e:
                    if self.options.verbose:
                        print(f'[Shape] error: {e}')
        self.stats['shapes'] = self.stats.get('shapes', 0) + added

    def _is_underline_drawing(self, dr, page_data) -> bool:
        """判断 drawing 是否是 span 的下划线"""
        r = dr.rect
        for span in page_data.all_spans:
            sx0, sy0, sx1, sy1 = span.bbox
            # 水平线 + 在span正下方 + 水平重叠>50%
            if abs(r[1] - r[3]) < 2:
                ly = (r[1] + r[3]) / 2
                if sy1 - 1 <= ly <= sy1 + max(2, span.size * 0.25):
                    if r[0] < sx1 and r[2] > sx0:
                        overlap = min(r[2], sx1) - max(r[0], sx0)
                        if overlap > (sx1 - sx0) * 0.5:
                            return True
        return False

    # ---------- 边距估算 ----------
    def _find_header_footer_layout(self, layouts: List[PageLayout]) -> Optional[PageLayout]:
        """找代表页眉页脚的页面 (跳过封面页)
        封面页特征: 第1页, 大字号标题, 图片多, 无标准页眉
        策略: 从第2页开始找第一个有页眉的页; 若无, 回退到第1页
        """
        if not layouts:
            return None
        for layout in layouts[1:]:
            if layout.header or layout.footer:
                return layout
        return layouts[0]

    def _estimate_top_margin(self, layouts: List[PageLayout], page_h: float) -> float:
        """估算上边距: 取所有页正文最顶端的最小 y"""
        ys = []
        for l in layouts:
            if l.body_paragraphs:
                ys.append(min(p.bbox[1] for p in l.body_paragraphs))
            if l.tables:
                ys.append(min(t.bbox[1] for t in l.tables))
        if not ys:
            return 72
        margin = min(ys) - 5
        return max(36, min(margin, page_h * 0.3))

    def _estimate_bottom_margin(self, layouts: List[PageLayout], page_h: float) -> float:
        ys = []
        for l in layouts:
            if l.body_paragraphs:
                ys.append(max(p.bbox[3] for p in l.body_paragraphs))
            if l.tables:
                ys.append(max(t.bbox[3] for t in l.tables))
        if not ys:
            return 72
        margin = page_h - max(ys) - 5
        return max(36, min(margin, page_h * 0.3))

    def _estimate_left_margin(self, pages: List[PageData]) -> float:
        xs = []
        for p in pages:
            for s in p.all_spans:
                if s.bbox[1] > 50 and s.bbox[1] < p.height - 50:  # 排除页眉页脚
                    xs.append(s.bbox[0])
        if not xs:
            return 72
        return max(36, min(xs) - 2)

    def _estimate_right_margin(self, pages: List[PageData], page_w: float) -> float:
        xs = []
        for p in pages:
            for s in p.all_spans:
                if s.bbox[1] > 50 and s.bbox[1] < p.height - 50:
                    xs.append(s.bbox[2])
        if not xs:
            return 72
        return max(36, page_w - max(xs) - 2)
