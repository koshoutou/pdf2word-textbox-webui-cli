#!/usr/bin/env python3
"""
main.py - PDF→DOCX 企业级转换器 CLI 入口
用法:
  python3 main.py -i input.pdf -o output.docx
  python3 main.py -i input.pdf -o output.docx --pages 1-10
  python3 main.py -i input.pdf -o output.docx --pages 1,3,5-8
  python3 main.py -i input.pdf -o output.docx --verbose
"""
import sys
import os
import argparse

# 把 src 加入 path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src'))

from converter import PDF2DocxConverter, ConversionOptions


def parse_pages(s: str):
    """解析 '1-10' 或 '1,3,5-8' → page_list"""
    if not s:
        return None, None
    if '-' in s and ',' not in s:
        parts = s.split('-')
        if len(parts) == 2:
            return (int(parts[0]), int(parts[1])), None
    if ',' in s:
        nums = []
        for part in s.split(','):
            part = part.strip()
            if '-' in part:
                a, b = part.split('-')
                nums.extend(range(int(a), int(b) + 1))
            else:
                nums.append(int(part))
        return None, nums
    if '-' in s:
        parts = s.split('-')
        return (int(parts[0]), int(parts[1])), None
    # 单页
    return None, [int(s)]


def main():
    ap = argparse.ArgumentParser(
        description='PDF→DOCX 企业级 1:1 复刻转换器',
        formatter_class=argparse.RawTextHelpFormatter,
    )
    ap.add_argument('-i', '--input', required=True, help='输入 PDF 文件路径')
    ap.add_argument('-o', '--output', required=True, help='输出 DOCX 文件路径')
    ap.add_argument('-p', '--pages', default=None,
                    help='页码范围: 1-10 或 1,3,5-8 (1-based), 不指定则全转换')
    ap.add_argument('-d', '--dpi', type=int, default=150, help='图片渲染 DPI (默认 150)')
    ap.add_argument('--no-tables', action='store_true', help='禁用表格检测')
    ap.add_argument('--no-lists', action='store_true', help='禁用列表/编号检测')
    ap.add_argument('--no-shapes', action='store_true', help='禁用矢量图形还原')
    ap.add_argument('--no-formulas', action='store_true', help='禁用数学公式检测')
    ap.add_argument('--no-links', action='store_true', help='禁用超链接/书签还原')
    ap.add_argument('--no-forms', action='store_true', help='禁用表单域识别')
    ap.add_argument('--no-align', action='store_true', help='禁用分页对齐(keepNext/cantSplit)')
    ap.add_argument('--fill-pages', action='store_true', help='启用页面填充对齐(重量级,可能增加页数)')
    ap.add_argument('--no-toc', action='store_true', help='禁用自动生成目录')
    ap.add_argument('--toc-level', type=int, default=3, help='目录最大级别(1-9, 默认3)')
    ap.add_argument('--no-metadata', action='store_true', help='禁用PDF元数据还原')
    ap.add_argument('--no-annotations', action='store_true', help='禁用PDF注释还原')
    ap.add_argument('--no-page-breaks', action='store_true', help='不保留分页符')
    ap.add_argument('--report', action='store_true', help='生成转换质量报告')
    ap.add_argument('-v', '--verbose', action='store_true', help='详细日志')
    args = ap.parse_args()

    if not os.path.exists(args.input):
        print(f'错误: 输入文件不存在: {args.input}', file=sys.stderr)
        sys.exit(1)

    page_range, page_list = parse_pages(args.pages) if args.pages else (None, None)

    options = ConversionOptions(
        page_range=page_range,
        page_list=page_list,
        dpi=args.dpi,
        detect_tables=not args.no_tables,
        detect_lists=not args.no_lists,
        render_shapes=not args.no_shapes,
        detect_formulas=not args.no_formulas,
        detect_links=not args.no_links,
        detect_form_fields=not args.no_forms,
        align_pages=not args.no_align,
        fill_page_boundary=args.fill_pages,
        generate_toc=not args.no_toc,
        toc_max_level=max(1, min(9, args.toc_level)),
        restore_metadata=not args.no_metadata,
        detect_annotations=not args.no_annotations,
        preserve_page_breaks=not args.no_page_breaks,
        verbose=args.verbose,
    )

    converter = PDF2DocxConverter(args.input, options)
    try:
        stats = converter.convert(args.output)
        print(f'\n✓ 转换成功')
        print(f'  输入: {args.input}')
        print(f'  输出: {args.output}')
        print(f'  页数: {stats["pages"]}')
        print(f'  段落: {stats["paragraphs"]}')
        print(f'  表格: {stats["tables"]}')
        print(f'  图片: {stats["images"]}')
        if stats['errors']:
            print(f'  警告: {len(stats["errors"])} 个错误')
            for e in stats['errors'][:5]:
                print(f'    - {e}')
        # 质量报告
        if args.report:
            print()
            try:
                sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src'))
                from reporter import ConversionReporter
                reporter = ConversionReporter(args.input, args.output,
                                               verbose=args.verbose,
                                               page_range=page_range,
                                               page_list=page_list)
                report = reporter.generate()
                reporter.print_report(report)
            except Exception as e:
                print(f'报告生成失败: {e}')
    except Exception as e:
        import traceback
        print(f'✗ 转换失败: {e}', file=sys.stderr)
        if args.verbose:
            traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
