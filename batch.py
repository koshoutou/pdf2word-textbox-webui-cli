#!/usr/bin/env python3
"""
batch.py - 批量处理模式
扫描目录下所有 PDF, 逐个转换为 DOCX
用法:
  python3 batch.py -i /input/dir -o /output/dir
  python3 batch.py -i /input/dir -o /output/dir --recursive
  python3 batch.py -i /input/dir -o /output/dir --pages 1-5 --dpi 200
"""
import sys
import os
import argparse
import time
import glob

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src'))
from converter import PDF2DocxConverter, ConversionOptions


def find_pdfs(input_dir: str, recursive: bool = False) -> list:
    """查找目录下所有 PDF"""
    pattern = '**/*.pdf' if recursive else '*.pdf'
    # 大小写不敏感
    pdfs = []
    for ext in ['*.pdf', '*.PDF']:
        pdfs.extend(glob.glob(os.path.join(input_dir, ext)))
        if recursive:
            pdfs.extend(glob.glob(os.path.join(input_dir, '**', ext), recursive=True))
    # 去重 + 排序
    return sorted(set(pdfs))


def parse_pages(s: str):
    if not s:
        return None, None
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
    return None, [int(s)]


def main():
    ap = argparse.ArgumentParser(
        description='PDF→DOCX 批量转换器',
        formatter_class=argparse.RawTextHelpFormatter,
    )
    ap.add_argument('-i', '--input', required=True, help='输入目录')
    ap.add_argument('-o', '--output', required=True, help='输出目录')
    ap.add_argument('-r', '--recursive', action='store_true', help='递归扫描子目录')
    ap.add_argument('-p', '--pages', default=None, help='页码范围 (应用到所有PDF)')
    ap.add_argument('-d', '--dpi', type=int, default=150, help='图片 DPI')
    ap.add_argument('--no-tables', action='store_true', help='禁用表格检测')
    ap.add_argument('--no-lists', action='store_true', help='禁用列表检测')
    ap.add_argument('--no-shapes', action='store_true', help='禁用矢量图形还原')
    ap.add_argument('--no-page-breaks', action='store_true', help='不保留分页符')
    ap.add_argument('-w', '--workers', type=int, default=1, help='并行工作进程数(默认1=串行)')
    ap.add_argument('--report', action='store_true', help='每个PDF生成质量报告')
    ap.add_argument('-v', '--verbose', action='store_true', help='详细日志')
    args = ap.parse_args()

    if not os.path.isdir(args.input):
        print(f'错误: 输入目录不存在: {args.input}', file=sys.stderr)
        sys.exit(1)
    os.makedirs(args.output, exist_ok=True)

    pdfs = find_pdfs(args.input, args.recursive)
    if not pdfs:
        print(f'未找到 PDF 文件: {args.input}')
        sys.exit(0)

    print(f'找到 {len(pdfs)} 个 PDF 文件')
    print('=' * 60)

    page_range, page_list = parse_pages(args.pages) if args.pages else (None, None)
    options = ConversionOptions(
        page_range=page_range,
        page_list=page_list,
        dpi=args.dpi,
        detect_tables=not args.no_tables,
        detect_lists=not args.no_lists,
        render_shapes=not args.no_shapes,
        preserve_page_breaks=not args.no_page_breaks,
        verbose=args.verbose,
    )

    success = 0
    failed = 0
    total_time = 0.0
    results = []

    # 准备任务列表
    tasks = []
    for i, pdf_path in enumerate(pdfs, 1):
        rel_path = os.path.relpath(pdf_path, args.input)
        out_name = os.path.splitext(rel_path)[0].replace('/', '_') + '.docx'
        out_path = os.path.join(args.output, out_name)
        if args.recursive:
            out_subdir = os.path.dirname(os.path.join(args.output, rel_path))
            os.makedirs(out_subdir, exist_ok=True)
            out_path = os.path.join(out_subdir, os.path.splitext(os.path.basename(rel_path))[0] + '.docx')
        tasks.append((i, pdf_path, rel_path, out_path))

    # ★ 并行/串行处理
    if args.workers > 1:
        results = _process_parallel(tasks, options, args)
    else:
        results = _process_serial(tasks, options, args)

    success = sum(1 for r in results if r['status'] == 'ok')
    failed = sum(1 for r in results if r['status'] == 'fail')
    total_time = sum(r.get('time', 0) for r in results)

    print('=' * 60)
    print(f'批量转换完成')
    print(f'  总数: {len(pdfs)}')
    print(f'  成功: {success}')
    print(f'  失败: {failed}')
    print(f'  总用时: {total_time:.1f}s  平均: {total_time/len(pdfs):.1f}s/个')

    # 输出 CSV 报告
    csv_path = os.path.join(args.output, '_batch_report.csv')
    with open(csv_path, 'w', encoding='utf-8') as f:
        f.write('文件,状态,用时(秒),段落,表格,图片,列表项,形状,错误\n')
        for r in results:
            stats = r.get('stats', {})
            f.write(f'"{r["pdf"]}",{r["status"]},{r["time"]:.1f},'
                    f'{stats.get("paragraphs","")},{stats.get("tables","")},'
                    f'{stats.get("images","")},{stats.get("list_items","")},'
                    f'{stats.get("shapes","")},{r.get("error","")}\n')
    print(f'  报告: {csv_path}')


# ---------- 串行处理 ----------
def _process_serial(tasks, options, args):
    """串行处理 PDF 列表"""
    results = []
    for i, pdf_path, rel_path, out_path in tasks:
        print(f'[{i}/{len(tasks)}] {rel_path}')
        t0 = time.time()
        try:
            converter = PDF2DocxConverter(pdf_path, options)
            stats = converter.convert(out_path)
            elapsed = time.time() - t0
            score = ''
            if args.report:
                try:
                    from reporter import ConversionReporter
                    page_range = options.page_range
                    page_list = options.page_list
                    reporter = ConversionReporter(pdf_path, out_path,
                                                   page_range=page_range, page_list=page_list)
                    report = reporter.generate()
                    score = f' 评分:{report["quality_score"]}/100'
                except Exception:
                    pass
            print(f'  ✓ 用时 {elapsed:.1f}s  段落:{stats.get("paragraphs",0)} 表格:{stats.get("tables",0)} 图片:{stats.get("images",0)} 列表:{stats.get("list_items",0)}{score}')
            results.append({'pdf': rel_path, 'status': 'ok', 'time': elapsed, 'stats': stats})
        except Exception as e:
            elapsed = time.time() - t0
            print(f'  ✗ 失败: {e}')
            results.append({'pdf': rel_path, 'status': 'fail', 'error': str(e), 'time': elapsed})
    return results


# ---------- 并行处理 ----------
def _process_worker(task):
    """并行处理单个 PDF (worker 函数, 必须在顶层以便 pickle)"""
    i, pdf_path, rel_path, out_path, options = task
    t0 = time.time()
    try:
        # 每个 worker 关闭 verbose 输出, 避免日志混乱
        opts = ConversionOptions(
            page_range=options.page_range,
            page_list=options.page_list,
            dpi=options.dpi,
            detect_tables=options.detect_tables,
            detect_lists=options.detect_lists,
            render_shapes=options.render_shapes,
            detect_formulas=options.detect_formulas,
            detect_links=options.detect_links,
            detect_form_fields=options.detect_form_fields,
            preserve_page_breaks=options.preserve_page_breaks,
            align_pages=options.align_pages,
            fill_page_boundary=options.fill_page_boundary,
            generate_toc=options.generate_toc,
            toc_max_level=options.toc_max_level,
            restore_metadata=options.restore_metadata,
            verbose=False,  # 关闭日志
        )
        converter = PDF2DocxConverter(pdf_path, opts)
        stats = converter.convert(out_path)
        elapsed = time.time() - t0
        return {'pdf': rel_path, 'status': 'ok', 'time': elapsed, 'stats': stats}
    except Exception as e:
        elapsed = time.time() - t0
        return {'pdf': rel_path, 'status': 'fail', 'error': str(e), 'time': elapsed}


def _process_parallel(tasks, options, args):
    """并行处理 PDF 列表 (multiprocessing)"""
    from multiprocessing import Pool
    # 准备 worker 任务
    worker_tasks = [(i, p, r, o, options) for i, p, r, o in tasks]
    print(f'启动 {args.workers} 个并行工作进程...')
    results = []
    with Pool(args.workers) as pool:
        for i, result in enumerate(pool.imap(_process_worker, worker_tasks), 1):
            elapsed = result.get('time', 0)
            status = result.get('status')
            if status == 'ok':
                stats = result.get('stats', {})
                print(f'[{i}/{len(tasks)}] ✓ {result["pdf"]} ({elapsed:.1f}s) 段落:{stats.get("paragraphs",0)} 表格:{stats.get("tables",0)}')
            else:
                print(f'[{i}/{len(tasks)}] ✗ {result["pdf"]} 失败: {result.get("error", "")}')
            results.append(result)
    return results


if __name__ == '__main__':
    main()
