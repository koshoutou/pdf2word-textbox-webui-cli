#!/usr/bin/env python3
"""
api_server.py - PDF→DOCX 转换 API 服务
独立的HTTP API服务, 支持其他程序通过REST API调用转换功能

用法:
  python3 api_server.py --port 8000
  python3 api_server.py --port 8000 --host 0.0.0.0

API接口:
  POST /api/convert      - 上传PDF并转换为DOCX
  GET  /api/download/:id - 下载转换后的DOCX
  GET  /api/status       - 服务状态
  GET  /api/files        - 已转换文件列表
  DELETE /api/files      - 清空已转换文件
"""
import sys
import os
import json
import time
import uuid
import asyncio
from pathlib import Path

# 添加src到path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src'))

from converter import PDF2DocxConverter, ConversionOptions

# 使用 aiohttp 或 Flask
try:
    from aiohttp import web, MultipartReader
    AIOHTTP_AVAILABLE = True
except ImportError:
    AIOHTTP_AVAILABLE = False

if not AIOHTTP_AVAILABLE:
    # 回退到简单的 http.server
    import http.server
    import socketserver
    from urllib.parse import urlparse, parse_qs

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
UPLOAD_DIR = '/tmp/pdf2docx-uploads'
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

CONVERTER_DIR = os.path.dirname(os.path.abspath(__file__))


def convert_pdf_to_docx(pdf_path, output_path, options_dict):
    """执行转换"""
    options = ConversionOptions(
        page_range=options_dict.get('page_range'),
        page_list=options_dict.get('page_list'),
        dpi=options_dict.get('dpi', 150),
        detect_tables=options_dict.get('detect_tables', True),
        detect_lists=options_dict.get('detect_lists', True),
        render_shapes=options_dict.get('render_shapes', True),
        detect_formulas=options_dict.get('detect_formulas', True),
        detect_links=options_dict.get('detect_links', True),
        detect_form_fields=options_dict.get('detect_form_fields', True),
        preserve_page_breaks=options_dict.get('preserve_page_breaks', True),
        align_pages=options_dict.get('align_pages', True),
        fill_page_boundary=options_dict.get('fill_page_boundary', False),
        generate_toc=options_dict.get('generate_toc', True),
        toc_max_level=options_dict.get('toc_max_level', 3),
        restore_metadata=options_dict.get('restore_metadata', True),
        detect_annotations=options_dict.get('detect_annotations', True),
        password=options_dict.get('password'),
        verbose=options_dict.get('verbose', False),
    )
    converter = PDF2DocxConverter(pdf_path, options)
    stats = converter.convert(output_path)
    return stats


# ============ aiohttp 实现 ============
if AIOHTTP_AVAILABLE:
    async def handle_status(request):
        return web.json_response({
            'service': 'pdf2docx-converter',
            'version': '1.0.0',
            'status': 'running',
            'modules': 16,
            'endpoints': [
                'POST /api/convert',
                'GET /api/download/:fileId',
                'GET /api/status',
                'GET /api/files',
                'DELETE /api/files',
            ],
        })

    async def handle_convert(request):
        reader = await request.multipart()
        file_part = await reader.next()
        if not file_part or not file_part.filename:
            return web.json_response({'error': '未提供文件'}, status=400)

        # 保存上传文件
        file_id = str(uuid.uuid4())
        filename = file_part.filename
        upload_path = os.path.join(UPLOAD_DIR, f'{file_id}-{filename}')
        with open(upload_path, 'wb') as f:
            while True:
                chunk = await file_part.read_chunk()
                if not chunk:
                    break
                f.write(chunk)

        # 解析选项
        options = {}
        async for field in reader:
            if field.name:
                value = await field.text()
                options[field.name] = value

        # 解析页码
        pages = options.get('pages', '')
        page_range = None
        page_list = None
        if '-' in pages and ',' not in pages:
            parts = pages.split('-')
            if len(parts) == 2:
                page_range = (int(parts[0]), int(parts[1]))
        elif ',' in pages:
            page_list = []
            for part in pages.split(','):
                part = part.strip()
                if '-' in part:
                    a, b = part.split('-')
                    page_list.extend(range(int(a), int(b) + 1))
                else:
                    page_list.append(int(part))

        # 转换
        out_path = os.path.join(OUTPUT_DIR, f'{file_id}.docx')
        try:
            stats = convert_pdf_to_docx(upload_path, out_path, {
                'page_range': page_range,
                'page_list': page_list,
                'dpi': int(options.get('dpi', 150)),
                'detect_tables': options.get('detectTables', 'true') != 'false',
                'detect_lists': options.get('detectLists', 'true') != 'false',
                'detect_formulas': options.get('detectFormulas', 'true') != 'false',
                'detect_links': options.get('detectLinks', 'true') != 'false',
                'detect_form_fields': options.get('detectForms', 'true') != 'false',
                'generate_toc': options.get('generateToc', 'true') != 'false',
                'restore_metadata': options.get('restoreMetadata', 'true') != 'false',
                'detect_annotations': options.get('detectAnnotations', 'true') != 'false',
                'preserve_page_breaks': options.get('preservePageBreaks', 'true') != 'false',
                'align_pages': options.get('alignPages', 'true') != 'false',
                'password': options.get('password'),
                'verbose': False,
            })
            return web.json_response({
                'success': True,
                'fileId': file_id,
                'fileName': filename,
                'stats': stats,
                'downloadUrl': f'/api/download/{file_id}',
            })
        except Exception as e:
            return web.json_response({'error': str(e)}, status=500)

    async def handle_download(request):
        file_id = request.match_info.get('fileId')
        if not file_id:
            return web.json_response({'error': '缺少fileId'}, status=400)
        file_path = os.path.join(OUTPUT_DIR, f'{file_id}.docx')
        if not os.path.exists(file_path):
            return web.json_response({'error': '文件不存在'}, status=404)
        return web.FileResponse(file_path, headers={
            'Content-Disposition': f'attachment; filename="{file_id}.docx"'
        })

    async def handle_list_files(request):
        files = []
        for f in os.listdir(OUTPUT_DIR):
            if f.endswith('.docx'):
                fpath = os.path.join(OUTPUT_DIR, f)
                stat = os.stat(fpath)
                files.append({
                    'fileId': f.replace('.docx', ''),
                    'name': f,
                    'size': stat.st_size,
                    'mtime': stat.st_mtime,
                })
        return web.json_response({'files': files})

    async def handle_clear_files(request):
        count = 0
        for f in os.listdir(OUTPUT_DIR):
            if f.endswith('.docx'):
                os.remove(os.path.join(OUTPUT_DIR, f))
                count += 1
        return web.json_response({'cleared': count})

    def create_app():
        app = web.Application()
        app.router.add_GET('/api/status', handle_status)
        app.router.add_POST('/api/convert', handle_convert)
        app.router.add_GET('/api/download/{fileId}', handle_download)
        app.router.add_GET('/api/files', handle_list_files)
        app.router.add_DELETE('/api/files', handle_clear_files)
        # CORS
        @web.middleware
        async def cors_middleware(request, handler):
            response = await handler(request)
            response.headers['Access-Control-Allow-Origin'] = '*'
            response.headers['Access-Control-Allow-Methods'] = 'GET, POST, DELETE, OPTIONS'
            response.headers['Access-Control-Allow-Headers'] = '*'
            return response
        app.middlewares.append(cors_middleware)
        return app

    def main():
        import argparse
        ap = argparse.ArgumentParser(description='PDF→DOCX API服务')
        ap.add_argument('--host', default='0.0.0.0', help='监听地址')
        ap.add_argument('--port', type=int, default=8000, help='监听端口')
        args = ap.parse_args()
        print(f'PDF→DOCX API服务启动: http://{args.host}:{args.port}')
        print(f'API文档: http://{args.host}:{args.port}/api/status')
        web.run_app(create_app(), host=args.host, port=args.port)

    if __name__ == '__main__':
        main()
