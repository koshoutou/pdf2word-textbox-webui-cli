/**
 * PDF→DOCX 转换 API
 * 接收 PDF 文件 + 选项, 调用 Python 转换器, 返回 DOCX 文件信息
 */
import { NextRequest, NextResponse } from 'next/server';
import { writeFile, readFile, mkdir } from 'fs/promises';
import { existsSync } from 'fs';
import path from 'path';
import { spawn } from 'child_process';
import { randomUUID } from 'crypto';

const UPLOAD_DIR = '/tmp/pdf2docx-uploads';
const OUTPUT_DIR = '/home/z/my-project/pdf2docx-pro/output';
const CONVERTER_DIR = '/home/z/my-project/pdf2docx-pro';

async function ensureDirs() {
  if (!existsSync(UPLOAD_DIR)) await mkdir(UPLOAD_DIR, { recursive: true });
  if (!existsSync(OUTPUT_DIR)) await mkdir(OUTPUT_DIR, { recursive: true });
}

export async function POST(request: NextRequest) {
  try {
    await ensureDirs();
    const formData = await request.formData();
    const file = formData.get('file') as File | null;
    if (!file) {
      return NextResponse.json({ error: '未提供文件' }, { status: 400 });
    }
    if (!file.name.toLowerCase().endsWith('.pdf')) {
      return NextResponse.json({ error: '仅支持 PDF 文件' }, { status: 400 });
    }

    // 选项
    const pages = (formData.get('pages') as string) || '';
    const dpi = parseInt((formData.get('dpi') as string) || '150', 10);
    const detectTables = formData.get('detectTables') !== 'false';
    const detectLists = formData.get('detectLists') !== 'false';
    const renderShapes = formData.get('renderShapes') !== 'false';
    const detectFormulas = formData.get('detectFormulas') !== 'false';
    const detectLinks = formData.get('detectLinks') !== 'false';
    const detectForms = formData.get('detectForms') !== 'false';
    const alignPages = formData.get('alignPages') !== 'false';
    const fillPages = formData.get('fillPages') === 'true';
    const generateToc = formData.get('generateToc') !== 'false';
    const tocLevel = parseInt((formData.get('tocLevel') as string) || '3', 10);
    const restoreMetadata = formData.get('restoreMetadata') !== 'false';
    const detectAnnotations = formData.get('detectAnnotations') !== 'false';
    const preservePageBreaks = formData.get('preservePageBreaks') !== 'false';
    const withReport = formData.get('report') === 'true';

    // 保存上传文件
    const fileId = randomUUID();
    const jobId = fileId; // 用 fileId 作为 jobId 供进度跟踪
    const uploadPath = path.join(UPLOAD_DIR, `${fileId}-${file.name}`);
    const bytes = await file.arrayBuffer();
    await writeFile(uploadPath, Buffer.from(bytes));

    // 初始化进度文件
    try {
      const PROGRESS_DIR = '/tmp/pdf2docx-progress';
      if (!existsSync(PROGRESS_DIR)) await mkdir(PROGRESS_DIR, { recursive: true });
      await writeFile(path.join(PROGRESS_DIR, `${jobId}.json`), JSON.stringify({
        jobId, progress: 5, message: '上传完成, 开始转换', status: 'progress', stats: {}, timestamp: Date.now(),
      }));
    } catch (e) {}

    // 输出文件
    const outName = `${fileId}.docx`;
    const outPath = path.join(OUTPUT_DIR, outName);

    // 构建 Python 命令
    const args = [
      'main.py',
      '-i', uploadPath,
      '-o', outPath,
      '-d', String(dpi),
    ];
    if (pages) {
      args.push('-p', pages);
    }
    if (!detectTables) args.push('--no-tables');
    if (!detectLists) args.push('--no-lists');
    if (!renderShapes) args.push('--no-shapes');
    if (!detectFormulas) args.push('--no-formulas');
    if (!detectLinks) args.push('--no-links');
    if (!detectForms) args.push('--no-forms');
    if (!alignPages) args.push('--no-align');
    if (fillPages) args.push('--fill-pages');
    if (!generateToc) args.push('--no-toc');
    args.push('--toc-level', String(tocLevel));
    if (!restoreMetadata) args.push('--no-metadata');
    if (!detectAnnotations) args.push('--no-annotations');
    if (!preservePageBreaks) args.push('--no-page-breaks');
    if (withReport) args.push('--report');
    args.push('-v');

    // 调用 Python (限制内存+超时, 避免OOM拖垮dev server)
    const result = await new Promise<{ ok: boolean; stdout: string; stderr: string; timeout: boolean }>((resolve) => {
      const proc = spawn('python3', args, {
        cwd: CONVERTER_DIR,
        env: { ...process.env, PYTHONUNBUFFERED: '1' },
        // 限制子进程资源
        stdio: ['pipe', 'pipe', 'pipe'],
      });
      let stdout = '';
      let stderr = '';
      let resolved = false;
      const timeoutMs = 120000; // 2分钟超时
      const timer = setTimeout(() => {
        if (!resolved) {
          resolved = true;
          try { proc.kill('SIGKILL'); } catch (e) {}
          resolve({ ok: false, stdout, stderr: '转换超时(' + timeoutMs/1000 + '秒)', timeout: true });
        }
      }, timeoutMs);
      proc.stdout.on('data', (d) => {
        const chunk = d.toString();
        // 限制stdout长度避免内存累积 (保留最后100KB)
        if (stdout.length < 200000) stdout += chunk.slice(0, 50000);
        // ★ 解析进度并写入进度文件
        try {
          const progressMatch = chunk.match(/\[Progress\]\s*(\d+)%\s*(.*)/);
          if (progressMatch) {
            const progress = parseInt(progressMatch[1], 10);
            const message = progressMatch[2].trim();
            const PROGRESS_DIR = '/tmp/pdf2docx-progress';
            writeFile(path.join(PROGRESS_DIR, `${jobId}.json`), JSON.stringify({
              jobId, progress, message, status: 'progress', stats: {}, timestamp: Date.now(),
            }));
          }
        } catch (e) {}
      });
      proc.stderr.on('data', (d) => {
        const chunk = d.toString();
        if (stderr.length < 100000) stderr += chunk.slice(0, 30000);
      });
      proc.on('close', (code) => {
        if (!resolved) {
          resolved = true;
          clearTimeout(timer);
          resolve({ ok: code === 0, stdout, stderr, timeout: false });
        }
      });
      proc.on('error', (err) => {
        if (!resolved) {
          resolved = true;
          clearTimeout(timer);
          resolve({ ok: false, stdout, stderr: err.message, timeout: false });
        }
      });
    });

    if (!result.ok || !existsSync(outPath)) {
      // 写入错误进度
      try {
        const PROGRESS_DIR = '/tmp/pdf2docx-progress';
        await writeFile(path.join(PROGRESS_DIR, `${jobId}.json`), JSON.stringify({
          jobId, progress: 0, message: '转换失败', status: 'error',
          error: result.stderr.slice(-500), stats: {}, timestamp: Date.now(),
        }));
      } catch (e) {}
      return NextResponse.json({
        error: '转换失败',
        stderr: result.stderr.slice(-2000),
        stdout: result.stdout.slice(-2000),
      }, { status: 500 });
    }

    // 解析统计
    const stats = parseStats(result.stdout + result.stderr);

    // 同时把 PDF 复制到 output 方便预览
    const pdfCopyPath = path.join(OUTPUT_DIR, `${fileId}.pdf`);
    await writeFile(pdfCopyPath, Buffer.from(bytes));

    // 写入完成进度
    try {
      const PROGRESS_DIR = '/tmp/pdf2docx-progress';
      await writeFile(path.join(PROGRESS_DIR, `${jobId}.json`), JSON.stringify({
        jobId, progress: 100, message: '转换完成', status: 'done',
        stats, timestamp: Date.now(),
      }));
    } catch (e) {}

    return NextResponse.json({
      success: true,
      fileId,
      jobId,
      fileName: file.name,
      docxPath: outPath,
      pdfPath: pdfCopyPath,
      docxUrl: `/api/pdf2docx/download?f=${fileId}&t=docx`,
      pdfUrl: `/api/pdf2docx/download?f=${fileId}&t=pdf`,
      previewUrl: `/api/pdf2docx/preview?f=${fileId}`,
      stats,
      log: result.stdout.slice(-3000),
    });

    // ★ 写入历史记录
    try {
      const HISTORY_DIR = '/tmp/pdf2docx-history';
      if (!existsSync(HISTORY_DIR)) await mkdir(HISTORY_DIR, { recursive: true });
      const historyFile = path.join(HISTORY_DIR, 'history.json');
      let history: any[] = [];
      if (existsSync(historyFile)) {
        try {
          history = JSON.parse(await readFile(historyFile, 'utf8'));
        } catch {}
      }
      history.unshift({
        fileId,
        fileName: file.name,
        timestamp: new Date().toISOString(),
        stats,
        fileSize: file.size,
        docxUrl: `/api/pdf2docx/download?f=${fileId}&t=docx`,
        pdfUrl: `/api/pdf2docx/download?f=${fileId}&t=pdf`,
        qualityScore: stats.quality_score,
      });
      await writeFile(historyFile, JSON.stringify(history.slice(0, 50), null, 2));
    } catch (e) {}
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}

function parseStats(log: string): any {
  const stats: any = {};
  const patterns: [string, RegExp][] = [
    ['pages', /页数:\s*(\d+)/],
    ['paragraphs', /段落:\s*(\d+)/],
    ['tables', /表格:\s*(\d+)/],
    ['images', /图片:\s*(\d+)/],
    ['list_items', /list_items['"]?\s*[:=]\s*(\d+)/],
    ['shapes', /shapes['"]?\s*[:=]\s*(\d+)/],
    ['formulas', /formulas['"]?\s*[:=]\s*(\d+)/],
    ['links', /links['"]?\s*[:=]\s*(\d+)/],
    ['form_fields', /form_fields['"]?\s*[:=]\s*(\d+)/],
    ['toc_entries', /toc_entries['"]?\s*[:=]\s*(\d+)/],
    ['metadata_restored', /metadata_restored['"]?\s*[:=]\s*(\d+)/],
    ['annotations', /annotations['"]?\s*[:=]\s*(\d+)/],
  ];
  for (const [k, p] of patterns) {
    const m = log.match(p);
    if (m) stats[k] = parseInt(m[1], 10);
  }
  // 解析质量评分
  const scoreMatch = log.match(/综合质量评分:\s*([\d.]+)\/100/);
  if (scoreMatch) stats.quality_score = parseFloat(scoreMatch[1]);
  return stats;
}

export async function GET() {
  return NextResponse.json({
    name: 'PDF→DOCX 转换 API',
    methods: ['POST'],
    usage: 'POST multipart/form-data with file=, pages=, dpi=, report=',
  });
}
