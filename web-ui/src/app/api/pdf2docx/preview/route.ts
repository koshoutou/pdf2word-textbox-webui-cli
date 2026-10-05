/**
 * PDF 页面预览渲染 API
 * 用 PyMuPDF 把指定页渲染成 PNG 返回
 */
import { NextRequest, NextResponse } from 'next/server';
import { existsSync } from 'fs';
import { writeFile, readFile, mkdir } from 'fs/promises';
import path from 'path';
import { spawn } from 'child_process';

const OUTPUT_DIR = '/home/z/my-project/pdf2docx-pro/output';
const CACHE_DIR = '/tmp/pdf2docx-preview';

async function ensureDirs() {
  if (!existsSync(CACHE_DIR)) await mkdir(CACHE_DIR, { recursive: true });
}

export async function GET(request: NextRequest) {
  await ensureDirs();
  const { searchParams } = new URL(request.url);
  const fileId = searchParams.get('f');
  const pageStr = searchParams.get('p') || '1';
  const source = searchParams.get('s') || 'pdf'; // pdf / docx
  const page = parseInt(pageStr, 10);
  if (!fileId) {
    return NextResponse.json({ error: '缺少 f 参数' }, { status: 400 });
  }
  // 找源文件
  const pdfPath = path.join(OUTPUT_DIR, `${fileId}.pdf`);
  const docxPath = path.join(OUTPUT_DIR, `${fileId}.docx`);
  let srcPath = source === 'docx' ? docxPath : pdfPath;
  // docx 需要先转 pdf (缓存)
  if (source === 'docx') {
    const docxPdfPath = path.join(CACHE_DIR, `${fileId}_docx.pdf`);
    if (!existsSync(docxPdfPath)) {
      if (!existsSync(docxPath)) {
        return NextResponse.json({ error: 'DOCX 不存在' }, { status: 404 });
      }
      // 用 soffice 转 (输出到 CACHE_DIR, 文件名会是 原名.pdf)
      await new Promise<void>((resolve) => {
        const proc = spawn('soffice', ['--headless', '--convert-to', 'pdf',
                                       '--outdir', CACHE_DIR, docxPath]);
        proc.on('close', () => resolve());
        proc.on('error', () => resolve());
      });
      // soffice 输出文件名 = 输入文件 basename + .pdf
      // 输入是 {fileId}.docx → 输出是 {fileId}.pdf (在 CACHE_DIR)
      const sofficeOut = path.join(CACHE_DIR, `${fileId}.pdf`);
      if (existsSync(sofficeOut)) {
        try {
          const data = await readFile(sofficeOut);
          await writeFile(docxPdfPath, data);
        } catch (e) {
          // 重命名失败
        }
      }
    }
    srcPath = docxPdfPath;
  }
  if (!existsSync(srcPath)) {
    return NextResponse.json({ error: '源文件不存在', path: srcPath }, { status: 404 });
  }
  // 渲染指定页
  const cacheKey = `${fileId}_${source}_p${page}.png`;
  const cachePath = path.join(CACHE_DIR, cacheKey);
  if (!existsSync(cachePath)) {
    const script = `
import fitz, sys
doc = fitz.open("${srcPath}")
pno = ${page - 1}
if 0 <= pno < len(doc):
    mat = fitz.Matrix(1.5, 1.5)
    pix = doc[pno].get_pixmap(matrix=mat)
    pix.save("${cachePath}")
    print(len(doc))
else:
    print("ERR: out of range")
doc.close()
`;
    await new Promise<void>((resolve) => {
      const proc = spawn('python3', ['-c', script]);
      proc.on('close', () => resolve());
      proc.on('error', () => resolve());
    });
  }
  if (!existsSync(cachePath)) {
    return NextResponse.json({ error: '渲染失败' }, { status: 500 });
  }
  const buffer = await readFile(cachePath);
  return new NextResponse(buffer, {
    headers: {
      'Content-Type': 'image/png',
      'Cache-Control': 'public, max-age=86400',
    },
  });
}
