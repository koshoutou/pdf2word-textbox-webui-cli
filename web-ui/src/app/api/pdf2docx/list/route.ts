/**
 * 获取已转换文件列表
 */
import { NextResponse } from 'next/server';
import { existsSync, readdirSync, statSync } from 'fs';
import path from 'path';

const OUTPUT_DIR = '/home/z/my-project/pdf2docx-pro/output';

export async function GET() {
  if (!existsSync(OUTPUT_DIR)) {
    return NextResponse.json({ files: [] });
  }
  const files = readdirSync(OUTPUT_DIR);
  const docxFiles = files.filter(f => f.endsWith('.docx')).map(f => {
    const fp = path.join(OUTPUT_DIR, f);
    const stat = statSync(fp);
    const fileId = f.replace(/\.docx$/, '');
    const hasPdf = existsSync(path.join(OUTPUT_DIR, `${fileId}.pdf`));
    return {
      fileId,
      name: f,
      size: stat.size,
      mtime: stat.mtime.toISOString(),
      hasPdf,
      docxUrl: `/api/pdf2docx/download?f=${fileId}&t=docx`,
      pdfUrl: hasPdf ? `/api/pdf2docx/download?f=${fileId}&t=pdf` : null,
    };
  });
  docxFiles.sort((a, b) => b.mtime.localeCompare(a.mtime));
  return NextResponse.json({ files: docxFiles });
}
