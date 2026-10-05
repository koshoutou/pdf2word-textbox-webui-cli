/**
 * 文件下载 API (DOCX / PDF)
 */
import { NextRequest, NextResponse } from 'next/server';
import { existsSync } from 'fs';
import { readFile } from 'fs/promises';
import path from 'path';

const OUTPUT_DIR = '/home/z/my-project/pdf2docx-pro/output';

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const fileId = searchParams.get('f');
  const type = searchParams.get('t') || 'docx';  // docx / pdf
  if (!fileId) {
    return NextResponse.json({ error: '缺少 f 参数' }, { status: 400 });
  }
  const ext = type === 'pdf' ? 'pdf' : 'docx';
  const filePath = path.join(OUTPUT_DIR, `${fileId}.${ext}`);
  if (!existsSync(filePath)) {
    return NextResponse.json({ error: '文件不存在', path: filePath }, { status: 404 });
  }
  const buffer = await readFile(filePath);
  const contentType = ext === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  return new NextResponse(buffer, {
    headers: {
      'Content-Type': contentType,
      'Content-Disposition': `attachment; filename="${fileId}.${ext}"`,
      'Content-Length': String(buffer.length),
    },
  });
}
