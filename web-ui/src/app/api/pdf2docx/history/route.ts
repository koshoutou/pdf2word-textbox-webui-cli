/**
 * 转换历史记录 API
 * GET: 获取历史记录列表
 * POST: 添加历史记录
 * DELETE: 清空历史记录
 */
import { NextRequest, NextResponse } from 'next/server';
import { existsSync } from 'fs';
import { readFile, writeFile, mkdir } from 'fs/promises';
import path from 'path';

const HISTORY_DIR = '/tmp/pdf2docx-history';
const HISTORY_FILE = path.join(HISTORY_DIR, 'history.json');

async function ensureDir() {
  if (!existsSync(HISTORY_DIR)) await mkdir(HISTORY_DIR, { recursive: true });
}

async function readHistory(): Promise<any[]> {
  await ensureDir();
  if (!existsSync(HISTORY_FILE)) return [];
  try {
    const content = await readFile(HISTORY_FILE, 'utf8');
    return JSON.parse(content);
  } catch {
    return [];
  }
}

async function writeHistory(history: any[]) {
  await ensureDir();
  await writeFile(HISTORY_FILE, JSON.stringify(history.slice(0, 50), null, 2));
}

export async function GET() {
  const history = await readHistory();
  return NextResponse.json({ history });
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const history = await readHistory();
    history.unshift({
      fileId: body.fileId,
      fileName: body.fileName,
      timestamp: body.timestamp || new Date().toISOString(),
      stats: body.stats || {},
      fileSize: body.fileSize || 0,
      docxUrl: body.docxUrl,
      pdfUrl: body.pdfUrl,
      qualityScore: body.stats?.quality_score,
    });
    await writeHistory(history);
    return NextResponse.json({ ok: true, count: history.length });
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}

export async function DELETE() {
  try {
    await writeHistory([]);
    return NextResponse.json({ ok: true });
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
