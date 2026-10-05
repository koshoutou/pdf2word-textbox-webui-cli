/**
 * 转换进度实时反馈 API (Server-Sent Events)
 * 客户端通过 EventSource 连接, 实时接收转换进度
 * 用法: GET /api/pdf2docx/progress?jobId=xxx
 */
import { NextRequest } from 'next/server';
import { existsSync, readFile, writeFile, mkdir } from 'fs/promises';
import path from 'path';

const PROGRESS_DIR = '/tmp/pdf2docx-progress';

async function ensureDir() {
  if (!existsSync(PROGRESS_DIR)) await mkdir(PROGRESS_DIR, { recursive: true });
}

async function readProgress(jobId: string): Promise<any> {
  const filePath = path.join(PROGRESS_DIR, `${jobId}.json`);
  if (!existsSync(filePath)) return null;
  try {
    const content = await readFile(filePath, 'utf8');
    return JSON.parse(content);
  } catch {
    return null;
  }
}

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const jobId = searchParams.get('jobId');
  if (!jobId) {
    return new Response('Missing jobId', { status: 400 });
  }
  await ensureDir();

  // SSE 流
  const encoder = new TextEncoder();
  const stream = new ReadableStream({
    async start(controller) {
      let lastProgress = -1;
      let lastMessage = '';
      const startTime = Date.now();

      const sendEvent = (data: any) => {
        const msg = `data: ${JSON.stringify(data)}\n\n`;
        controller.enqueue(encoder.encode(msg));
      };

      // 初始连接消息
      sendEvent({ type: 'connected', jobId, timestamp: Date.now() });

      // 轮询进度文件 (每500ms)
      const interval = setInterval(async () => {
        try {
          const progress = await readProgress(jobId);
          if (progress) {
            const currentProgress = progress.progress || 0;
            const currentMessage = progress.message || '';
            // 仅在进度变化或消息变化时推送
            if (currentProgress !== lastProgress || currentMessage !== lastMessage) {
              lastProgress = currentProgress;
              lastMessage = currentMessage;
              sendEvent({
                type: 'progress',
                jobId,
                progress: currentProgress,
                message: currentMessage,
                stats: progress.stats || {},
                timestamp: Date.now(),
              });
            }
            // 完成
            if (progress.status === 'done' || progress.status === 'error') {
              sendEvent({
                type: progress.status,
                jobId,
                progress: currentProgress,
                message: currentMessage,
                stats: progress.stats || {},
                error: progress.error || '',
                elapsed: (Date.now() - startTime) / 1000,
                timestamp: Date.now(),
              });
              clearInterval(interval);
              controller.close();
            }
          }
        } catch (e) {
          // 忽略读取错误
        }
      }, 500);

      // 超时 (5分钟)
      setTimeout(() => {
        clearInterval(interval);
        try {
          sendEvent({ type: 'timeout', jobId, message: '转换超时' });
          controller.close();
        } catch (e) {}
      }, 300000);

      // 客户端断开时清理
      request.signal.addEventListener('abort', () => {
        clearInterval(interval);
        try { controller.close(); } catch (e) {}
      });
    },
  });

  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache, no-transform',
      'Connection': 'keep-alive',
      'Access-Control-Allow-Origin': '*',
    },
  });
}

export async function POST(request: NextRequest) {
  // 写入进度文件 (供 SSE 轮询)
  await ensureDir();
  try {
    const body = await request.json();
    const { jobId, progress, message, status, stats, error } = body;
    if (!jobId) {
      return new Response('Missing jobId', { status: 400 });
    }
    const filePath = path.join(PROGRESS_DIR, `${jobId}.json`);
    const data = {
      jobId,
      progress: progress || 0,
      message: message || '',
      status: status || 'progress',
      stats: stats || {},
      error: error || '',
      timestamp: Date.now(),
    };
    await writeFile(filePath, JSON.stringify(data, null, 2));
    return new Response(JSON.stringify({ ok: true }), {
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (e: any) {
    return new Response(JSON.stringify({ error: e.message }), { status: 500 });
  }
}
