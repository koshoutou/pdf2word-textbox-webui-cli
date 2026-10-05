'use client';

import { useState, useRef, useCallback, useEffect } from 'react';
import { Upload, FileText, Loader2, Download, Eye, Settings, ArrowRight, CheckCircle2, AlertCircle, Layers, Image as ImageIcon, Table as TableIcon, ListOrdered, Type, FunctionSquare, Link2, ClipboardCheck, AlignJustify, ListTree, FileHeart, Highlighter, History, Trash2, Clock } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Switch } from '@/components/ui/switch';
import { Badge } from '@/components/ui/badge';
import { Separator } from '@/components/ui/separator';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Slider } from '@/components/ui/slider';
import { Progress } from '@/components/ui/progress';
import { Toaster } from '@/components/ui/toaster';
import { useToast } from '@/hooks/use-toast';
import { ScrollArea } from '@/components/ui/scroll-area';

interface ConvertStats {
  pages?: number;
  paragraphs?: number;
  tables?: number;
  images?: number;
  list_items?: number;
  shapes?: number;
  quality_score?: number;
}

interface ConvertResult {
  success: boolean;
  fileId: string;
  fileName: string;
  docxUrl: string;
  pdfUrl: string;
  previewUrl: string;
  stats: ConvertStats;
  log: string;
  error?: string;
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [converting, setConverting] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressMsg, setProgressMsg] = useState('');
  const [result, setResult] = useState<ConvertResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const { toast } = useToast();

  // 选项
  const [pages, setPages] = useState('');
  const [dpi, setDpi] = useState(150);
  const [detectTables, setDetectTables] = useState(true);
  const [detectLists, setDetectLists] = useState(true);
  const [renderShapes, setRenderShapes] = useState(true);
  const [preservePageBreaks, setPreservePageBreaks] = useState(true);
  const [withReport, setWithReport] = useState(true);
  const [detectFormulas, setDetectFormulas] = useState(true);
  const [detectLinks, setDetectLinks] = useState(true);
  const [detectForms, setDetectForms] = useState(true);
  const [alignPages, setAlignPages] = useState(true);
  const [fillPages, setFillPages] = useState(false);
  const [generateToc, setGenerateToc] = useState(true);
  const [tocLevel, setTocLevel] = useState(3);
  const [restoreMetadata, setRestoreMetadata] = useState(true);
  const [detectAnnotations, setDetectAnnotations] = useState(true);

  // 预览
  const [previewPage, setPreviewPage] = useState(1);
  const [previewSource, setPreviewSource] = useState<'pdf' | 'docx'>('pdf');

  const handleFile = useCallback((f: File) => {
    if (!f.name.toLowerCase().endsWith('.pdf')) {
      toast({ title: '格式错误', description: '仅支持 PDF 文件', variant: 'destructive' });
      return;
    }
    setFile(f);
    setResult(null);
    setError(null);
  }, [toast]);

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const f = e.dataTransfer.files?.[0];
    if (f) handleFile(f);
  }, [handleFile]);

  const handleConvert = async () => {
    if (!file) {
      toast({ title: '请先选择 PDF 文件', variant: 'destructive' });
      return;
    }
    setConverting(true);
    setProgress(10);
    setProgressMsg('上传中...');
    setError(null);
    setResult(null);
    try {
      const formData = new FormData();
      formData.append('file', file);
      if (pages) formData.append('pages', pages);
      formData.append('dpi', String(dpi));
      formData.append('detectTables', String(detectTables));
      formData.append('detectLists', String(detectLists));
      formData.append('renderShapes', String(renderShapes));
      formData.append('detectFormulas', String(detectFormulas));
      formData.append('detectLinks', String(detectLinks));
      formData.append('detectForms', String(detectForms));
      formData.append('alignPages', String(alignPages));
      formData.append('fillPages', String(fillPages));
      formData.append('generateToc', String(generateToc));
      formData.append('tocLevel', String(tocLevel));
      formData.append('restoreMetadata', String(restoreMetadata));
      formData.append('detectAnnotations', String(detectAnnotations));
      formData.append('preservePageBreaks', String(preservePageBreaks));
      formData.append('report', String(withReport));

      setProgress(10);
      setProgressMsg('上传中...');
      // 异步发起转换
      const res = await fetch('/api/pdf2docx/convert', { method: 'POST', body: formData });
      setProgress(80);
      setProgressMsg('处理中...');
      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || '转换失败');
      }
      setProgress(100);
      setProgressMsg('转换完成');
      setResult(data);
      toast({
        title: '转换成功',
        description: `已生成 DOCX${data.stats?.quality_score ? ` · 评分 ${data.stats.quality_score}/100` : ''}`,
      });
    } catch (e: any) {
      setError(e.message);
      toast({ title: '转换失败', description: e.message, variant: 'destructive' });
    } finally {
      setConverting(false);
      setTimeout(() => { setProgress(0); setProgressMsg(''); }, 1500);
    }
  };

  return (
    <div className="min-h-screen flex flex-col bg-gradient-to-br from-slate-50 via-white to-slate-100 dark:from-slate-950 dark:via-slate-900 dark:to-slate-950">
      {/* Header */}
      <header className="border-b bg-white/80 dark:bg-slate-950/80 backdrop-blur-sm sticky top-0 z-50">
        <div className="container mx-auto px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-rose-500 to-orange-500 flex items-center justify-center shadow-lg">
              <FileText className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight">PDF → DOCX 企业级转换器</h1>
              <p className="text-xs text-muted-foreground">1:1 精准复刻 · 自研内核 · 不依赖 pdf2docx</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <Badge variant="outline" className="hidden sm:flex">v12.0</Badge>
            <Badge className="bg-emerald-500 hover:bg-emerald-600">企业级</Badge>
          </div>
        </div>
      </header>

      <main className="flex-1 container mx-auto px-4 py-6">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          {/* 左侧: 上传 + 选项 */}
          <div className="lg:col-span-5 space-y-4">
            {/* 上传区 */}
            <Card className="border-2 border-dashed transition-all" >
              <CardContent className="p-6">
                <div
                  onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                  onDragLeave={() => setDragOver(false)}
                  onDrop={onDrop}
                  onClick={() => inputRef.current?.click()}
                  className={`cursor-pointer rounded-xl border-2 border-dashed transition-all p-8 text-center
                    ${dragOver ? 'border-rose-500 bg-rose-50 dark:bg-rose-950/20' :
                      file ? 'border-emerald-500 bg-emerald-50 dark:bg-emerald-950/20' :
                      'border-slate-300 dark:border-slate-700 hover:border-rose-400 hover:bg-slate-50 dark:hover:bg-slate-900'}`}
                >
                  <input
                    ref={inputRef}
                    type="file"
                    accept=".pdf"
                    className="hidden"
                    onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
                  />
                  {file ? (
                    <div className="space-y-2">
                      <CheckCircle2 className="w-12 h-12 mx-auto text-emerald-500" />
                      <p className="font-medium text-sm">{file.name}</p>
                      <p className="text-xs text-muted-foreground">{(file.size / 1024 / 1024).toFixed(2)} MB</p>
                    </div>
                  ) : (
                    <div className="space-y-2">
                      <Upload className="w-12 h-12 mx-auto text-slate-400" />
                      <p className="font-medium text-sm">拖拽 PDF 到此处, 或点击选择</p>
                      <p className="text-xs text-muted-foreground">支持任意页数 · 自动检测格式</p>
                    </div>
                  )}
                </div>
                {converting && (
                  <div className="mt-4 space-y-2">
                    <Progress value={progress} className="h-2" />
                    <p className="text-xs text-center text-muted-foreground">
                      {progressMsg || `${progress}%`}
                    </p>
                  </div>
                )}
                <Button
                  className="w-full mt-4 bg-gradient-to-r from-rose-500 to-orange-500 hover:from-rose-600 hover:to-orange-600"
                  onClick={handleConvert}
                  disabled={!file || converting}
                >
                  {converting ? <Loader2 className="w-4 h-4 mr-2 animate-spin" /> : <ArrowRight className="w-4 h-4 mr-2" />}
                  {converting ? '转换中...' : '开始转换'}
                </Button>
                {error && (
                  <div className="mt-3 flex items-start gap-2 p-3 rounded-lg bg-red-50 dark:bg-red-950/30 text-red-700 dark:text-red-300 text-xs">
                    <AlertCircle className="w-4 h-4 flex-shrink-0 mt-0.5" />
                    <span className="break-all">{error}</span>
                  </div>
                )}
              </CardContent>
            </Card>

            {/* 选项 */}
            <Card>
              <CardHeader className="pb-3">
                <CardTitle className="text-base flex items-center gap-2">
                  <Settings className="w-4 h-4" /> 转换选项
                </CardTitle>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <Label htmlFor="pages" className="text-xs">页码范围 (可选)</Label>
                    <Input
                      id="pages"
                      placeholder="如 1-10 或 1,3,5"
                      value={pages}
                      onChange={(e) => setPages(e.target.value)}
                      className="text-sm h-9"
                    />
                  </div>
                  <div>
                    <Label className="text-xs">图片 DPI: {dpi}</Label>
                    <div className="pt-2">
                      <Slider
                        value={[dpi]}
                        onValueChange={(v) => setDpi(v[0])}
                        min={72}
                        max={400}
                        step={10}
                      />
                    </div>
                  </div>
                  <div>
                    <Label className="text-xs">目录级别: 1-{tocLevel}</Label>
                    <div className="pt-2">
                      <Slider
                        value={[tocLevel]}
                        onValueChange={(v) => setTocLevel(v[0])}
                        min={1}
                        max={9}
                        step={1}
                      />
                    </div>
                  </div>
                </div>
                <Separator />
                <div className="space-y-2.5">
                  <ToggleRow label="表格检测" desc="pdfplumber 精确提取" checked={detectTables} onChange={setDetectTables} icon={<TableIcon className="w-4 h-4" />} />
                  <ToggleRow label="列表/编号检测" desc="自动识别项目符号和编号" checked={detectLists} onChange={setDetectLists} icon={<ListOrdered className="w-4 h-4" />} />
                  <ToggleRow label="矢量图形还原" desc="DrawingML 浮动形状" checked={renderShapes} onChange={setRenderShapes} icon={<Layers className="w-4 h-4" />} />
                  <ToggleRow label="数学公式检测" desc="OMML 原生公式" checked={detectFormulas} onChange={setDetectFormulas} icon={<FunctionSquare className="w-4 h-4" />} />
                  <ToggleRow label="超链接/书签" desc="URL/Email自动识别" checked={detectLinks} onChange={setDetectLinks} icon={<Link2 className="w-4 h-4" />} />
                  <ToggleRow label="表单域识别" desc="签名/复选框/文本框" checked={detectForms} onChange={setDetectForms} icon={<ClipboardCheck className="w-4 h-4" />} />
                  <ToggleRow label="分页对齐" desc="标题与下段同页/表格行不分裂" checked={alignPages} onChange={setAlignPages} icon={<AlignJustify className="w-4 h-4" />} />
                  <ToggleRow label="页面填充对齐" desc="强制每页内容边界(可能增页)" checked={fillPages} onChange={setFillPages} icon={<AlignJustify className="w-4 h-4" />} />
                  <ToggleRow label="自动生成目录" desc="基于标题级别插入TOC字段" checked={generateToc} onChange={setGenerateToc} icon={<ListTree className="w-4 h-4" />} />
                  <ToggleRow label="PDF元数据还原" desc="标题/作者/日期/生产者" checked={restoreMetadata} onChange={setRestoreMetadata} icon={<FileHeart className="w-4 h-4" />} />
                  <ToggleRow label="PDF注释还原" desc="高亮/批注/删除线/波浪线" checked={detectAnnotations} onChange={setDetectAnnotations} icon={<Highlighter className="w-4 h-4" />} />
                  <ToggleRow label="保留分页符" desc="维持原 PDF 分页" checked={preservePageBreaks} onChange={setPreservePageBreaks} icon={<FileText className="w-4 h-4" />} />
                  <ToggleRow label="生成质量报告" desc="PDF/DOCX 对比评分" checked={withReport} onChange={setWithReport} icon={<CheckCircle2 className="w-4 h-4" />} />
                </div>
              </CardContent>
            </Card>

            {/* 历史记录面板 */}
            <HistoryPanel onRestore={(r) => { setResult(r); }} />
          </div>

          {/* 右侧: 结果 + 预览 */}
          <div className="lg:col-span-7 space-y-4">
            {result ? (
              <>
                {/* 统计卡片 */}
                <Card>
                  <CardHeader className="pb-3">
                    <div className="flex items-center justify-between">
                      <CardTitle className="text-base flex items-center gap-2">
                        <CheckCircle2 className="w-5 h-5 text-emerald-500" />
                        转换完成
                      </CardTitle>
                      {result.stats?.quality_score !== undefined && (
                        <Badge className={
                          result.stats.quality_score >= 85 ? 'bg-emerald-500' :
                          result.stats.quality_score >= 70 ? 'bg-amber-500' : 'bg-red-500'
                        }>
                          评分 {result.stats.quality_score}/100
                        </Badge>
                      )}
                    </div>
                  </CardHeader>
                  <CardContent>
                    <div className="grid grid-cols-3 sm:grid-cols-4 gap-3">
                      <StatCard icon={<FileText className="w-4 h-4" />} label="页数" value={result.stats?.pages ?? '-'} />
                      <StatCard icon={<Type className="w-4 h-4" />} label="段落" value={result.stats?.paragraphs ?? '-'} />
                      <StatCard icon={<TableIcon className="w-4 h-4" />} label="表格" value={result.stats?.tables ?? '-'} />
                      <StatCard icon={<ImageIcon className="w-4 h-4" />} label="图片" value={result.stats?.images ?? '-'} />
                      <StatCard icon={<ListOrdered className="w-4 h-4" />} label="列表" value={result.stats?.list_items ?? '-'} />
                      <StatCard icon={<Layers className="w-4 h-4" />} label="形状" value={result.stats?.shapes ?? '-'} />
                      <StatCard icon={<FunctionSquare className="w-4 h-4" />} label="公式" value={result.stats?.formulas ?? '-'} />
                      <StatCard icon={<Link2 className="w-4 h-4" />} label="链接" value={result.stats?.links ?? '-'} />
                      <StatCard icon={<ClipboardCheck className="w-4 h-4" />} label="表单域" value={result.stats?.form_fields ?? '-'} />
                      <StatCard icon={<ListTree className="w-4 h-4" />} label="目录条目" value={result.stats?.toc_entries ?? '-'} />
                      <StatCard icon={<FileHeart className="w-4 h-4" />} label="元数据" value={result.stats?.metadata_restored ?? '-'} />
                      <StatCard icon={<Highlighter className="w-4 h-4" />} label="注释" value={result.stats?.annotations ?? '-'} />
                    </div>
                    <div className="flex gap-2 mt-4">
                      <Button asChild className="flex-1 bg-gradient-to-r from-rose-500 to-orange-500 hover:from-rose-600 hover:to-orange-600">
                        <a href={result.docxUrl} download>
                          <Download className="w-4 h-4 mr-2" /> 下载 DOCX
                        </a>
                      </Button>
                      <Button asChild variant="outline" className="flex-1">
                        <a href={result.pdfUrl} download>
                          <Download className="w-4 h-4 mr-2" /> 下载原 PDF
                        </a>
                      </Button>
                    </div>
                  </CardContent>
                </Card>

                {/* 对比预览 */}
                <Card>
                  <CardHeader className="pb-3">
                    <CardTitle className="text-base flex items-center gap-2">
                      <Eye className="w-4 h-4" /> 可视化对比预览
                    </CardTitle>
                    <CardDescription className="text-xs">左: 原 PDF · 右: 转换后 DOCX</CardDescription>
                  </CardHeader>
                  <CardContent>
                    <div className="flex items-center gap-2 mb-3">
                      <Label className="text-xs">页码:</Label>
                      <Input
                        type="number"
                        min={1}
                        value={previewPage}
                        onChange={(e) => setPreviewPage(Math.max(1, parseInt(e.target.value) || 1))}
                        className="w-20 h-8 text-sm"
                      />
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={() => setPreviewPage(p => Math.max(1, p - 1))}
                      >上一页</Button>
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={() => setPreviewPage(p => p + 1)}
                      >下一页</Button>
                      <div className="ml-auto flex gap-1">
                        <Button
                          size="sm"
                          variant={previewSource === 'pdf' ? 'default' : 'outline'}
                          onClick={() => setPreviewSource('pdf')}
                        >原 PDF</Button>
                        <Button
                          size="sm"
                          variant={previewSource === 'docx' ? 'default' : 'outline'}
                          onClick={() => setPreviewSource('docx')}
                        >DOCX 渲染</Button>
                      </div>
                    </div>
                    <div className="grid grid-cols-2 gap-3">
                      <PreviewImage fileId={result.fileId} page={previewPage} source="pdf" />
                      <PreviewImage fileId={result.fileId} page={previewPage} source="docx" />
                    </div>
                  </CardContent>
                </Card>

                {/* 转换日志 */}
                <Card>
                  <CardHeader className="pb-3">
                    <CardTitle className="text-base">转换日志</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <ScrollArea className="h-40 w-full rounded border bg-slate-50 dark:bg-slate-900 p-3">
                      <pre className="text-xs font-mono whitespace-pre-wrap">{result.log}</pre>
                    </ScrollArea>
                  </CardContent>
                </Card>
              </>
            ) : (
              <Card className="h-full min-h-[400px] flex items-center justify-center">
                <CardContent className="text-center py-12">
                  <div className="w-20 h-20 rounded-full bg-gradient-to-br from-rose-100 to-orange-100 dark:from-rose-950 dark:to-orange-950 flex items-center justify-center mx-auto mb-4">
                    <FileText className="w-10 h-10 text-rose-500" />
                  </div>
                  <h3 className="text-lg font-semibold mb-2">等待转换</h3>
                  <p className="text-sm text-muted-foreground max-w-md mx-auto">
                    上传 PDF 文件后点击「开始转换」,<br />
                    将自动完成 1:1 格式复刻并生成对比预览
                  </p>
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mt-6 max-w-md mx-auto">
                    {[
                      { icon: <Type className="w-4 h-4" />, label: '字体复刻' },
                      { icon: <TableIcon className="w-4 h-4" />, label: '表格提取' },
                      { icon: <ImageIcon className="w-4 h-4" />, label: '印章定位' },
                      { icon: <ListOrdered className="w-4 h-4" />, label: '列表编号' },
                    ].map((f, i) => (
                      <div key={i} className="flex flex-col items-center gap-1 p-2 rounded-lg bg-slate-50 dark:bg-slate-900">
                        <span className="text-rose-500">{f.icon}</span>
                        <span className="text-xs">{f.label}</span>
                      </div>
                    ))}
                  </div>
                </CardContent>
              </Card>
            )}
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="border-t bg-white/80 dark:bg-slate-950/80 backdrop-blur-sm mt-auto">
        <div className="container mx-auto px-4 py-3 text-center text-xs text-muted-foreground">
          PDF→DOCX 企业级转换器 · 基于 PyMuPDF + pdfplumber + python-docx 自研 · 不依赖 pdf2docx
        </div>
      </footer>
    </div>
  );
}

function ToggleRow({ label, desc, checked, onChange, icon }: {
  label: string; desc: string; checked: boolean; onChange: (v: boolean) => void; icon: React.ReactNode;
}) {
  return (
    <div className="flex items-center justify-between gap-2 py-1">
      <div className="flex items-center gap-2">
        <span className="text-muted-foreground">{icon}</span>
        <div>
          <p className="text-sm font-medium">{label}</p>
          <p className="text-xs text-muted-foreground">{desc}</p>
        </div>
      </div>
      <Switch checked={checked} onCheckedChange={onChange} />
    </div>
  );
}

function StatCard({ icon, label, value }: { icon: React.ReactNode; label: string; value: any }) {
  return (
    <div className="rounded-lg border bg-slate-50 dark:bg-slate-900 p-2.5 text-center">
      <div className="flex justify-center text-rose-500 mb-1">{icon}</div>
      <div className="text-lg font-bold leading-none">{value}</div>
      <div className="text-xs text-muted-foreground mt-1">{label}</div>
    </div>
  );
}

function PreviewImage({ fileId, page, source }: { fileId: string; page: number; source: 'pdf' | 'docx' }) {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(false);
  const url = `/api/pdf2docx/preview?f=${fileId}&p=${page}&s=${source}&t=${Date.now()}`;
  useEffect(() => { setLoading(true); setError(false); }, [url]);
  return (
    <div className="relative aspect-[3/4] rounded-lg border bg-slate-100 dark:bg-slate-900 overflow-hidden">
      {loading && (
        <div className="absolute inset-0 flex items-center justify-center">
          <Loader2 className="w-6 h-6 animate-spin text-muted-foreground" />
        </div>
      )}
      {error ? (
        <div className="absolute inset-0 flex items-center justify-center text-xs text-muted-foreground p-4 text-center">
          {source === 'docx' ? 'DOCX 渲染中, 请稍候...' : '加载失败'}
        </div>
      ) : (
        <img
          src={url}
          alt={`${source} page ${page}`}
          className="w-full h-full object-contain"
          onLoad={() => setLoading(false)}
          onError={() => { setLoading(false); setError(true); }}
        />
      )}
      <div className="absolute top-2 left-2">
        <Badge variant="secondary" className="text-xs">
          {source === 'pdf' ? '原 PDF' : 'DOCX'}
        </Badge>
      </div>
    </div>
  );
}

function HistoryPanel({ onRestore }: { onRestore: (r: any) => void }) {
  const [history, setHistory] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  const fetchHistory = useCallback(async () => {
    try {
      const res = await fetch('/api/pdf2docx/history');
      const data = await res.json();
      setHistory(data.history || []);
    } catch (e) {
      // 忽略
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { fetchHistory(); }, [fetchHistory]);

  const clearHistory = async () => {
    try {
      await fetch('/api/pdf2docx/history', { method: 'DELETE' });
      setHistory([]);
    } catch (e) {}
  };

  const formatTime = (iso: string) => {
    try {
      const d = new Date(iso);
      const now = new Date();
      const diff = (now.getTime() - d.getTime()) / 1000;
      if (diff < 60) return '刚刚';
      if (diff < 3600) return `${Math.floor(diff/60)}分钟前`;
      if (diff < 86400) return `${Math.floor(diff/3600)}小时前`;
      return d.toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
    } catch { return ''; }
  };

  const formatSize = (bytes: number) => {
    if (!bytes) return '-';
    if (bytes < 1024) return `${bytes}B`;
    if (bytes < 1024*1024) return `${(bytes/1024).toFixed(1)}KB`;
    return `${(bytes/1024/1024).toFixed(2)}MB`;
  };

  return (
    <Card>
      <CardHeader className="pb-3">
        <div className="flex items-center justify-between">
          <CardTitle className="text-base flex items-center gap-2">
            <History className="w-4 h-4" />
            转换历史
            {history.length > 0 && (
              <Badge variant="secondary" className="text-xs">{history.length}</Badge>
            )}
          </CardTitle>
          {history.length > 0 && (
            <Button size="sm" variant="ghost" onClick={clearHistory} className="h-7 text-xs text-muted-foreground hover:text-destructive">
              <Trash2 className="w-3 h-3 mr-1" /> 清空
            </Button>
          )}
        </div>
      </CardHeader>
      <CardContent>
        {loading ? (
          <div className="flex items-center justify-center py-6 text-xs text-muted-foreground">
            <Loader2 className="w-4 h-4 mr-2 animate-spin" /> 加载中...
          </div>
        ) : history.length === 0 ? (
          <div className="text-center py-6 text-xs text-muted-foreground">
            <Clock className="w-8 h-8 mx-auto mb-2 opacity-30" />
            暂无转换记录
          </div>
        ) : (
          <div className="space-y-2 max-h-64 overflow-y-auto">
            {history.map((h, i) => (
              <div
                key={i}
                className="flex items-center gap-2 p-2 rounded-lg border bg-slate-50 dark:bg-slate-900 hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer transition-colors"
                onClick={() => onRestore(h)}
              >
                <div className="flex-1 min-w-0">
                  <p className="text-xs font-medium truncate">{h.fileName}</p>
                  <div className="flex items-center gap-2 mt-0.5">
                    <span className="text-[10px] text-muted-foreground">{formatTime(h.timestamp)}</span>
                    <span className="text-[10px] text-muted-foreground">·</span>
                    <span className="text-[10px] text-muted-foreground">{formatSize(h.fileSize)}</span>
                    {h.stats?.pages && (
                      <>
                        <span className="text-[10px] text-muted-foreground">·</span>
                        <span className="text-[10px] text-muted-foreground">{h.stats.pages}页</span>
                      </>
                    )}
                  </div>
                </div>
                {h.qualityScore && (
                  <Badge className={
                    h.qualityScore >= 85 ? 'bg-emerald-500 text-xs' :
                    h.qualityScore >= 70 ? 'bg-amber-500 text-xs' : 'bg-red-500 text-xs'
                  }>
                    {h.qualityScore}
                  </Badge>
                )}
                <a
                  href={h.docxUrl}
                  download
                  onClick={(e) => e.stopPropagation()}
                  className="text-rose-500 hover:text-rose-600"
                >
                  <Download className="w-3.5 h-3.5" />
                </a>
              </div>
            ))}
          </div>
        )}
      </CardContent>
    </Card>
  );
}
