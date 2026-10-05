# API 接口文档

> PDF→DOCX 转换器提供 RESTful API，支持其他程序通过 HTTP 调用。

## 启动 API 服务

### 方式1: 独立Python服务 (推荐)

```bash
# 安装依赖
pip install aiohttp

# 启动服务
python3 api_server.py --port 8000

# 或指定host
python3 api_server.py --host 0.0.0.0 --port 8000
```

### 方式2: 通过Next.js API (Web UI内置)

Web UI已内置所有API路由，启动Web UI即可使用：

```bash
cd web-ui
npm install
npm run dev  # 默认端口3000
```

## API 接口列表

### 1. 转换 PDF → DOCX

```
POST /api/convert
Content-Type: multipart/form-data
```

**请求参数:**

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `file` | File | 是 | PDF文件 |
| `pages` | String | 否 | 页码范围，如 `1-10` 或 `1,3,5-8` |
| `dpi` | String | 否 | 图片DPI，默认 `150` |
| `detectTables` | String | 否 | 表格检测，`true`/`false`，默认 `true` |
| `detectLists` | String | 否 | 列表检测，`true`/`false`，默认 `true` |
| `detectFormulas` | String | 否 | 公式检测，`true`/`false`，默认 `true` |
| `detectLinks` | String | 否 | 超链接检测，`true`/`false`，默认 `true` |
| `detectForms` | String | 否 | 表单域检测，`true`/`false`，默认 `true` |
| `generateToc` | String | 否 | TOC目录，`true`/`false`，默认 `true` |
| `restoreMetadata` | String | 否 | 元数据还原，`true`/`false`，默认 `true` |
| `detectAnnotations` | String | 否 | 注释还原，`true`/`false`，默认 `true` |
| `alignPages` | String | 否 | 分页对齐，`true`/`false`，默认 `true` |
| `password` | String | 否 | 加密PDF密码 |

**响应:**

```json
{
  "success": true,
  "fileId": "uuid-string",
  "fileName": "input.pdf",
  "stats": {
    "pages": 97,
    "paragraphs": 1048,
    "tables": 43,
    "images": 4,
    "list_items": 36,
    "formulas": 7,
    "links": 2,
    "form_fields": 16,
    "toc_entries": 69,
    "metadata_restored": 5,
    "annotations": 0,
    "quality_score": 86.0
  },
  "downloadUrl": "/api/download/uuid-string"
}
```

**调用示例:**

```bash
# curl
curl -X POST http://localhost:8000/api/convert \
  -F "file=@input.pdf" \
  -F "pages=1-20" \
  -F "dpi=200"

# Python (requests)
import requests
files = {'file': open('input.pdf', 'rb')}
data = {'pages': '1-20', 'dpi': '200'}
resp = requests.post('http://localhost:8000/api/convert', files=files, data=data)
result = resp.json()
print(result['stats'])
```

### 2. 下载转换后的 DOCX

```
GET /api/download/:fileId
```

**示例:**
```bash
curl -o output.docx http://localhost:8000/api/download/uuid-string
```

### 3. 服务状态

```
GET /api/status
```

**响应:**
```json
{
  "service": "pdf2docx-converter",
  "version": "1.0.0",
  "status": "running",
  "modules": 16,
  "endpoints": ["POST /api/convert", "GET /api/download/:fileId", ...]
}
```

### 4. 已转换文件列表

```
GET /api/files
```

**响应:**
```json
{
  "files": [
    {
      "fileId": "uuid-string",
      "name": "uuid-string.docx",
      "size": 406443,
      "mtime": 1696000000
    }
  ]
}
```

### 5. 清空已转换文件

```
DELETE /api/files
```

### 6. 进度推送 (SSE)

```
GET /api/progress?jobId=:jobId
```

通过 Server-Sent Events 实时推送转换进度。

### 7. 页面预览

```
GET /api/preview?f=:fileId&p=:page&s=:source
```

渲染指定页为PNG图片，`source` 为 `pdf` 或 `docx`。

## Python SDK 调用

```python
import requests

class PDF2DocxClient:
    def __init__(self, base_url='http://localhost:8000'):
        self.base_url = base_url

    def convert(self, pdf_path, **options):
        """转换PDF为DOCX"""
        files = {'file': open(pdf_path, 'rb')}
        data = {k: str(v) for k, v in options.items()}
        resp = requests.post(f'{self.base_url}/api/convert', files=files, data=data)
        return resp.json()

    def download(self, file_id, output_path):
        """下载转换后的DOCX"""
        resp = requests.get(f'{self.base_url}/api/download/{file_id}')
        with open(output_path, 'wb') as f:
            f.write(resp.content)

    def status(self):
        """获取服务状态"""
        return requests.get(f'{self.base_url}/api/status').json()

# 使用示例
client = PDF2DocxClient('http://localhost:8000')
result = client.convert('input.pdf', pages='1-20', dpi=200)
print(f"质量评分: {result['stats']['quality_score']}")
client.download(result['fileId'], 'output.docx')
```

## 命令行调用

无需启动API服务，直接调用Python模块：

```bash
# 单文件
python3 main.py -i input.pdf -o output.docx

# 批量
python3 batch.py -i /input/dir -o /output/dir -w 4

# 加密PDF
python3 main.py -i encrypted.pdf -o output.docx --password YOUR_PASSWORD
```

## 错误码

| HTTP状态码 | 说明 |
|-----------|------|
| 200 | 成功 |
| 400 | 参数错误 |
| 404 | 文件不存在 |
| 500 | 转换失败 |
