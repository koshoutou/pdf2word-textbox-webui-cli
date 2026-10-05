FROM node:18-slim AS web-builder

# 构建Web UI
WORKDIR /app/web-ui
COPY web-ui/package.json ./
RUN npm install --legacy-peer-deps
COPY web-ui/ ./
RUN npm run build

# 最终镜像
FROM python:3.12-slim

# 安装系统依赖
RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1-mesa-glx \
    libglib2.0-0 \
    libsm6 \
    libxext6 \
    libxrender-dev \
    libfontconfig1 \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# 安装Python依赖
RUN pip install --no-cache-dir PyMuPDF python-docx pdfplumber lxml Pillow aiohttp

WORKDIR /app

# 复制转换器代码
COPY src/ ./src/
COPY main.py batch.py api_server.py setup_fonts.sh ./
COPY README.md FONTS.md API.md CODE_REVIEW.md LICENSE ./

# 复制Web UI构建产物
COPY --from=web-builder /app/web-ui/.next ./.next
COPY --from=web-builder /app/web-ui/public ./public
COPY --from=web-builder /app/web-ui/package.json ./web-ui.json

# 创建输出目录
RUN mkdir -p output

# 安装字体
RUN chmod +x setup_fonts.sh && ./setup_fonts.sh || true

EXPOSE 3000 8000

# 启动Web UI和API服务
CMD ["sh", "-c", "python3 api_server.py --port 8000 & exec npx next start -p 3000"]
