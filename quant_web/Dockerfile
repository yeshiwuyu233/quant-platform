FROM python:3.11-slim

# 安装 Chromium 浏览器及 DrissionPage 所需系统依赖
RUN apt-get update && apt-get install -y \
    chromium \
    wget \
    gnupg \
    ca-certificates \
    fonts-liberation \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdrm2 \
    libgbm1 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libu2f-udev \
    libvulkan1 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxkbcommon0 \
    libxrandr2 \
    xdg-utils \
    && rm -rf /var/lib/apt/lists/*

# 安装 Python 依赖
COPY quant_web/requirements.txt /tmp/requirements.txt
RUN pip install --no-cache-dir -r /tmp/requirements.txt

# 复制项目代码（排除 xlsx 等大文件，见 .dockerignore）
COPY . /app/
WORKDIR /app

RUN chmod +x /app/quant_web/entrypoint.sh

# 默认启动 Flask (gunicorn)
EXPOSE 5000
ENTRYPOINT ["/app/quant_web/entrypoint.sh"]
CMD ["web"]
