# syntax=docker/dockerfile:1.7
FROM python:3.12-slim-bookworm

LABEL org.opencontainers.image.title="hero-digital-etl" \
      org.opencontainers.image.description="Hero Digital ETL FastAPI service"

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONUTF8=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    LANG=C.UTF-8 \
    LC_ALL=C.UTF-8

WORKDIR /app

RUN groupadd --gid 10001 app \
    && useradd --uid 10001 --gid app --create-home --shell /usr/sbin/nologin app

COPY requirements.txt ./
RUN python -m pip install --upgrade pip \
    && python -m pip install --no-cache-dir -r requirements.txt

# 按部署要求将 .env 直接打入镜像；该镜像必须只存放在受控私有仓库。
COPY --chown=app:app api.py run.py .env ./
COPY --chown=app:app etl ./etl
COPY --chown=app:app resources ./resources

RUN mkdir -p /app/output && chown app:app /app/output

USER app

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD ["python", "-c", "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=3).read()"]

# 必须保持单个 API worker；任务级并行由服务内部的隔离子进程池负责。
CMD ["python", "-m", "uvicorn", "api:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
