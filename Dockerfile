FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    REMBG_MODEL=u2netp \
    REMBG_HOME=/data/rembg \
    U2NET_HOME=/data/rembg \
    OCR_LANG=ara+eng \
    TMPDIR=/tmp \
    NUMBA_CACHE_DIR=/tmp/numba \
    XDG_CACHE_HOME=/tmp/cache

RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-dejavu-core \
    libgl1 \
    libglib2.0-0 \
    libmagic1 \
    tesseract-ocr \
    tesseract-ocr-ara \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install --upgrade pip && pip install -r requirements.txt

RUN useradd --create-home --uid 10001 --shell /usr/sbin/nologin botuser \
    && mkdir -p /data/rembg \
    && chown -R botuser:botuser /data

COPY --chown=botuser:botuser . .
RUN chmod +x /app/docker-entrypoint.sh

USER botuser

ENTRYPOINT ["/app/docker-entrypoint.sh"]
CMD ["python", "bot.py"]

HEALTHCHECK --interval=30s --timeout=5s --start-period=90s --retries=3 \
    CMD ["python", "-c", "from pathlib import Path; import sys; sys.exit(0 if b'bot.py' in Path('/proc/1/cmdline').read_bytes() else 1)"]
