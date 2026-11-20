FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    HOME=/tmp

# LibreOffice does PDF export and legacy-format conversion; fontconfig makes
# the brand font visible to it.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer libreoffice-calc libreoffice-impress \
        fonts-dejavu-core fontconfig curl \
    && rm -rf /var/lib/apt/lists/*

# The brand font, so PDFs never silently fall back to DejaVu.
COPY assets/fonts/Onest-Regular.ttf assets/fonts/Onest-Bold.ttf /usr/share/fonts/truetype/onest/
RUN fc-cache -f

WORKDIR /app

COPY pyproject.toml README.md ./
COPY core ./core
RUN pip install --no-cache-dir .

COPY app.py cli.py ./
COPY .streamlit ./.streamlit
COPY templates ./templates
COPY assets ./assets
COPY SapientLogoNoText.png ./

RUN useradd -m -u 10001 app && chown -R app:app /app
USER app

EXPOSE 8501
HEALTHCHECK --interval=30s --timeout=5s --start-period=30s \
    CMD curl -fsS http://localhost:8501/_stcore/health || exit 1

CMD ["streamlit", "run", "app.py", \
     "--server.port=8501", "--server.address=0.0.0.0", \
     "--server.maxUploadSize=100", "--server.enableXsrfProtection=true", \
     "--browser.gatherUsageStats=false"]
