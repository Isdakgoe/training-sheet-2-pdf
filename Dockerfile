FROM python:3.12-slim
 
# LibreOffice(PPTX->PDF変換)と日本語フォントを導入
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*
 
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py .
 
# Linuxには Meiryo が無いため Noto Sans CJK JP を使用
ENV PPTX_FONT="Noto Sans CJK JP"
ENV PORT=8080
EXPOSE 8080
 
CMD streamlit run app.py --server.port=${PORT} --server.address=0.0.0.0 --server.headless=true
