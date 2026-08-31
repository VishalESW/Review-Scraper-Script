FROM python:3.12-slim

WORKDIR /app

# Install deps first (better layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App code + assets (fonts, tool page, about video)
COPY . .

# Bind to all interfaces inside the container; no desktop browser on a server.
ENV HOST=0.0.0.0 \
    PORT=8765 \
    NO_BROWSER=1

EXPOSE 8765

CMD ["python", "app.py"]
