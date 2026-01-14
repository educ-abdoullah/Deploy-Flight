FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV RENDER=true

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Installe Chromium pour Playwright
RUN python -m playwright install --with-deps chromium

COPY . .

CMD gunicorn -b 0.0.0.0:$PORT app:app
