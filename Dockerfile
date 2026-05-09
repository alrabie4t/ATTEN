# Python 3.11 صراحة — مش هنستخدم latest
FROM python:3.11-slim

# تثبيت system dependencies للـ asyncpg و bcrypt
RUN apt-get update && apt-get install -y     gcc     libpq-dev     && rm -rf /var/lib/apt/lists/*

# مجلد العمل
WORKDIR /app

# نسخ المتطلبات أولاً (للـ cache)
COPY requirements.txt .

# تثبيت المتطلبات
RUN pip install --no-cache-dir --upgrade pip &&     pip install --no-cache-dir -r requirements.txt

# نسخ باقي الملفات
COPY . .

# البورت
ENV PORT=8080
EXPOSE 8080

# تشغيل التطبيق
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080"]
