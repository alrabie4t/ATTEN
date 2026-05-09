# استخدم Python 3.11
FROM python:3.11-slim

# مجلد العمل
WORKDIR /app

# نسخ ملف المتطلبات أولاً (للـ cache)
COPY requirements.txt .

# تثبيت المتطلبات
RUN pip install --no-cache-dir -r requirements.txt

# نسخ باقي الملفات
COPY . .

# البورت اللي هيسمع عليه Railway
ENV PORT=8080
EXPOSE 8080

# تشغيل التطبيق
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080"]
