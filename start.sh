#!/bin/bash
# سكربت تشغيل السيستم على Render أو أي خادم

echo "🚀 Starting Attendance System v6..."
echo "📡 PORT: $PORT"

# تشغيل uvicorn
uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1
