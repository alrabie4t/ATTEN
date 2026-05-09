# نظام الحضور والانصراف v6 — تعليمات الرفع على Render

## الخطوات خطوة بخطوة1

### 1. إنشاء حساب على Render
- ادخل على https://render.com
- سجل بحساب GitHub

### 2. رفع الكود على GitHub
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/attendance-system.git
git push -u origin main
```

### 3. إنشاء قاعدة بيانات PostgreSQL
1. في Render Dashboard، اضغط **New** → **PostgreSQL**
2. اختر **Free Plan**
3. اكتب اسم: `attendance-db`
4. اضغط **Create Database**
5. انسخ **Internal Database URL** — ده هنحتاجه

### 4. إنشاء Web Service
1. اضغط **New** → **Web Service**
2. اربط repo بتاعك من GitHub
3. املأ البيانات:
   - **Name**: `attendance-system`
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT`
4. في **Environment Variables**:
   - `DATABASE_URL`: (اللي نسخته من قاعدة البيانات)
   - `JWT_SECRET`: اكتب أي نص عشوائي طويل (مثلاً 32 حرف)
5. اضغط **Create Web Service**

### 5. إنشاء Static Site للـ Frontend
1. اضغط **New** → **Static Site**
2. اربط نفس الـ repo
3. املأ البيانات:
   - **Name**: `attendance-frontend`
   - **Build Command**: (فارغ)
   - **Publish Directory**: `.` (أو `/`)
4. في **Environment Variables**:
   - `API_URL`: رابط الـ Backend (مثلاً `https://attendance-system.onrender.com`)
5. اضغط **Create Static Site**

### 6. تحديث CORS في Backend (لو مشتغلش)
لو ظهرت مشكلة CORS، عدّل في `main.py`:
```python
allow_origins=["https://attendance-frontend.onrender.com", "http://localhost:8000"]
```

### 7. إنشاء الجداول في قاعدة البيانات
```bash
# استخدم Render Shell أو PostgreSQL client
# شغّل migration script (لو موجود)
# أو أنشئ الجداول يدوياً من SQL
```

### ملاحظات مهمة
- **Free tier** على Render: Web service ينام بعد 15 دقيقة عدم استخدام (يصحى تلقائياً)
- **PostgreSQL Free**: 1GB storage
- **Static Site**: مجاني دائم بدون قيود
- لو عايز تتجنب sleep: اشتري paid plan ($7/شهر)

### روابط مهمة
- Render Dashboard: https://dashboard.render.com
- Docs: https://render.com/docs
