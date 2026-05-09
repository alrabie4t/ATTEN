-- Migration Script: نظام الحضور والانصراف v6
-- شغّل ده على PostgreSQL قبل ما تستخدم السيستم

-- 1. جدول الشركات (Tenants)
CREATE TABLE IF NOT EXISTS tenants (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    name VARCHAR(255) NOT NULL,
    subdomain VARCHAR(100) UNIQUE NOT NULL,
    address TEXT,
    phone VARCHAR(50),
    email VARCHAR(255),
    tax_number VARCHAR(100),
    office_latitude DECIMAL(10, 8),
    office_longitude DECIMAL(11, 8),
    geofence_radius_meters INTEGER DEFAULT 300,
    work_start_time TIME DEFAULT '09:00:00',
    work_end_time TIME DEFAULT '17:00:00',
    late_threshold_minutes INTEGER DEFAULT 15,
    currency VARCHAR(50) DEFAULT 'جنيه',
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- 2. جدول الأقسام
CREATE TABLE IF NOT EXISTS departments (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    tenant_id UUID REFERENCES tenants(id) ON DELETE CASCADE,
    name VARCHAR(255) NOT NULL,
    description TEXT,
    manager_id UUID,
    work_start_time TIME,
    work_end_time TIME,
    late_threshold_minutes INTEGER,
    is_flexible BOOLEAN DEFAULT FALSE,
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT NOW()
);

-- 3. جدول الموظفين
CREATE TABLE IF NOT EXISTS users (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    tenant_id UUID REFERENCES tenants(id) ON DELETE CASCADE,
    department_id UUID REFERENCES departments(id) ON DELETE SET NULL,
    full_name VARCHAR(255) NOT NULL,
    username VARCHAR(100) NOT NULL,
    password_hash VARCHAR(255) NOT NULL,
    role VARCHAR(50) DEFAULT 'employee',
    job_title VARCHAR(255),
    employee_code VARCHAR(100),
    phone VARCHAR(50),
    national_id VARCHAR(100),
    hire_date DATE,
    hourly_rate DECIMAL(10, 2) DEFAULT 0.0,
    base_salary DECIMAL(10, 2) DEFAULT 0.0,
    salary_type VARCHAR(50) DEFAULT 'hourly',
    is_active BOOLEAN DEFAULT TRUE,
    last_login_at TIMESTAMP,
    created_at TIMESTAMP DEFAULT NOW()
);

-- 4. جدول الحضور والانصراف
CREATE TABLE IF NOT EXISTS attendance (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    tenant_id UUID REFERENCES tenants(id) ON DELETE CASCADE,
    user_id UUID REFERENCES users(id) ON DELETE CASCADE,
    attendance_date DATE NOT NULL,
    check_in_time TIMESTAMP,
    check_out_time TIMESTAMP,
    check_in_latitude DECIMAL(10, 8),
    check_in_longitude DECIMAL(11, 8),
    check_out_latitude DECIMAL(10, 8),
    check_out_longitude DECIMAL(11, 8),
    status VARCHAR(50) DEFAULT 'present',
    late_minutes INTEGER DEFAULT 0,
    work_hours DECIMAL(5, 2),
    is_remote BOOLEAN DEFAULT FALSE,
    is_manual BOOLEAN DEFAULT FALSE,
    notes TEXT,
    session_number INTEGER DEFAULT 1,
    created_by UUID REFERENCES users(id) ON DELETE SET NULL,
    created_at TIMESTAMP DEFAULT NOW()
);

-- 5. جدول التسويات اليومية
CREATE TABLE IF NOT EXISTS daily_adjustments (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    tenant_id UUID REFERENCES tenants(id) ON DELETE CASCADE,
    user_id UUID REFERENCES users(id) ON DELETE CASCADE,
    adjustment_date DATE NOT NULL,
    bonus DECIMAL(10, 2) DEFAULT 0.0,
    deduction DECIMAL(10, 2) DEFAULT 0.0,
    reason TEXT,
    created_by UUID REFERENCES users(id) ON DELETE SET NULL,
    created_at TIMESTAMP DEFAULT NOW(),
    UNIQUE(tenant_id, user_id, adjustment_date)
);

-- 6. جدول السلف
CREATE TABLE IF NOT EXISTS advances (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    tenant_id UUID REFERENCES tenants(id) ON DELETE CASCADE,
    user_id UUID REFERENCES users(id) ON DELETE CASCADE,
    amount DECIMAL(10, 2) NOT NULL,
    advance_date DATE NOT NULL,
    reason TEXT,
    notes TEXT,
    status VARCHAR(50) DEFAULT 'approved',
    approved_by UUID REFERENCES users(id) ON DELETE SET NULL,
    approved_at TIMESTAMP,
    created_at TIMESTAMP DEFAULT NOW()
);

-- 7. جدول تقارير الرواتب
CREATE TABLE IF NOT EXISTS salary_reports (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    tenant_id UUID REFERENCES tenants(id) ON DELETE CASCADE,
    user_id UUID REFERENCES users(id) ON DELETE CASCADE,
    report_name VARCHAR(255),
    report_from DATE NOT NULL,
    report_to DATE NOT NULL,
    total_days INTEGER DEFAULT 0,
    total_sessions INTEGER DEFAULT 0,
    total_hours DECIMAL(8, 2) DEFAULT 0.0,
    hourly_rate DECIMAL(10, 2) DEFAULT 0.0,
    base_salary DECIMAL(10, 2) DEFAULT 0.0,
    salary_type VARCHAR(50),
    gross_salary DECIMAL(10, 2) DEFAULT 0.0,
    total_bonus DECIMAL(10, 2) DEFAULT 0.0,
    total_deduction DECIMAL(10, 2) DEFAULT 0.0,
    total_advances DECIMAL(10, 2) DEFAULT 0.0,
    net_salary DECIMAL(10, 2) DEFAULT 0.0,
    notes TEXT,
    created_by UUID REFERENCES users(id) ON DELETE SET NULL,
    created_at TIMESTAMP DEFAULT NOW()
);

-- 8. بيانات افتراضية (Tenant + Admin)
INSERT INTO tenants (name, subdomain, is_active) 
VALUES ('شركة النور', 'alnoor-tech', TRUE)
ON CONFLICT DO NOTHING;

INSERT INTO users (tenant_id, full_name, username, password_hash, role, is_active)
SELECT 
    (SELECT id FROM tenants WHERE subdomain = 'alnoor-tech'),
    'مدير النظام',
    'admin',
    '$2b$12$LQv3c1yqBWVHxkd0LHAkCOYz6TtxMQJqhN8/LewKyNiAYMyzJ/I1O', -- bcrypt hash for "Admin@123"
    'admin',
    TRUE
WHERE NOT EXISTS (SELECT 1 FROM users WHERE username = 'admin');

-- Indexes للأداء
CREATE INDEX IF NOT EXISTS idx_attendance_user_date ON attendance(user_id, attendance_date);
CREATE INDEX IF NOT EXISTS idx_attendance_tenant_date ON attendance(tenant_id, attendance_date);
CREATE INDEX IF NOT EXISTS idx_adjustments_user_date ON daily_adjustments(user_id, adjustment_date);
CREATE INDEX IF NOT EXISTS idx_advances_user ON advances(user_id);
CREATE INDEX IF NOT EXISTS idx_salary_reports_user ON salary_reports(user_id);
