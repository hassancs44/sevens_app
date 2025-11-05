from flask import Flask, request, jsonify, render_template, send_from_directory
from flask_cors import CORS
import pandas as pd
import os
from datetime import datetime
import requests
import re




# ✅ تعريف المجلد الأساسي للمشروع
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

import sqlite3

DB_SQLITE = os.path.join(BASE_DIR, "sevens.db")

def init_sqlite():
    conn = sqlite3.connect(DB_SQLITE)
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        role TEXT,
        password TEXT,
        email TEXT UNIQUE,
        department TEXT,
        status TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        req_id TEXT UNIQUE,
        date TEXT,
        title TEXT,
        description TEXT,
        sender_dept TEXT,
        receiver_dept TEXT,
        status TEXT,
        assigned_to TEXT,
        updated_by TEXT,
        duration TEXT,
        file_name TEXT
    )
    """)

    conn.commit()
    conn.close()
    print("✅ SQLite database initialized and ready.")


# ✅ إنشاء مجلد الرفع
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ✅ مسار ملف دردشة الطلبات
CHAT_PATH = os.path.join(BASE_DIR, "chat_messages.xlsx")

def load_chats():
    """تحميل سجل دردشات الطلبات من ملف Excel أو إنشاؤه إن لم يوجد"""
    if not os.path.exists(CHAT_PATH):
        df = pd.DataFrame(columns=['رقم الطلب', 'المرسل', 'القسم', 'الرسالة', 'الملف', 'الوقت'])
        df.to_excel(CHAT_PATH, index=False)
        print("✅ Created chat_messages.xlsx")
        return df
    try:
        df = pd.read_excel(CHAT_PATH)
        # تنظيف الأعمدة وتوحيد الأسماء
        df.columns = [str(c).strip() for c in df.columns]
        for col in ['رقم الطلب', 'المرسل', 'القسم', 'الرسالة', 'الملف', 'الوقت']:
            if col not in df.columns:
                df[col] = ''
        return df
    except Exception as e:
        print("❌ load_chats error:", e)
        return pd.DataFrame(columns=['رقم الطلب', 'المرسل', 'القسم', 'الرسالة', 'الملف', 'الوقت'])

def normalize_arabic(text):
    """توحيد النصوص العربية لتفادي اختلاف الهمزات والمسافات"""
    if not isinstance(text, str):
        text = str(text)
    text = text.strip()
    text = re.sub(r'[إأآا]', 'ا', text)  # توحيد الألف والهمزات
    text = re.sub(r'\s+', '', text)      # إزالة كل المسافات
    text = text.replace('ة','ه')         # توحيد التاء المربوطة مع الهاء
    return text

# ============== إعدادات عامة ==============
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "database.xlsx")
REQUESTS_PATH = os.path.join(BASE_DIR, "requests.xlsx")
REQUESTS_SHEET = "الطلبات جميع"
EXPORT_DIR = os.path.join(BASE_DIR, "exports")
os.makedirs(EXPORT_DIR, exist_ok=True)

## مفتاح واجهة OpenRouter API  (احصل عليه من https://openrouter.ai)
OPENROUTER_API_KEY = "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"


app = Flask(__name__, template_folder='templates', static_folder='static')
CORS(app, resources={r"/api/*": {"origins": "*"}})

# ============== دوال مساعدة ==============
def ensure_excel_exists():
    if not os.path.exists(DB_PATH):
        users_cols = ['الاسم', 'الصلاحية', 'كلمة المرور', 'البريد الإلكتروني', 'القسم']
        pd.DataFrame(columns=users_cols).to_excel(DB_PATH, index=False)
        print("✅ Created users DB")

    if not os.path.exists(REQUESTS_PATH):
        req_cols = ['رقم الطلب', 'التاريخ', 'العنوان', 'الوصف', 'القسم المرسل',
                    'القسم المستلم', 'الحالة', 'الموظف المعين', 'آخر تحديث بواسطة', 'الوقت', 'الملف']
        pd.DataFrame(columns=req_cols).to_excel(REQUESTS_PATH, index=False, sheet_name=REQUESTS_SHEET)
        print("✅ Created requests DB")
    else:
        print("📂 Excel files already exist ✅")

# ✅ استدعِها مرة واحدة عند بدء التشغيل
ensure_excel_exists()


def normalize_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df

def load_users():
    try:
        df = pd.read_excel(DB_PATH)

        # 🔹 تنظيف الأعمدة من أي رموز أو فراغات غريبة
        df.columns = (
            df.columns
            .astype(str)
            .str.replace('\u200f', '', regex=True)
            .str.replace('\u200e', '', regex=True)
            .str.replace(' ', '', regex=True)
            .str.strip()
        )

        # ✅ توحيد أسماء الأعمدة مهما كانت كتابتها
        rename_map = {
            'الاسم': 'الاسم',
            'الاسمالكامل': 'الاسم',
            'الاسم_الكامل': 'الاسم',
            'الا سم': 'الاسم',
            'الإسم': 'الاسم',
            'اسم': 'الاسم',

            'البريدالإلكتروني': 'البريد الإلكتروني',
            'البريدالالكتروني': 'البريد الإلكتروني',
            'البريدالالكترونى': 'البريد الإلكتروني',
            'الايميل': 'البريد الإلكتروني',
            'email': 'البريد الإلكتروني',
            'ايميل': 'البريد الإلكتروني',

            'القسم': 'القسم',
            'القسم_الموظف': 'القسم',
            'ادارة': 'القسم',

            'الصلاحيه': 'الصلاحية',
            'الوظيفة': 'الصلاحية',
            'role': 'الصلاحية'
        }

        # 🧩 إعادة التسمية بناءً على التطابق الجزئي (حتى لو ناقص حرف)
        for col in list(df.columns):
            normalized = re.sub(r'[إأآا]', 'ا', col).replace(' ', '').lower()
            for k, v in rename_map.items():
                if re.sub(r'[إأآا]', 'ا', k).replace(' ', '').lower() in normalized:
                    df.rename(columns={col: v}, inplace=True)

        # ✅ التأكد أن كل الأعمدة المهمة موجودة حتى لو ناقصة
        for col in ['الاسم', 'البريد الإلكتروني', 'القسم', 'الصلاحية', 'كلمة المرور']:
            if col not in df.columns:
                df[col] = ''

        return normalize_department_names(df)
    except Exception as e:
        print("❌ load_users error:", e)
        return pd.DataFrame()


def normalize_department_names(df):
    """توحيد أسماء الأقسام داخل قاعدة المستخدمين"""
    if 'القسم' in df.columns:
        df['القسم'] = (
            df['القسم']
            .astype(str)
            .str.strip()
            .str.replace('\u200f','', regex=True)
            .str.replace('\u200e','', regex=True)
            .str.replace('  ',' ', regex=True)
            .str.replace('الادارة','إدارة', regex=False)
        )
    return df

def load_requests():
    try:
        if not os.path.exists(REQUESTS_PATH):
            return pd.DataFrame()
        xls = pd.ExcelFile(REQUESTS_PATH)
        sheet = REQUESTS_SHEET if REQUESTS_SHEET in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(REQUESTS_PATH, sheet_name=sheet)
        return normalize_columns(df)
    except Exception as e:
        print("load_requests error:", e)
        return pd.DataFrame()

def save_requests(df):
    df = normalize_columns(df)
    required_cols = [
        'رقم الطلب', 'التاريخ', 'العنوان', 'الوصف',
        'القسم المرسل', 'القسم المستلم', 'الحالة',
        'الموظف المعين', 'آخر تحديث بواسطة', 'الوقت', 'الملف'
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""
    df.to_excel(REQUESTS_PATH, index=False, sheet_name=REQUESTS_SHEET)


def generate_request_id():
    df = load_requests()
    if df.empty or 'رقم الطلب' not in df.columns or df['رقم الطلب'].dropna().empty:
        return f"REQ-{datetime.now().year}-001"
    try:
        last_id = str(df['رقم الطلب'].dropna().iloc[-1])
        number = int(last_id.split('-')[-1]) + 1
        return f"REQ-{datetime.now().year}-{number:03}"
    except:
        return f"REQ-{datetime.now().year}-001"

# ============== الصفحات ==============
@app.route('/')
def index(): return render_template('Login.html')

@app.route('/Login.html')
def login_page(): return render_template('Login.html')

@app.route('/EmployeePage.html')
def emp_page(): return render_template('EmployeePage.html')

@app.route('/DepartmentManagerPage.html')
def mgr_page(): return render_template('DepartmentManagerPage.html')

@app.route('/GeneralManager.html')
def gm_page(): return render_template('GeneralManager.html')
@app.route('/HrPage.html')
def hr_page():
    return render_template('HrPage.html')
@app.route('/ForgotYourPassword.html')
def forgot_page(): return render_template('ForgotYourPassword.html')

# ============== API: الدخول ==============
@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json() or {}
    email = (data.get('email', '') or '').strip().lower()
    password = (data.get('password', '') or '').strip()

    df = load_users()
    if df.empty:
        return jsonify({"success": False, "message": "قاعدة المستخدمين فارغة"}), 500

    # ✅ البحث عن عمود البريد الإلكتروني حتى لو مكتوب بصيغة مختلفة
    email_col = next((c for c in df.columns if 'بريد' in str(c) or 'email' in str(c) or 'ايميل' in str(c)), None)
    pass_col  = next((c for c in df.columns if 'مرور' in str(c) or 'password' in str(c)), None)
    role_col  = next((c for c in df.columns if 'صلاح' in str(c) or 'وظيف' in str(c) or 'role' in str(c)), None)

    if not email_col or not pass_col:
        return jsonify({"success": False, "message": "أعمدة البريد أو كلمة المرور غير موجودة في قاعدة البيانات"}), 500

    # 🔹 تنظيف النصوص داخل الأعمدة
    df[email_col] = df[email_col].astype(str).str.lower().str.strip()
    df[pass_col]  = df[pass_col].astype(str).str.strip()

    # 🔹 دالة مقارنة ذكية تتجاهل المسافات والاختلافات الطفيفة
    def normalize_text(t):
        return re.sub(r'\s+', '', str(t).strip().lower())

    # ✅ البحث الذكي عن المستخدم
    match = df[df.apply(
        lambda row: normalize_text(row[email_col]) == normalize_text(email)
        and normalize_text(row[pass_col]) == normalize_text(password),
        axis=1
    )]

    if match.empty:
        return jsonify({"success": False, "message": "البريد أو كلمة المرور غير صحيحة"}), 401

    user = match.iloc[0].to_dict()

    # ✅ معالجة الصلاحية
    role = str(user.get(role_col or 'الصلاحية', '')).strip()
    role = role.replace('\u200f', '').replace('\u200e', '')

    # ✅ تحسين التعرف على الصلاحيات (يدعم كل الصيغ الممكنة)
    role_norm = re.sub(r'[إأآا]', 'ا', role).replace(' ', '').replace('ة', 'ه').lower()

    if any(k in role_norm for k in ['المواردالبشريه', 'الموارد', 'بشر']):
        role = 'إدارة الموارد البشرية'

    elif any(k in role_norm for k in ['مديرقسم', 'مديرالقسم', 'رئيسقسم', 'رئيسالقسم', 'مديرالاقسام']):
        role = 'مدير قسم'

    elif any(k in role_norm for k in ['مديرعام', 'الادارهالعامه', 'الاداره', 'ادارهعامه']):
        role = 'مدير عام'

    elif any(k in role_norm for k in ['موظف', 'عامل', 'عضو']):
        role = 'موظف'

    else:
        print(f"⚠️ صلاحية غير معروفة: {role_norm}")
        role = 'موظف'  # افتراضي آمن

    # ✅ تحديد الاسم والقسم حتى لو كان بأسماء مختلفة
    name_col = next((c for c in df.columns if 'اسم' in str(c)), 'الاسم')
    dept_col = next((c for c in df.columns if 'قسم' in str(c)), 'القسم')

    name_value = str(user.get(name_col, '')).strip()
    dept_value = str(user.get(dept_col, '')).strip()

    # 🧠 في حال الاسم فاضي، نستخرج الاسم من البريد
    if not name_value:
        name_value = email.split('@')[0] if '@' in email else email

    return jsonify({
        "success": True,
        "user": {
            "email": str(user.get(email_col, '')).strip(),
            "name": name_value,
            "role": role,
            "department": dept_value
        }
    })


# ============== API: جلب الموظفين لكل قسم ==============
@app.route('/api/get_employees', methods=['POST'])
def get_employees():
    """
    جلب الموظفين بناءً على المدير (كل الموظفين التابعين له بغض النظر عن القسم)
    """
    try:
        data = request.get_json() or {}
        manager_name = (data.get('manager_name', '') or '').strip()
        dept = (data.get('department', '') or '').strip()

        df = load_users()
        if df.empty:
            return jsonify({"success": False, "message": "لا توجد بيانات مستخدمين"})

        # 🔹 اكتشاف الأعمدة الأساسية
        name_col = next((c for c in df.columns if 'اسم' in str(c)), 'الاسم')
        role_col = next((c for c in df.columns if 'صلاح' in str(c)), 'الصلاحية')
        dept_col = next((c for c in df.columns if 'قسم' in str(c)), 'القسم')

        df['الاسم'] = df[name_col].astype(str).str.strip()
        df['الصلاحية'] = df[role_col].astype(str).str.strip()
        df['القسم'] = df[dept_col].astype(str).str.strip()

        # ✅ المنطق الجديد:
        # إذا المستخدم مدير قسم → يشوف كل الموظفين اللي صلاحيتهم "موظف"
        if manager_name:
            df = df[df['الصلاحية'].isin(['موظف', 'عامل'])]

        # ✅ المدير العام يشوف الكل
        employees = df[['الاسم', 'القسم', 'الصلاحية']].dropna().to_dict(orient='records')
        return jsonify({"success": True, "employees": employees})

    except Exception as e:
        print("❌ get_employees error:", e)
        return jsonify({"success": False, "message": str(e)})


# ============== API: الطلبات ==============
@app.route('/api/get_requests', methods=['POST'])
def get_requests():
    try:
        data = request.get_json() or {}
        role = data.get('role', '')
        dept = data.get('department', '')
        df = load_requests()

        if df.empty:
            return jsonify([])

        df = normalize_columns(df)
        df['القسم المرسل'] = df['القسم المرسل'].astype(str).str.strip()
        df['القسم المستلم'] = df['القسم المستلم'].astype(str).str.strip()
        df['الحالة'] = df['الحالة'].astype(str).str.strip()

        # ✅ فلترة مطابقة للنسخة القديمة:
        dept_std = normalize_arabic(dept)

        if role == 'موظف':
            filtered = df[
                df['القسم المرسل'].apply(lambda x: dept_std in normalize_arabic(x) or normalize_arabic(x) in dept_std)
                | df['القسم المستلم'].apply(lambda x: dept_std in normalize_arabic(x) or normalize_arabic(x) in dept_std)
            ]
        elif role == 'مدير قسم':
            filtered = df[
                df['القسم المرسل'].apply(lambda x: dept_std in normalize_arabic(x) or normalize_arabic(x) in dept_std)
                | df['القسم المستلم'].apply(lambda x: dept_std in normalize_arabic(x) or normalize_arabic(x) in dept_std)
            ]
        elif role == 'مدير عام':
            filtered = df.copy()
        else:
            filtered = pd.DataFrame()

        # 🔹 إخفاء الحالات المغلقة أو المرفوضة فقط من عرض الصفحة
        filtered = filtered

        return jsonify(filtered.fillna('').to_dict(orient='records'))

    except Exception as e:
        print("get_requests error:", e)
        return jsonify([])

@app.route('/api/create_request', methods=['POST'])
def create_request():
    try:
        title  = request.form.get('title', '').strip()
        desc   = request.form.get('description', '').strip()
        target = request.form.get('targetDept', '').strip()
        sender = request.form.get('senderDept', '').strip()
        sender_name = request.form.get('senderName', '').strip()

        file = request.files.get('file')

        if not all([title, desc, target, sender]):
            return jsonify({"success": False, "message": "جميع الحقول مطلوبة"}), 400

        df = load_requests()
        for col in ['رقم الطلب','التاريخ','العنوان','الوصف','القسم المرسل','القسم المستلم',
                    'الحالة','الموظف المعين','آخر تحديث بواسطة','الوقت','بدأ التنفيذ بواسطة','أغلق بواسطة','الملف']:
            if col not in df.columns:
                df[col] = ""

        req_id = generate_request_id()
        file_name = ""
        if file:
            safe_name = f"{req_id}_{file.filename}"
            file_path = os.path.join(UPLOAD_DIR, safe_name)
            file.save(file_path)
            file_name = safe_name

        new_row = {
            'رقم الطلب': req_id,
            'التاريخ': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'العنوان': title,
            'الوصف': desc,
            'القسم المرسل': sender,
            'اسم المرسل': sender_name,
            'القسم المستلم': target,
            'اسم المستلم': '',
            'الحالة': 'جديد',
            'الموظف المعين': '-',
            'آخر تحديث بواسطة': sender_name or '-',
            'بدأ التنفيذ بواسطة': '',
            'أغلق بواسطة': '',
            'الوقت': '',
            'الملف': file_name
        }

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_requests(df)
        sync_sqlite_to_excel()

        return jsonify({"success": True})
    except Exception as e:
        print("❌ create_request error:", e)
        return jsonify({"success": False, "message": str(e)}), 500


@app.route('/uploads/<path:filename>')
def get_uploaded_file(filename):
    # ✅ يعرض الملف مباشرة داخل المتصفح بدل التحميل
    return send_from_directory(UPLOAD_DIR, filename)

@app.route('/api/update_request_status', methods=['POST'])
def update_request_status():
    data = request.get_json()
    req_id = (data.get('requestId','') or '').strip()
    new_status = (data.get('status','') or '').strip()
    updater = (data.get('updater','') or '').strip()
    duration = data.get('duration')

    df = load_requests()
    if df.empty or 'رقم الطلب' not in df.columns:
        return jsonify({"success": False}), 404

    idx_list = df.index[df['رقم الطلب'] == req_id].tolist()
    if not idx_list:
        return jsonify({"success": False}), 404
    idx = idx_list[0]

    # ✅ ضمان أن الأعمدة النصية هي من نوع str لتفادي تحذير pandas
    text_cols = ['اسم المستلم', 'بدأ التنفيذ بواسطة', 'أغلق بواسطة', 'آخر تحديث بواسطة', 'الوقت']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str)

    # 🔹 تحديث الحالة والاسم
    df.at[idx, 'الحالة'] = new_status
    df.at[idx, 'آخر تحديث بواسطة'] = updater

    # 🔹 تعيين اسم المستلم فقط إذا لم يكن موجود سابقًا
    if not df.at[idx, 'اسم المستلم']:
        df.at[idx, 'اسم المستلم'] = updater

    if new_status == 'جاري التنفيذ':
        df.at[idx, 'بدأ التنفيذ بواسطة'] = updater
        df.at[idx, 'وقت البداية'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    elif new_status == 'معلق':
        df.at[idx, 'وقت التوقف المؤقت'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    elif new_status == 'مغلق':
        if 'وقت البداية' in df.columns:
            start_str = df.at[idx, 'وقت البداية']
            if start_str:
                start_time = datetime.strptime(start_str, '%Y-%m-%d %H:%M:%S')
                diff = datetime.now() - start_time
                df.at[idx, 'الوقت'] = str(diff).split('.')[0]
        if duration:
            df.at[idx, 'الوقت'] = duration
        df.at[idx, 'أغلق بواسطة'] = updater

    if new_status == 'معلق':
        # حفظ وقت الإيقاف المؤقت فقط
        if 'وقت البداية' in df.columns:
            start_str = df.at[idx, 'وقت البداية']
            if start_str:
                start_time = datetime.strptime(start_str, '%Y-%m-%d %H:%M:%S')
                diff = datetime.now() - start_time
                df.at[idx, 'الوقت'] = str(diff).split('.')[0]

    save_requests(df)
    return jsonify({"success": True})


@app.route('/api/delegate_request', methods=['POST'])
def delegate_request():
    data = request.get_json() or {}

    # ✅ يدعم كل أنواع المفاتيح (camelCase أو snake_case)
    req_id = data.get('requestId') or data.get('request_id')
    delegate = data.get('delegate') or data.get('delegateName')
    delegated_by = data.get('delegatedBy') or data.get('delegated_by')

    if not req_id or not delegate:
        return jsonify({'success': False, 'message': 'بيانات غير مكتملة (رقم الطلب أو اسم الموظف مفقود)'})

    df = load_requests()
    if df.empty or 'رقم الطلب' not in df.columns:
        return jsonify({'success': False, 'message': 'قاعدة الطلبات فارغة'})

    mask = df['رقم الطلب'] == req_id
    if not mask.any():
        return jsonify({'success': False, 'message': f'الطلب {req_id} غير موجود'})

    # ✅ تحديث الحقول
    df.loc[mask, 'اسم المستلم'] = delegate
    df.loc[mask, 'آخر تحديث بواسطة'] = delegated_by
    df.loc[mask, 'الحالة'] = 'موكل'

    save_requests(df)
    print(f"✅ تم توكيل الطلب {req_id} إلى {delegate} بواسطة {delegated_by}")
    return jsonify({'success': True})



# ============== API: تصدير الطلبات ==============
@app.route('/api/export_requests', methods=['POST'])
def export_requests():
    """
    📦 تصدير الطلبات إلى ملف Excel يحتوي على عدة أوراق:
    ✅ فقط الطلبات التي استلمها القسم (القسم المستلم)
    كل ورقة تمثل حالة من حالات الطلب (جديد، جاري التنفيذ، مغلق، مرفوض، إلخ)
    """
    try:
        data = request.get_json() or {}
        dept = (data.get('department', '') or '').strip()
        start = (data.get('start_date', '') or '').strip()
        end   = (data.get('end_date', '') or '').strip()

        if not os.path.exists(REQUESTS_PATH):
            return jsonify({"success": False, "message": "ملف الطلبات غير موجود."})

        df = pd.read_excel(REQUESTS_PATH)
        if df.empty:
            return jsonify({"success": False, "message": "لا توجد بيانات لتصديرها."})

        # 🧹 تنظيف الأعمدة المهمة
        for col in ['القسم المستلم', 'الحالة', 'التاريخ']:
            if col in df.columns:
                df[col] = (
                    df[col]
                    .astype(str)
                    .str.strip()
                    .str.replace('\u200f', '', regex=True)
                    .str.replace('\u200e', '', regex=True)
                )

        # ✅ فلترة الطلبات التي استلمها القسم فقط
        dept_norm = normalize_arabic(dept)
        df = df[df['القسم المستلم'].apply(lambda x: dept_norm in normalize_arabic(x) or normalize_arabic(x) in dept_norm)]

        # ✅ فلترة حسب التاريخ إن وجد
        if start:
            df = df[pd.to_datetime(df['التاريخ'], errors='coerce') >= pd.to_datetime(start)]
        if end:
            end_dt = pd.to_datetime(end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df = df[pd.to_datetime(df['التاريخ'], errors='coerce') <= end_dt]

        if df.empty:
            return jsonify({"success": False, "message": "لا توجد طلبات استلمها القسم ضمن الشروط المحددة."})

        # 🗂️ تقسيم البيانات حسب الحالة
        grouped = {status: sub_df for status, sub_df in df.groupby('الحالة')}

        # 📘 إنشاء ملف Excel بعدة أوراق (كل ورقة = حالة)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fname = f"طلبات_الواردة_{dept}_{ts}.xlsx".replace(' ', '_')
        fpath = os.path.join(EXPORT_DIR, fname)

        with pd.ExcelWriter(fpath, engine='openpyxl') as writer:
            for status, sub_df in grouped.items():
                clean_status = str(status).replace('/', '-').strip() or 'غير_محدد'
                sub_df.to_excel(writer, index=False, sheet_name=clean_status[:31])

        return jsonify({"success": True, "file": fname})

    except Exception as e:
        print("❌ export_requests error:", e)
        return jsonify({"success": False, "message": f"حدث خطأ أثناء التصدير: {str(e)}"})

@app.route('/download/<path:filename>')
def download(filename):
    return send_from_directory(EXPORT_DIR, filename, as_attachment=True)

# ============== API: الشات العام ==============
@app.route("/chatbot", methods=["POST"])
def chatbot():
    """رد ذكي باستخدام OpenRouter بسرعة أعلى"""
    user_input = request.json.get("message", "").strip()
    if not user_input:
        return jsonify({"reply": "الرسالة فارغة!"})

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": "qwen/qwen-2.5-7b-instruct",
        "messages": [
            {"role": "system", "content": "أنت مساعد ذكي تتحدث العربية وتساعد موظفي نظام SEVENS."},
            {"role": "user", "content": user_input}
        ],
        "temperature": 0.6,
        "max_tokens": 200
    }

    try:
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=15,   # ⏱️ أقصى مهلة للرد 15 ثانية فقط
        )

        if response.status_code == 200:
            data = response.json()
            if "choices" in data and len(data["choices"]) > 0:
                reply = data["choices"][0]["message"]["content"].strip()
                return jsonify({"reply": reply})
            else:
                return jsonify({"reply": "لم يصل رد من نموذج الذكاء الاصطناعي."})
        else:
            print("❌ OpenRouter Error:", response.text)
            return jsonify({"reply": "حدث خطأ في الخادم أثناء معالجة الطلب."})

    except requests.Timeout:
        return jsonify({"reply": "الخادم تأخر في الرد، حاول مرة أخرى لاحقاً."})
    except Exception as e:
        print("❌ chatbot error:", e)
        return jsonify({"reply": "تعذر الاتصال بخدمة الذكاء الاصطناعي."})

# ============== API: دردشة بين الموظفين ==============
CHAT_UPLOAD_DIR = os.path.join(BASE_DIR, "chat_uploads")
os.makedirs(CHAT_UPLOAD_DIR, exist_ok=True)

@app.route('/api/chat_send_file', methods=['POST'])
def chat_send_file():
    req_id = request.form.get('request_id')
    sender = request.form.get('sender')
    dept = request.form.get('department')
    msg = request.form.get('message', '')
    file = request.files.get('file')
    filename = ""

    if file:
        safe_name = f"{req_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}"
        path = os.path.join(CHAT_UPLOAD_DIR, safe_name)
        file.save(path)
        filename = safe_name

    df = load_chats()
    new = pd.DataFrame([{
        'رقم الطلب': req_id,
        'المرسل': sender,
        'القسم': dept,
        'الرسالة': msg,
        'الملف': filename,
        'الوقت': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }])
    df = pd.concat([df, new], ignore_index=True)
    df.to_excel(CHAT_PATH, index=False)

    # ✅ تحديث الطلب بآخر مرسل
    req_df = load_requests()
    mask = req_df['رقم الطلب'] == req_id
    if mask.any():
        req_df.loc[mask, 'آخر تحديث بواسطة'] = sender
        save_requests(req_df)

    return jsonify({"success": True})

@app.route('/api/chat_get/<req_id>', methods=['GET'])
def chat_get(req_id):
    """إرجاع جميع الرسائل الخاصة بطلب محدد"""
    try:
        df = load_chats()
        if df.empty:
            return jsonify([])
        msgs = df[df['رقم الطلب'].astype(str) == str(req_id)].fillna('').to_dict(orient='records')
        return jsonify(msgs)
    except Exception as e:
        print("❌ chat_get error:", e)
        return jsonify([])


@app.route('/chat_uploads/<path:filename>')
def chat_uploads(filename):
    return send_from_directory(CHAT_UPLOAD_DIR, filename)

# ============== API: استعادة / إعادة تعيين كلمة المرور ==============
@app.route('/api/forgot_reset_password', methods=['POST'])
def forgot_reset_password():
    """تحديث كلمة المرور عبر البريد الإلكتروني"""
    try:
        data = request.get_json() or {}
        email = (data.get('email', '') or '').strip().lower()
        new_password = (data.get('newPassword', '') or '').strip()

        if not email or not new_password:
            return jsonify({"success": False, "message": "يرجى إدخال البريد وكلمة المرور الجديدة"}), 400

        df = load_users()
        if df.empty:
            return jsonify({"success": False, "message": "قاعدة بيانات المستخدمين فارغة"}), 500
        if 'البريد الإلكتروني' not in df.columns:
            return jsonify({"success": False, "message": "عمود البريد الإلكتروني غير موجود"}), 500

        # 🔹 توحيد البريد الإلكتروني للمقارنة
        df['البريد الإلكتروني'] = df['البريد الإلكتروني'].astype(str).str.lower().str.strip()

        # 🔍 البحث عن المستخدم
        mask = df['البريد الإلكتروني'] == email
        if not mask.any():
            return jsonify({"success": False, "message": "البريد الإلكتروني غير موجود"}), 404

        # ✏️ تحديث كلمة المرور
        df.loc[mask, 'كلمة المرور'] = new_password
        df.to_excel(DB_PATH, index=False)
        sync_sqlite_to_excel()

        return jsonify({"success": True, "message": "تم تحديث كلمة المرور بنجاح ✅"})

    except Exception as e:
        print("❌ forgot_reset_password error:", e)
        return jsonify({"success": False, "message": "حدث خطأ أثناء تحديث كلمة المرور"})

# ====== ★★★ HR APIs ★★★ ======

def ensure_users_status_col():
    """نضيف عمود 'الحالة' لملف المستخدمين لو غير موجود"""
    try:
        if not os.path.exists(DB_PATH):
            return
        df = pd.read_excel(DB_PATH)
        if 'الحالة' not in df.columns:
            df['الحالة'] = 'نشط'
            df.to_excel(DB_PATH, index=False)
            sync_sqlite_to_excel()

            print("✅ Added 'الحالة' column to users DB")
    except Exception as e:
        print("ensure_users_status_col error:", e)

ensure_users_status_col()

@app.route('/api/hr/list_users', methods=['GET'])
def hr_list_users():
    """إرجاع كل المستخدمين مع الإيميل وكلمة المرور (HR فقط في الواجهة الأمامية)"""
    try:
        df = load_users()
        if df.empty:
            return jsonify([])
        # ضمان الأعمدة المهمة
        for col in ['الاسم','الصلاحية','كلمة المرور','البريد الإلكتروني','القسم','الحالة']:
            if col not in df.columns:
                df[col] = ''
        return jsonify(df.fillna('').to_dict(orient='records'))
    except Exception as e:
        print("hr_list_users error:", e)
        return jsonify([]), 500

@app.route('/api/hr/add_user', methods=['POST'])
def hr_add_user():
    data = request.get_json() or {}
    name  = (data.get('name','') or '').strip()
    role  = (data.get('role','') or '').strip()
    pwd   = (data.get('password','') or '').strip()
    email = (data.get('email','') or '').strip().lower()
    dept  = (data.get('department','') or '').strip()
    status= (data.get('status','نشط') or 'نشط').strip()

    if not all([name, role, pwd, email, dept]):
        return jsonify({"success": False, "message": "الحقول مطلوبة"}), 400

    df = load_users()
    # ضمان الأعمدة
    for col in ['الاسم','الصلاحية','كلمة المرور','البريد الإلكتروني','القسم','الحالة']:
        if col not in df.columns: df[col] = ''

    # منع تكرار البريد
    if not df[df['البريد الإلكتروني'].astype(str).str.lower().str.strip() == email].empty:
        return jsonify({"success": False, "message": "البريد موجود مسبقاً"}), 409

    new_row = {
        'الاسم': name, 'الصلاحية': role, 'كلمة المرور': pwd,
        'البريد الإلكتروني': email, 'القسم': dept, 'الحالة': status
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(DB_PATH, index=False)
    return jsonify({"success": True})


@app.route('/api/hr/update_user', methods=['POST'])
def hr_update_user():
    data = request.get_json() or {}
    email = (data.get('email','') or '').strip().lower()
    if not email:
        return jsonify({"success": False, "message": "البريد مطلوب"}), 400

    fields_map = {
        'name':'الاسم', 'role':'الصلاحية', 'password':'كلمة المرور',
        'department':'القسم', 'status':'الحالة'
    }
    df = load_users()
    for col in ['الاسم','الصلاحية','كلمة المرور','البريد الإلكتروني','القسم','الحالة']:
        if col not in df.columns: df[col] = ''
    mask = df['البريد الإلكتروني'].astype(str).str.lower().str.strip() == email
    if not mask.any():
        return jsonify({"success": False, "message": "المستخدم غير موجود"}), 404

    for k, ar_col in fields_map.items():
        if k in data and data[k] is not None:
            df.loc[mask, ar_col] = str(data[k]).strip()

    df.to_excel(DB_PATH, index=False)
    sync_sqlite_to_excel()
    return jsonify({"success": True})

@app.route('/api/hr/archive_user', methods=['POST'])
def hr_archive_user():
    data = request.get_json() or {}
    email = (data.get('email','') or '').strip().lower()
    if not email:
        return jsonify({"success": False, "message": "البريد مطلوب"}), 400

    df = load_users()
    if 'الحالة' not in df.columns:
        df['الحالة'] = 'نشط'
    mask = df['البريد الإلكتروني'].astype(str).str.lower().str.strip() == email
    if not mask.any():
        return jsonify({"success": False, "message": "المستخدم غير موجود"}), 404

    df.loc[mask, 'الحالة'] = 'مؤرشف'
    df.to_excel(DB_PATH, index=False)
    return jsonify({"success": True})

def sync_excel_to_sqlite():
    """ينسخ محتوى Excel إلى SQLite إذا تم التعديل على Excel"""
    try:
        conn = sqlite3.connect(DB_SQLITE)
        cur = conn.cursor()

        # 🧱 مزامنة المستخدمين
        if os.path.exists(DB_PATH):
            df_users = pd.read_excel(DB_PATH)
            df_users.columns = [c.strip() for c in df_users.columns]
            for _, row in df_users.iterrows():
                cur.execute("""
                    INSERT OR REPLACE INTO users (email, name, role, password, department, status)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    str(row.get('البريد الإلكتروني', '')).strip().lower(),
                    str(row.get('الاسم', '')).strip(),
                    str(row.get('الصلاحية', '')).strip(),
                    str(row.get('كلمة المرور', '')).strip(),
                    str(row.get('القسم', '')).strip(),
                    str(row.get('الحالة', 'نشط')).strip()
                ))

        # 🧾 مزامنة الطلبات
        if os.path.exists(REQUESTS_PATH):
            df_req = pd.read_excel(REQUESTS_PATH)
            df_req.columns = [c.strip() for c in df_req.columns]
            for _, row in df_req.iterrows():
                cur.execute("""
                    INSERT OR REPLACE INTO requests (req_id, date, title, description, sender_dept, receiver_dept, status, assigned_to, updated_by, duration, file_name)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    str(row.get('رقم الطلب', '')).strip(),
                    str(row.get('التاريخ', '')).strip(),
                    str(row.get('العنوان', '')).strip(),
                    str(row.get('الوصف', '')).strip(),
                    str(row.get('القسم المرسل', '')).strip(),
                    str(row.get('القسم المستلم', '')).strip(),
                    str(row.get('الحالة', '')).strip(),
                    str(row.get('الموظف المعين', '')).strip(),
                    str(row.get('آخر تحديث بواسطة', '')).strip(),
                    str(row.get('الوقت', '')).strip(),
                    str(row.get('الملف', '')).strip(),
                ))

        conn.commit()
        conn.close()
        print("🔁 Excel → SQLite sync done successfully ✅")

    except Exception as e:
        print("❌ sync_excel_to_sqlite error:", e)


def sync_sqlite_to_excel():
    """ينسخ محتوى SQLite إلى Excel عند التحديث من داخل النظام"""
    try:
        conn = sqlite3.connect(DB_SQLITE)
        users_df = pd.read_sql_query("SELECT name AS 'الاسم', role AS 'الصلاحية', password AS 'كلمة المرور', email AS 'البريد الإلكتروني', department AS 'القسم', status AS 'الحالة' FROM users", conn)
        req_df = pd.read_sql_query("SELECT req_id AS 'رقم الطلب', date AS 'التاريخ', title AS 'العنوان', description AS 'الوصف', sender_dept AS 'القسم المرسل', receiver_dept AS 'القسم المستلم', status AS 'الحالة', assigned_to AS 'الموظف المعين', updated_by AS 'آخر تحديث بواسطة', duration AS 'الوقت', file_name AS 'الملف' FROM requests", conn)
        conn.close()

        users_df.to_excel(DB_PATH, index=False)
        req_df.to_excel(REQUESTS_PATH, index=False, sheet_name=REQUESTS_SHEET)

        print("✅ SQLite → Excel backup done successfully")
    except Exception as e:
        print("❌ sync_sqlite_to_excel error:", e)

# ✅ مزامنة قواعد البيانات قبل التشغيل
init_sqlite()
sync_excel_to_sqlite()

import threading
import time

def watch_excel_changes(interval=30):
    """يراقب أي تغييرات في ملفات Excel ويعمل مزامنة تلقائية"""
    last_users_time = os.path.getmtime(DB_PATH)
    last_requests_time = os.path.getmtime(REQUESTS_PATH)

    while True:
        time.sleep(interval)
        try:
            # تحقق من آخر وقت تعديل
            new_users_time = os.path.getmtime(DB_PATH)
            new_requests_time = os.path.getmtime(REQUESTS_PATH)

            # إذا تغير أي ملف → أعد المزامنة
            if new_users_time != last_users_time or new_requests_time != last_requests_time:
                print("🔄 Detected Excel file change, syncing to SQLite...")
                sync_excel_to_sqlite()
                last_users_time = new_users_time
                last_requests_time = new_requests_time

        except Exception as e:
            print("⚠️ watch_excel_changes error:", e)

# 🔁 تشغيل المراقبة في خيط منفصل
threading.Thread(target=watch_excel_changes, daemon=True).start()

# ============== التشغيل ==============
if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)


