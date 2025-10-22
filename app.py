from flask import Flask, request, jsonify, render_template, send_from_directory
from flask_cors import CORS
import pandas as pd
import os
from datetime import datetime
import requests
import re

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
                    'القسم المستلم', 'الحالة', 'الموظف المعين', 'آخر تحديث بواسطة', 'الوقت']
        pd.DataFrame(columns=req_cols).to_excel(REQUESTS_PATH, index=False, sheet_name=REQUESTS_SHEET)
        print("✅ Created requests DB")

def normalize_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df

def load_users():
    try:
        df = pd.read_excel(DB_PATH)
        df.columns = df.columns.str.strip().str.replace('\u200f','', regex=True).str.replace('\u200e','', regex=True)
        rename_map = {'البريد الالكتروني':'البريد الإلكتروني','البريد الالكترونى':'البريد الإلكتروني','الايميل':'البريد الإلكتروني'}
        df.rename(columns=rename_map, inplace=True)
        df = normalize_department_names(df)
        return normalize_columns(df)
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
    for col in ['رقم الطلب','التاريخ','العنوان','الوصف','القسم المرسل','القسم المستلم',
                'الحالة','الموظف المعين','آخر تحديث بواسطة','الوقت']:
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

@app.route('/ForgotYourPassword.html')
def forgot_page(): return render_template('ForgotYourPassword.html')

# ============== API: الدخول ==============
@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json()
    email = (data.get('email','') or '').strip().lower()
    password = (data.get('password','') or '').strip()

    df = load_users()
    if df.empty:
        return jsonify({"success": False, "message": "قاعدة المستخدمين فارغة"}), 500
    if 'البريد الإلكتروني' not in df.columns:
        return jsonify({"success": False, "message": "عمود البريد الإلكتروني غير موجود في قاعدة البيانات"}), 500

    df['البريد الإلكتروني'] = df['البريد الإلكتروني'].astype(str).str.lower().str.strip()
    df['كلمة المرور'] = df['كلمة المرور'].astype(str).str.strip()

    match = df[(df['البريد الإلكتروني']==email) & (df['كلمة المرور']==password)]
    if match.empty:
        return jsonify({"success": False, "message": "البريد أو كلمة المرور غير صحيحة"}), 401

    user = match.iloc[0].to_dict()
    role = str(user.get('الصلاحية','')).strip().replace('\u200f','').replace('\u200e','')
    if role in ['مدير القسم','مدير أقسام','رئيس قسم']: role='مدير قسم'
    elif role in ['موظف','موظفه','عامل']:             role='موظف'
    elif role in ['مدير عام','الإدارة العامة']:        role='مدير عام'

    return jsonify({"success": True,"user":{
        "email": str(user.get('البريد الإلكتروني','')).strip(),
        "name": str(user.get('الاسم','')).strip(),
        "role": role,
        "department": str(user.get('القسم','')).strip()
    }})

# ============== API: الطلبات ==============

@app.route('/api/get_requests', methods=['POST'])
def get_requests():
    try:
        data = request.get_json() or {}
        role = (data.get('role', '') or '').strip()
        dept = (data.get('department', '') or '').strip()
        df = load_requests()
        if df.empty:
            return jsonify([])

        # تنظيف النصوص من الفراغات والرموز
        for col in ['القسم المرسل', 'القسم المستلم', 'الحالة']:
            if col in df.columns:
                df[col] = (
                    df[col].astype(str)
                    .str.strip()
                    .str.replace('\u200f', '', regex=True)
                    .str.replace('\u200e', '', regex=True)
                    .str.replace('  ', ' ', regex=True)
                )

        # قاموس تطابق الأقسام (توحيد الأسماء)
        dept_aliases = {
            "ادارة التقنية والشبكات": "إدارة تقنية المعلومات",
            "إدارة تقنية المعلومات": "إدارة تقنية المعلومات",
            "الادارة التقنية": "إدارة تقنية المعلومات",
            "ادارة تقنية المعلومات": "إدارة تقنية المعلومات",
            "ادارة  التقنية والشبكات": "إدارة تقنية المعلومات",

            "الادارة المالية": "الإدارة المالية",
            "إدارة المالية": "الإدارة المالية",

            "ادارة الصيانة وقطع الغيار": "إدارة الصيانة وقطع الغيار",
            "قسم استقبال عملاء الصيانة": "إدارة الصيانة وقطع الغيار",
            "قسم خدمة عملاء الصيانة": "إدارة الصيانة وقطع الغيار",

            "قسم المبيعات الهاتفية": "إدارة المبيعات الهاتفية - المدينة",
            "إدارة المبيعات الهاتفية - المدينة": "إدارة المبيعات الهاتفية - المدينة",
            "قسم المبيعات الهاتفية الرياض": "إدارة المبيعات الهاتفية - الرياض",
            "إدارة المبيعات الهاتفية - الرياض": "إدارة المبيعات الهاتفية - الرياض",

            "مبيعات الصالة - المدينة": "قسم مبيعات الصالة المدينة المنورة",
            "قسم مبيعات الصالة المدينة المنورة": "قسم مبيعات الصالة المدينة المنورة",
            "قسم مبيعات الصالة - جدة": "قسم مبيعات الصالة - جدة",

            "قسم النظافة والاستقبال المدينة": "قسم الاستقبال والنظافة",
            "قسم النظافة والاستقبال المدينه": "قسم الاستقبال والنظافة",
            "قسم النظافة والاستقبال جدة": "قسم النظافة والاستقبال جدة",
            "قسم النظافة والاستقبال الرياض": "قسم النظافة والاستقبال الرياض",
            "قسم النظافة والاستقبال الرياض 1": "قسم النظافة والاستقبال الرياض 1",
            "قسم النظافة والاستقبال الرياض 2": "قسم النظافة والاستقبال الرياض 2",
            "قسم الاستقبال والنظافة": "قسم الاستقبال والنظافة",

            "ادارة التسويق": "إدارة التسويق",
            "إدارة التسويق": "إدارة التسويق",
            "ادارة المبيعات": "إدارة المبيعات",
            "إدارة المبيعات": "إدارة المبيعات",

            "ادارة المراجعة": "إدارة المراجعة",
            "إدارة المراجعة": "إدارة المراجعة",

            "الادارة العامة": "الإدارة العامة",
            "الادارة الادارية": "الإدارة الإدارية",
            "ادارة التشغيل": "إدارة التشغيل",
            "مكتب التأجير": "إدارة مكتب التأجير",
            "ادارة مكتب التأجير": "إدارة مكتب التأجير",
            "قسم مخزون السيارات": "إدارة المخزون - مستودع السيارات",
        }

        def normalize_dept(name):
            n = str(name).strip()
            n = n.replace('\u200f', '').replace('\u200e', '').replace('  ', ' ')
            if n in dept_aliases:
                return dept_aliases[n]
            # fallback normalization (ignore "ادارة"/"إدارة" differences)
            if 'التقنية' in n or 'الشبكات' in n or 'المعلومات' in n:
                return "إدارة تقنية المعلومات"
            if 'الصيانة' in n:
                return "إدارة الصيانة وقطع الغيار"
            if 'المبيعات الهاتفية' in n:
                if 'الرياض' in n:
                    return "إدارة المبيعات الهاتفية - الرياض"
                else:
                    return "إدارة المبيعات الهاتفية - المدينة"
            return n

        dept_std = normalize_dept(dept)
        df['sent_norm'] = df['القسم المرسل'].apply(normalize_dept)
        df['recv_norm'] = df['القسم المستلم'].apply(normalize_dept)

        # الفلترة حسب الدور
        # ✅ مدير القسم يشاهد فقط الطلبات التي تخص قسمه (مرسلة أو مستلمة)
        if role == 'موظف':
            # الموظف يشاهد فقط الطلبات الواردة لقسمه
            filtered = df[df['recv_norm'] == dept_std]


        elif role == 'مدير قسم':

            # مدير القسم يشاهد الطلبات سواء كانت باسمه الأصلي أو بعد التطبيع

            filtered = df[

                (df['recv_norm'] == dept_std) |

                (df['sent_norm'] == dept_std) |

                (df['القسم المستلم'].str.contains(dept, case=False, na=False)) |

                (df['القسم المرسل'].str.contains(dept, case=False, na=False))

                ]


        elif role == 'مدير عام':
            # المدير العام يشاهد كل الطلبات بالنظام
            filtered = df

        else:
            filtered = df.iloc[0:0]

        # استبعاد الطلبات المغلقة والمرفوضة
        if 'الحالة' in filtered.columns:
            filtered = filtered[~filtered['الحالة'].isin(['مغلق', 'مرفوض'])]

        return jsonify(filtered.fillna('').to_dict(orient='records'))
    except Exception as e:
        print("get_requests error:", e)
        return jsonify([])

@app.route('/api/create_request', methods=['POST'])
def create_request():
    try:
        data = request.get_json()
        title  = (data.get('title','') or '').strip()
        desc   = (data.get('description','') or '').strip()
        target = (data.get('targetDept','') or '').strip()
        sender = (data.get('senderDept','') or '').strip()
        sender_name = (data.get('senderName','') or '').strip()

        if not all([title,desc,target,sender]):
            return jsonify({"success": False, "message": "جميع الحقول مطلوبة"}), 400

        df = load_requests()
        for col in ['رقم الطلب','التاريخ','العنوان','الوصف','القسم المرسل','القسم المستلم',
                    'الحالة','الموظف المعين','آخر تحديث بواسطة','الوقت']:
            if col not in df.columns: df[col] = ""

        new_row = {
            'رقم الطلب': generate_request_id(),
            'التاريخ': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'العنوان': title, 'الوصف': desc,
            'القسم المرسل': sender, 'القسم المستلم': target,
            'الحالة': 'جديد', 'الموظف المعين': '-',
            'آخر تحديث بواسطة': sender_name or '-', 'الوقت':''
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_requests(df)
        return jsonify({"success": True})
    except Exception as e:
        print("❌ create_request error:", e)
        return jsonify({"success": False, "message": str(e)}), 500

@app.route('/api/update_request_status', methods=['POST'])
def update_request_status():
    data = request.get_json()
    req_id = (data.get('requestId','') or '').strip()
    new_status = (data.get('status','') or '').strip()
    updater = (data.get('updater','') or '').strip()

    df = load_requests()
    if df.empty or 'رقم الطلب' not in df.columns:
        return jsonify({"success": False}), 404

    idx_list = df.index[df['رقم الطلب'] == req_id].tolist()
    if not idx_list: return jsonify({"success": False}), 404

    idx = idx_list[0]
    if 'الحالة' in df.columns: df.at[idx, 'الحالة'] = new_status
    if 'آخر تحديث بواسطة' in df.columns: df.at[idx, 'آخر تحديث بواسطة'] = updater
    duration = data.get('duration')
    if duration:
        if 'الوقت' not in df.columns: df['الوقت'] = ''
        df.at[idx, 'الوقت'] = duration

    save_requests(df)
    return jsonify({"success": True})

# ============== API: تصدير الطلبات ==============
@app.route('/api/export_requests', methods=['POST'])
def export_requests():
    """
    يسمح لمدير القسم بتصدير الطلبات المغلقة لقسمه فقط
    خلال فترة زمنية محددة، ويتم تحميل الملف على جهازه.
    """
    try:
        data = request.get_json() or {}
        dept = (data.get('department', '') or '').strip()
        start = (data.get('start_date', '') or '').strip()
        end   = (data.get('end_date', '') or '').strip()

        df = pd.read_excel(REQUESTS_PATH)
        if df.empty:
            return jsonify({"success": False, "message": "لا توجد بيانات لتصديرها."})

        # تنظيف الأعمدة
        for col in ['القسم المستلم', 'الحالة', 'التاريخ']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.replace('\u200f','', regex=True).str.replace('\u200e','', regex=True)

        # 🔹 نفس قاموس التطبيع المستخدم في get_requests
        dept_aliases = {
            "ادارة التقنية والشبكات": "إدارة تقنية المعلومات",
            "إدارة تقنية المعلومات": "إدارة تقنية المعلومات",
            "الادارة التقنية": "إدارة تقنية المعلومات",
            "ادارة تقنية المعلومات": "إدارة تقنية المعلومات",
            "ادارة  التقنية والشبكات": "إدارة تقنية المعلومات",

            "الادارة المالية": "الإدارة المالية",
            "إدارة المالية": "الإدارة المالية",

            "ادارة الصيانة وقطع الغيار": "إدارة الصيانة وقطع الغيار",
            "قسم استقبال عملاء الصيانة": "إدارة الصيانة وقطع الغيار",
            "قسم خدمة عملاء الصيانة": "إدارة الصيانة وقطع الغيار",

            "قسم المبيعات الهاتفية": "إدارة المبيعات الهاتفية - المدينة",
            "إدارة المبيعات الهاتفية - المدينة": "إدارة المبيعات الهاتفية - المدينة",
            "قسم المبيعات الهاتفية الرياض": "إدارة المبيعات الهاتفية - الرياض",
            "إدارة المبيعات الهاتفية - الرياض": "إدارة المبيعات الهاتفية - الرياض",

            "مبيعات الصالة - المدينة": "قسم مبيعات الصالة المدينة المنورة",
            "قسم مبيعات الصالة المدينة المنورة": "قسم مبيعات الصالة المدينة المنورة",
            "قسم مبيعات الصالة - جدة": "قسم مبيعات الصالة - جدة",

            "ادارة التسويق": "إدارة التسويق",
            "إدارة التسويق": "إدارة التسويق",
            "ادارة المبيعات": "إدارة المبيعات",
            "إدارة المبيعات": "إدارة المبيعات",

            "الادارة العامة": "الإدارة العامة",
            "الادارة الادارية": "الإدارة الإدارية",
            "ادارة التشغيل": "إدارة التشغيل",
            "مكتب التأجير": "إدارة مكتب التأجير",
            "ادارة مكتب التأجير": "إدارة مكتب التأجير",
            "قسم مخزون السيارات": "إدارة المخزون - مستودع السيارات",
        }

        def normalize_dept(name):
            n = str(name).strip().replace('\u200f','').replace('\u200e','')
            if n in dept_aliases:
                return dept_aliases[n]
            if 'التقنية' in n or 'الشبكات' in n or 'المعلومات' in n:
                return "إدارة تقنية المعلومات"
            if 'الصيانة' in n:
                return "إدارة الصيانة وقطع الغيار"
            if 'المبيعات الهاتفية' in n:
                if 'الرياض' in n:
                    return "إدارة المبيعات الهاتفية - الرياض"
                else:
                    return "إدارة المبيعات الهاتفية - المدينة"
            return n

        dept_std = normalize_dept(dept)
        df['recv_norm'] = df['القسم المستلم'].apply(normalize_dept)
        df['status_norm'] = df['الحالة'].astype(str).str.strip()

        # ✅ فلترة الطلبات المغلقة الخاصة بالقسم
        mask = (df['recv_norm'] == dept_std) & (df['status_norm'] == 'مغلق')
        out = df[mask].copy()

        # ✅ فلترة حسب التاريخ
        if start:
            out = out[pd.to_datetime(out['التاريخ']) >= pd.to_datetime(start)]
        if end:
            end_dt = pd.to_datetime(end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            out = out[pd.to_datetime(out['التاريخ']) <= end_dt]

        if out.empty:
            return jsonify({"success": False, "message": "لا توجد طلبات مغلقة ضمن الشروط المحددة."})

        # إنشاء الملف
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fname = f"طلبات_مغلقة_{dept_std}_{ts}.xlsx".replace(' ', '_')
        fpath = os.path.join(EXPORT_DIR, fname)
        out.to_excel(fpath, index=False)

        return jsonify({"success": True, "file": fname})

    except Exception as e:
        print("❌ export_requests error:", e)
        return jsonify({"success": False, "message": "حدث خطأ أثناء تصدير البيانات."})

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


# ============== التشغيل ==============
if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

