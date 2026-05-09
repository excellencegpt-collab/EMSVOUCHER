import base64
import csv
import os
import sqlite3
from datetime import date, datetime, timedelta
from io import BytesIO, StringIO

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


APP_NAME = "Fee Voucher"
DB_FILE = os.environ.get("VOUCHER_DB_PATH", "voucher_records.db")
AOE_LOGO_FILE = "academy_of_excellence_logo.png"
EMS_LOGO_FILE = "excellence_model_school_logo.png"

INSTITUTES = {
    "Academy of Excellence": {
        "short": "AOE",
        "address": "Kharadar, Karachi",
        "phone": "02132312333",
        "logo": AOE_LOGO_FILE,
    },
    "Excellence Model School": {
        "short": "EMS",
        "address": "Kharadar, Karachi",
        "phone": "02132312333",
        "logo": EMS_LOGO_FILE,
    },
}

# ── Excel Import Column Mapping ──────────────────────────────────────────────
# Excel mein yeh columns hone chahiye (case-insensitive)
EXCEL_COLUMNS = {
    "student_name":  ["student name", "student", "name"],
    "father_name":   ["father name", "father", "parent name"],
    "class_name":    ["class", "class name", "grade"],
    "roll_no":       ["roll no", "roll number", "roll"],
    "month":         ["month", "fee month"],
    "admission_fee": ["admission fee", "admission"],
    "tuition_fee":   ["tuition fee", "tuition", "monthly fee"],
    "annual_fee":    ["annual fee", "annual"],
    "exam_fee":      ["exam fee", "exam"],
    "late_fee":      ["late fee", "late"],
    "discount":      ["discount"],
    "notes":         ["notes", "note", "remarks"],
    "institute":     ["institute", "school", "institution"],
}


def connect_db():
    return sqlite3.connect(DB_FILE)


def init_database():
    with connect_db() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS vouchers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                voucher_no TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL,
                institute TEXT NOT NULL,
                student_name TEXT NOT NULL,
                father_name TEXT NOT NULL,
                class_name TEXT NOT NULL,
                roll_no TEXT NOT NULL,
                month TEXT NOT NULL,
                issue_date TEXT NOT NULL,
                due_date TEXT NOT NULL,
                admission_fee INTEGER NOT NULL,
                tuition_fee INTEGER NOT NULL,
                annual_fee INTEGER NOT NULL,
                exam_fee INTEGER NOT NULL,
                late_fee INTEGER NOT NULL,
                discount INTEGER NOT NULL,
                total_amount INTEGER NOT NULL,
                notes TEXT NOT NULL
            )
            """
        )


def next_voucher_no(institute):
    prefix = INSTITUTES[institute]["short"]
    return f"{prefix}-{datetime.now().strftime('%Y%m%d%H%M%S')}"


def money(value):
    return f"Rs. {int(value):,}"


def total_fee(values):
    gross = (
        values["admission_fee"]
        + values["tuition_fee"]
        + values["annual_fee"]
        + values["exam_fee"]
        + values["late_fee"]
    )
    return max(0, gross - values["discount"])


def save_voucher(values):
    with connect_db() as conn:
        conn.execute(
            """
            INSERT INTO vouchers (
                voucher_no, created_at, institute, student_name, father_name,
                class_name, roll_no, month, issue_date, due_date, admission_fee,
                tuition_fee, annual_fee, exam_fee, late_fee, discount,
                total_amount, notes
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                values["voucher_no"],
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                values["institute"],
                values["student_name"],
                values["father_name"],
                values["class_name"],
                values["roll_no"],
                values["month"],
                values["issue_date"],
                values["due_date"],
                values["admission_fee"],
                values["tuition_fee"],
                values["annual_fee"],
                values["exam_fee"],
                values["late_fee"],
                values["discount"],
                values["total_amount"],
                values["notes"],
            ),
        )


def get_vouchers(query=""):
    sql = """
        SELECT voucher_no, created_at, institute, student_name, father_name,
               class_name, roll_no, month, due_date, total_amount
        FROM vouchers
    """
    params = []
    if query:
        sql += """
            WHERE voucher_no LIKE ? OR student_name LIKE ? OR father_name LIKE ?
               OR class_name LIKE ? OR roll_no LIKE ? OR month LIKE ?
        """
        like = f"%{query}%"
        params = [like, like, like, like, like, like]
    sql += " ORDER BY created_at DESC"
    with connect_db() as conn:
        return conn.execute(sql, params).fetchall()


def logo_flowable(institute_name, width=22 * mm, height=22 * mm):
    logo_path = INSTITUTES[institute_name]["logo"]
    if os.path.exists(logo_path):
        return Image(logo_path, width=width, height=height)
    return Paragraph("LOGO", ParagraphStyle("LogoBox", alignment=1, fontSize=10, leading=12))


def fee_rows(values):
    rows = [
        ["Admission Fee", money(values["admission_fee"])],
        ["Tuition Fee", money(values["tuition_fee"])],
        ["Annual Fee", money(values["annual_fee"])],
        ["Exam Fee", money(values["exam_fee"])],
        ["Late Fee", money(values["late_fee"])],
        ["Discount", f"- {money(values['discount'])}"],
    ]
    return [row for row in rows if row[1] != "Rs. 0" and row[1] != "- Rs. 0"]


def voucher_copy(values, copy_label):
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "VoucherTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=17,
        alignment=1,
        textColor=colors.black,
    )
    sub = ParagraphStyle("VoucherSub", parent=styles["Normal"], fontSize=7, leading=9, alignment=1)
    small = ParagraphStyle("Small", parent=styles["Normal"], fontSize=8, leading=10)

    institute = INSTITUTES[values["institute"]]
    header = Table(
        [
            [
                logo_flowable(values["institute"]),
                [
                    Paragraph(values["institute"].upper(), title),
                    Paragraph(f"{institute['address']} | {institute['phone']}", sub),
                    Paragraph(copy_label, sub),
                ],
            ]
        ],
        colWidths=[24 * mm, 142 * mm],
    )
    header.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    student_info = Table(
        [
            ["Voucher No", values["voucher_no"], "Issue Date", values["issue_date"]],
            ["Student Name", values["student_name"], "Due Date", values["due_date"]],
            ["Father Name", values["father_name"], "Month", values["month"]],
            ["Class", values["class_name"], "Roll No", values["roll_no"]],
        ],
        colWidths=[28 * mm, 58 * mm, 28 * mm, 52 * mm],
    )
    student_info.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.35, colors.black),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )

    fees = [["Particular", "Amount"]] + fee_rows(values) + [["Total Amount", money(values["total_amount"])]]
    fee_table = Table(fees, colWidths=[112 * mm, 54 * mm])
    fee_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.35, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#111111")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    footer = Table(
        [
            [Paragraph(f"Note: {values['notes'] or 'Fee once paid is non-refundable.'}", small), "Signature: ____________"]
        ],
        colWidths=[116 * mm, 50 * mm],
    )
    footer.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.35, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )

    return [header, Spacer(1, 4 * mm), student_info, Spacer(1, 4 * mm), fee_table, Spacer(1, 4 * mm), footer]


def create_voucher_pdf(values):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=14 * mm,
        leftMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="Fee Voucher",
    )
    story = []
    story.extend(voucher_copy(values, "PARENT COPY"))
    story.append(Spacer(1, 8 * mm))
    story.extend(voucher_copy(values, "ACADEMY COPY"))
    doc.build(story)
    return buffer.getvalue()


def create_bulk_pdf(all_values):
    """Multiple students ke liye ek PDF mein sab vouchers"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=14 * mm,
        leftMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="Bulk Fee Vouchers",
    )
    story = []
    for i, values in enumerate(all_values):
        story.extend(voucher_copy(values, "PARENT COPY"))
        story.append(Spacer(1, 8 * mm))
        story.extend(voucher_copy(values, "ACADEMY COPY"))
        if i < len(all_values) - 1:
            from reportlab.platypus import PageBreak
            story.append(PageBreak())
    doc.build(story)
    return buffer.getvalue()


def image_to_data_url(path):
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as file:
        encoded = base64.b64encode(file.read()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


# ── Excel Import Helper Functions ────────────────────────────────────────────

def find_column(df_columns, field_key):
    """Excel column ko map karo field name se (case-insensitive)"""
    aliases = EXCEL_COLUMNS.get(field_key, [field_key])
    df_cols_lower = {c.lower().strip(): c for c in df_columns}
    for alias in aliases:
        if alias.lower() in df_cols_lower:
            return df_cols_lower[alias.lower()]
    return None


def parse_excel_import(uploaded_file, default_institute, default_issue_date, default_due_date, default_month):
    """Excel/CSV file parse karke list of student dicts return karo"""
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        return None, f"File read error: {e}"

    df.columns = [str(c).strip() for c in df.columns]
    students = []
    errors = []

    for idx, row in df.iterrows():
        row_num = idx + 2  # header + 1-indexed

        def get_val(field, default=""):
            col = find_column(df.columns, field)
            if col and pd.notna(row.get(col, None)):
                return str(row[col]).strip()
            return default

        def get_int(field, default=0):
            col = find_column(df.columns, field)
            if col and pd.notna(row.get(col, None)):
                try:
                    return int(float(str(row[col]).replace(",", "").replace("Rs.", "").strip()))
                except:
                    return default
            return default

        student_name = get_val("student_name")
        if not student_name:
            errors.append(f"Row {row_num}: Student Name missing — skip kiya gaya")
            continue

        # Institute validate karo
        inst_val = get_val("institute", default_institute)
        if inst_val not in INSTITUTES:
            inst_val = default_institute

        rec = {
            "institute": inst_val,
            "voucher_no": next_voucher_no(inst_val),
            "student_name": student_name,
            "father_name": get_val("father_name", "N/A"),
            "class_name": get_val("class_name", "-"),
            "roll_no": get_val("roll_no", "-"),
            "month": get_val("month", default_month),
            "issue_date": default_issue_date,
            "due_date": default_due_date,
            "admission_fee": get_int("admission_fee"),
            "tuition_fee": get_int("tuition_fee"),
            "annual_fee": get_int("annual_fee"),
            "exam_fee": get_int("exam_fee"),
            "late_fee": get_int("late_fee"),
            "discount": get_int("discount"),
            "notes": get_val("notes", "Late fee will be charged after due date."),
        }
        rec["total_amount"] = total_fee(rec)
        students.append(rec)

    return students, errors


def generate_sample_excel():
    """Sample Excel template download ke liye"""
    sample_data = {
        "Student Name": ["Ali Hassan", "Sara Ahmed"],
        "Father Name": ["Hassan Ali", "Ahmed Khan"],
        "Class": ["8-A", "7-B"],
        "Roll No": ["101", "102"],
        "Month": ["May 2026", "May 2026"],
        "Tuition Fee": [3000, 3000],
        "Admission Fee": [0, 500],
        "Annual Fee": [0, 1000],
        "Exam Fee": [500, 500],
        "Late Fee": [0, 0],
        "Discount": [0, 200],
        "Notes": ["", "Scholarship student"],
        "Institute": ["Academy of Excellence", "Excellence Model School"],
    }
    df = pd.DataFrame(sample_data)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Students")
    return buffer.getvalue()


def export_history_excel(records):
    """Voucher history ko Excel mein export karo"""
    rows = []
    for r in records:
        rows.append({
            "Voucher No": r[0],
            "Created At": r[1],
            "Institute": r[2],
            "Student Name": r[3],
            "Father Name": r[4],
            "Class": r[5],
            "Roll No": r[6],
            "Month": r[7],
            "Due Date": r[8],
            "Total Amount": r[9],
        })
    df = pd.DataFrame(rows)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Voucher History")
    return buffer.getvalue()


# ── Streamlit UI ─────────────────────────────────────────────────────────────

st.set_page_config(page_title=APP_NAME, layout="wide")
init_database()
if "selected_institute" not in st.session_state:
    st.session_state.selected_institute = "Academy of Excellence"
if "imported_students" not in st.session_state:
    st.session_state.imported_students = []

st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(180deg, #101010 0%, #191919 100%);
        color: #f8fafc;
    }
    .block-container {padding-top: 2rem; padding-bottom: 2rem; max-width: 1250px;}
    h1, h2, h3, h4, h5, h6, p, label, span {color: #f8fafc;}
    .hero {
        background: linear-gradient(120deg, #111111, #252525);
        border: 1px solid #333;
        border-radius: 14px;
        padding: 24px;
        display: flex;
        align-items: center;
        gap: 18px;
        margin-bottom: 18px;
        box-shadow: 0 16px 34px rgba(0,0,0,0.32);
    }
    .hero-logo {
        width: 82px;
        height: 82px;
        object-fit: contain;
        background: white;
        border-radius: 10px;
        padding: 8px;
    }
    .hero-title {font-size: 34px; font-weight: 900; line-height: 1.05;}
    .hero-sub {color: #f3ca66; font-weight: 700; margin-top: 6px;}
    .panel {
        background: #202020;
        border: 1px solid #343434;
        border-radius: 12px;
        padding: 18px;
    }
    .import-box {
        background: #1a1a2e;
        border: 1px dashed #f0b429;
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 12px;
    }
    .stButton > button, .stDownloadButton > button {
        background: #f0b429 !important;
        color: #111111 !important;
        border: 0 !important;
        border-radius: 999px !important;
        font-weight: 900 !important;
    }
    .stTextInput input, .stNumberInput input, .stTextArea textarea, .stDateInput input, .stSelectbox div[data-baseweb="select"] {
        background: #141414 !important;
        color: #ffffff !important;
        border-color: #3c3c3c !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

active_logo = INSTITUTES[st.session_state.selected_institute]["logo"]
logo_url = image_to_data_url(active_logo)
logo_html = f'<img class="hero-logo" src="{logo_url}" alt="Logo">' if logo_url else '<div class="hero-logo"></div>'
st.markdown(
    f"""
    <div class="hero">
        {logo_html}
        <div>
            <div class="hero-title">Fee Voucher</div>
            <div class="hero-sub">Excellence Online</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

voucher_tab, import_tab, history_tab, setup_tab = st.tabs([
    "➕ Create Voucher",
    "📥 Import from Excel",
    "📋 Voucher History",
    "🖼️ Logo Setup",
])

# ── TAB 1: Create Single Voucher ─────────────────────────────────────────────
with voucher_tab:
    left, right = st.columns([1, 1], gap="large")
    with left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.subheader("Student Details")
        institute = st.selectbox("Institute", list(INSTITUTES.keys()), key="selected_institute")
        voucher_no = st.text_input("Voucher No", value=next_voucher_no(institute))
        student_name = st.text_input("Student Name")
        father_name = st.text_input("Father Name")
        class_cols = st.columns(2)
        class_name = class_cols[0].text_input("Class")
        roll_no = class_cols[1].text_input("Roll No")
        month = st.text_input("Fee Month", value=datetime.now().strftime("%B %Y"))
        date_cols = st.columns(2)
        issue_date = date_cols[0].date_input("Issue Date", value=date.today())
        due_date = date_cols[1].date_input("Due Date", value=date.today() + timedelta(days=10))
        notes = st.text_area("Note", value="Late fee will be charged after due date.")
        st.markdown('</div>', unsafe_allow_html=True)

    with right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.subheader("Fee Details")
        fee_cols = st.columns(2)
        admission_fee = fee_cols[0].number_input("Admission Fee", min_value=0, value=0, step=100)
        tuition_fee = fee_cols[1].number_input("Tuition Fee", min_value=0, value=0, step=100)
        annual_fee = fee_cols[0].number_input("Annual Fee", min_value=0, value=0, step=100)
        exam_fee = fee_cols[1].number_input("Exam Fee", min_value=0, value=0, step=100)
        late_fee = fee_cols[0].number_input("Late Fee", min_value=0, value=0, step=50)
        discount = fee_cols[1].number_input("Discount", min_value=0, value=0, step=50)

        values = {
            "voucher_no": voucher_no.strip() or next_voucher_no(institute),
            "institute": institute,
            "student_name": student_name.strip() or "Student Name",
            "father_name": father_name.strip() or "Father Name",
            "class_name": class_name.strip() or "-",
            "roll_no": roll_no.strip() or "-",
            "month": month.strip() or datetime.now().strftime("%B %Y"),
            "issue_date": issue_date.strftime("%d-%m-%Y"),
            "due_date": due_date.strftime("%d-%m-%Y"),
            "admission_fee": int(admission_fee),
            "tuition_fee": int(tuition_fee),
            "annual_fee": int(annual_fee),
            "exam_fee": int(exam_fee),
            "late_fee": int(late_fee),
            "discount": int(discount),
            "notes": notes.strip(),
        }
        values["total_amount"] = total_fee(values)
        st.metric("Total Amount", money(values["total_amount"]))

        pdf_bytes = create_voucher_pdf(values)
        save_cols = st.columns(2)
        if save_cols[0].button("Save Record", use_container_width=True):
            try:
                save_voucher(values)
                st.success("Voucher record saved.")
            except sqlite3.IntegrityError:
                st.error("Ye voucher number already saved hai. Voucher No change karein.")
        save_cols[1].download_button(
            "Download PDF",
            data=pdf_bytes,
            file_name=f"{values['voucher_no']}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)

# ── TAB 2: Import from Excel ─────────────────────────────────────────────────
with import_tab:
    st.subheader("📥 Excel / CSV se Bulk Import")

    # Sample template download
    st.markdown("#### Step 1: Sample Template Download Karen")
    st.info(
        "Pehle sample Excel template download karein, usme data fill karein, phir upload karein. "
        "Columns ka naam change mat karein."
    )
    st.download_button(
        "⬇️ Sample Excel Template Download",
        data=generate_sample_excel(),
        file_name="fee_voucher_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    # Upload section
    st.markdown("#### Step 2: Apni Excel File Upload Karen")

    imp_cols = st.columns([1, 1])
    with imp_cols[0]:
        default_institute_imp = st.selectbox(
            "Default Institute (agar Excel mein column na ho)",
            list(INSTITUTES.keys()),
            key="imp_institute",
        )
    with imp_cols[1]:
        default_month_imp = st.text_input(
            "Default Fee Month",
            value=datetime.now().strftime("%B %Y"),
            key="imp_month",
        )

    date_imp_cols = st.columns(2)
    default_issue = date_imp_cols[0].date_input("Issue Date (sab ke liye)", value=date.today(), key="imp_issue")
    default_due = date_imp_cols[1].date_input("Due Date (sab ke liye)", value=date.today() + timedelta(days=10), key="imp_due")

    uploaded_file = st.file_uploader(
        "Excel (.xlsx) ya CSV (.csv) file select karein",
        type=["xlsx", "xls", "csv"],
        key="excel_import",
    )

    if uploaded_file:
        students, errors = parse_excel_import(
            uploaded_file,
            default_institute=default_institute_imp,
            default_issue_date=default_issue.strftime("%d-%m-%Y"),
            default_due_date=default_due.strftime("%d-%m-%Y"),
            default_month=default_month_imp,
        )

        if errors:
            for err in errors:
                st.warning(err)

        if students:
            st.success(f"✅ {len(students)} students ka data successfully parse hua!")
            st.session_state.imported_students = students

            # Preview table
            st.markdown("#### Preview — Data Check Karen")
            preview_rows = []
            for s in students:
                preview_rows.append({
                    "Student": s["student_name"],
                    "Father": s["father_name"],
                    "Class": s["class_name"],
                    "Roll No": s["roll_no"],
                    "Month": s["month"],
                    "Institute": s["institute"],
                    "Tuition": s["tuition_fee"],
                    "Total": s["total_amount"],
                })
            st.dataframe(preview_rows, use_container_width=True, hide_index=True)

            st.divider()
            st.markdown("#### Step 3: Bulk Actions")

            action_cols = st.columns(3)

            # Bulk save to DB
            if action_cols[0].button("💾 Sab Records Save Karein", use_container_width=True):
                saved = 0
                failed = 0
                for s in students:
                    try:
                        save_voucher(s)
                        saved += 1
                    except sqlite3.IntegrityError:
                        failed += 1
                if saved:
                    st.success(f"{saved} vouchers database mein save ho gaye!")
                if failed:
                    st.warning(f"{failed} vouchers already exist the, skip kiye gaye.")

            # Bulk PDF download
            if action_cols[1].button("📄 Sab PDF Generate Karein", use_container_width=True):
                with st.spinner("PDF bana raha hoon..."):
                    bulk_pdf = create_bulk_pdf(students)
                st.download_button(
                    "⬇️ Bulk PDF Download",
                    data=bulk_pdf,
                    file_name=f"bulk_vouchers_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

            # Export back to Excel with totals
            export_rows = []
            for s in students:
                export_rows.append({
                    "Voucher No": s["voucher_no"],
                    "Institute": s["institute"],
                    "Student Name": s["student_name"],
                    "Father Name": s["father_name"],
                    "Class": s["class_name"],
                    "Roll No": s["roll_no"],
                    "Month": s["month"],
                    "Issue Date": s["issue_date"],
                    "Due Date": s["due_date"],
                    "Admission Fee": s["admission_fee"],
                    "Tuition Fee": s["tuition_fee"],
                    "Annual Fee": s["annual_fee"],
                    "Exam Fee": s["exam_fee"],
                    "Late Fee": s["late_fee"],
                    "Discount": s["discount"],
                    "Total Amount": s["total_amount"],
                    "Notes": s["notes"],
                })
            df_export = pd.DataFrame(export_rows)
            exp_buffer = BytesIO()
            with pd.ExcelWriter(exp_buffer, engine="xlsxwriter") as writer:
                df_export.to_excel(writer, index=False, sheet_name="Processed Vouchers")
            action_cols[2].download_button(
                "📊 Processed Excel Export",
                data=exp_buffer.getvalue(),
                file_name=f"processed_vouchers_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            if not errors:
                st.error("File mein koi valid data nahi mila.")

# ── TAB 3: History ────────────────────────────────────────────────────────────
with history_tab:
    st.subheader("Voucher History")
    h_cols = st.columns([3, 1])
    query = h_cols[0].text_input("Search voucher, student, father, class, roll no, or month")
    records = get_vouchers(query.strip())

    if records:
        rows = [
            {
                "Voucher No": row[0],
                "Created At": row[1],
                "Institute": row[2],
                "Student": row[3],
                "Father": row[4],
                "Class": row[5],
                "Roll No": row[6],
                "Month": row[7],
                "Due Date": row[8],
                "Total": money(row[9]),
            }
            for row in records
        ]
        st.dataframe(rows, use_container_width=True, hide_index=True)

        exp_cols = st.columns(2)

        # CSV export
        output = StringIO()
        writer = csv.writer(output)
        writer.writerow(["Voucher No", "Created At", "Institute", "Student", "Father", "Class", "Roll No", "Month", "Due Date", "Total"])
        writer.writerows(records)
        exp_cols[0].download_button(
            "⬇️ CSV Export",
            data=output.getvalue(),
            file_name=f"voucher_history_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

        # Excel export
        excel_data = export_history_excel(records)
        exp_cols[1].download_button(
            "📊 Excel Export",
            data=excel_data,
            file_name=f"voucher_history_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.info("No voucher records found.")

# ── TAB 4: Logo Setup ─────────────────────────────────────────────────────────
with setup_tab:
    st.subheader("Logo Setup")
    st.info("Dropdown mein institute select karne par uska logo automatically PDF voucher mein aayega.")
    logo_cols = st.columns(2)
    for logo_col, (institute_name, info) in zip(logo_cols, INSTITUTES.items()):
        with logo_col:
            st.markdown(f"#### {institute_name}")
            if os.path.exists(info["logo"]):
                st.image(info["logo"], width=180)
            uploaded_logo = st.file_uploader(
                f"Replace {info['short']} Logo",
                type=["png", "jpg", "jpeg"],
                key=f"logo_{info['short']}",
            )
            if uploaded_logo:
                with open(info["logo"], "wb") as file:
                    file.write(uploaded_logo.getbuffer())
                st.success(f"{info['short']} logo saved.")
