import base64
import csv
import os
import sqlite3
from datetime import date, datetime, timedelta
from io import BytesIO, StringIO

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


def image_to_data_url(path):
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as file:
        encoded = base64.b64encode(file.read()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


st.set_page_config(page_title=APP_NAME, layout="wide")
init_database()
if "selected_institute" not in st.session_state:
    st.session_state.selected_institute = "Academy of Excellence"

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

voucher_tab, history_tab, setup_tab = st.tabs(["Create Voucher", "Voucher History", "Logo Setup"])

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

with history_tab:
    st.subheader("Voucher History")
    query = st.text_input("Search voucher, student, father, class, roll no, or month")
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

        output = StringIO()
        writer = csv.writer(output)
        writer.writerow(["Voucher No", "Created At", "Institute", "Student", "Father", "Class", "Roll No", "Month", "Due Date", "Total"])
        writer.writerows(records)
        st.download_button(
            "Download CSV",
            data=output.getvalue(),
            file_name=f"voucher_history_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )
    else:
        st.info("No voucher records found.")
