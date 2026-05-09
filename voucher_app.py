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


# =========================
# DATABASE
# =========================

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
                voucher_no, created_at, institute, student_name,
                father_name, class_name, roll_no, month,
                issue_date, due_date, admission_fee,
                tuition_fee, annual_fee, exam_fee,
                late_fee, discount, total_amount, notes
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
        SELECT voucher_no, created_at, institute,
               student_name, father_name,
               class_name, roll_no,
               month, due_date, total_amount
        FROM vouchers
    """

    params = []

    if query:

        sql += """
            WHERE voucher_no LIKE ?
            OR student_name LIKE ?
            OR father_name LIKE ?
            OR class_name LIKE ?
            OR roll_no LIKE ?
            OR month LIKE ?
        """

        like = f"%{query}%"

        params = [like, like, like, like, like, like]

    sql += " ORDER BY created_at DESC"

    with connect_db() as conn:
        return conn.execute(sql, params).fetchall()


# =========================
# PDF
# =========================

def logo_flowable(institute_name, width=22 * mm, height=22 * mm):

    logo_path = INSTITUTES[institute_name]["logo"]

    if os.path.exists(logo_path):
        return Image(logo_path, width=width, height=height)

    return Paragraph(
        "LOGO",
        ParagraphStyle(
            "LogoBox",
            alignment=1,
            fontSize=10,
            leading=12
        )
    )


def fee_rows(values):

    rows = [
        ["Admission Fee", money(values["admission_fee"])],
        ["Tuition Fee", money(values["tuition_fee"])],
        ["Annual Fee", money(values["annual_fee"])],
        ["Exam Fee", money(values["exam_fee"])],
        ["Late Fee", money(values["late_fee"])],
        ["Discount", f"- {money(values['discount'])}"],
    ]

    return [
        row
        for row in rows
        if row[1] != "Rs. 0"
        and row[1] != "- Rs. 0"
    ]


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

    sub = ParagraphStyle(
        "VoucherSub",
        parent=styles["Normal"],
        fontSize=7,
        leading=9,
        alignment=1
    )

    small = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontSize=8,
        leading=10
    )

    institute = INSTITUTES[values["institute"]]

    header = Table(
        [
            [
                logo_flowable(values["institute"]),
                [
                    Paragraph(values["institute"].upper(), title),
                    Paragraph(
                        f"{institute['address']} | {institute['phone']}",
                        sub
                    ),
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
            ]
        )
    )

    fees = (
        [["Particular", "Amount"]]
        + fee_rows(values)
        + [["Total Amount", money(values["total_amount"])]]
    )

    fee_table = Table(
        fees,
        colWidths=[112 * mm, 54 * mm]
    )

    fee_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.35, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#111111")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ]
        )
    )

    footer = Table(
        [
            [
                Paragraph(
                    f"Note: {values['notes']}",
                    small
                ),
                "Signature: ____________"
            ]
        ],
        colWidths=[116 * mm, 50 * mm],
    )

    footer.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.35, colors.black),
            ]
        )
    )

    return [
        header,
        Spacer(1, 4 * mm),
        student_info,
        Spacer(1, 4 * mm),
        fee_table,
        Spacer(1, 4 * mm),
        footer,
    ]


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


# =========================
# APP
# =========================

st.set_page_config(
    page_title=APP_NAME,
    layout="wide"
)

init_database()

st.title("Fee Voucher System")


voucher_tab, history_tab = st.tabs(
    ["Create Voucher", "Voucher History"]
)


# =========================
# CREATE VOUCHER
# =========================

with voucher_tab:

    imported_data = {}

    st.subheader("Import Excel / CSV")

    uploaded_file = st.file_uploader(
        "Upload Excel or CSV File",
        type=["xlsx", "csv"]
    )

    if uploaded_file:

        try:

            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)

            else:
                df = pd.read_excel(uploaded_file)

            st.success("File Imported Successfully")

            st.dataframe(
                df,
                use_container_width=True
            )

            selected_student = st.selectbox(
                "Select Student",
                df.index
            )

            row = df.loc[selected_student]

            imported_data = {
                "student_name": str(row.get("student_name", "")),
                "father_name": str(row.get("father_name", "")),
                "class_name": str(row.get("class", "")),
                "roll_no": str(row.get("roll_no", "")),
                "tuition_fee": int(row.get("tuition_fee", 0)),
                "admission_fee": int(row.get("admission_fee", 0)),
                "annual_fee": int(row.get("annual_fee", 0)),
                "exam_fee": int(row.get("exam_fee", 0)),
                "late_fee": int(row.get("late_fee", 0)),
                "discount": int(row.get("discount", 0)),
            }

            st.success("Student Data Loaded Successfully")

        except Exception as e:
            st.error(f"Import Error: {e}")

    left, right = st.columns([1, 1])

    with left:

        institute = st.selectbox(
            "Institute",
            list(INSTITUTES.keys())
        )

        voucher_no = st.text_input(
            "Voucher No",
            value=next_voucher_no(institute)
        )

        student_name = st.text_input(
            "Student Name",
            value=imported_data.get("student_name", "")
        )

        father_name = st.text_input(
            "Father Name",
            value=imported_data.get("father_name", "")
        )

        class_name = st.text_input(
            "Class",
            value=imported_data.get("class_name", "")
        )

        roll_no = st.text_input(
            "Roll No",
            value=imported_data.get("roll_no", "")
        )

        month = st.text_input(
            "Fee Month",
            value=datetime.now().strftime("%B %Y")
        )

        issue_date = st.date_input(
            "Issue Date",
            value=date.today()
        )

        due_date = st.date_input(
            "Due Date",
            value=date.today() + timedelta(days=10)
        )

        notes = st.text_area(
            "Notes",
            value="Late fee will be charged after due date."
        )

    with right:

        admission_fee = st.number_input(
            "Admission Fee",
            min_value=0,
            value=imported_data.get("admission_fee", 0),
            step=100
        )

        tuition_fee = st.number_input(
            "Tuition Fee",
            min_value=0,
            value=imported_data.get("tuition_fee", 0),
            step=100
        )

        annual_fee = st.number_input(
            "Annual Fee",
            min_value=0,
            value=imported_data.get("annual_fee", 0),
            step=100
        )

        exam_fee = st.number_input(
            "Exam Fee",
            min_value=0,
            value=imported_data.get("exam_fee", 0),
            step=100
        )

        late_fee = st.number_input(
            "Late Fee",
            min_value=0,
            value=imported_data.get("late_fee", 0),
            step=50
        )

        discount = st.number_input(
            "Discount",
            min_value=0,
            value=imported_data.get("discount", 0),
            step=50
        )

        values = {
            "voucher_no": voucher_no,
            "institute": institute,
            "student_name": student_name,
            "father_name": father_name,
            "class_name": class_name,
            "roll_no": roll_no,
            "month": month,
            "issue_date": issue_date.strftime("%d-%m-%Y"),
            "due_date": due_date.strftime("%d-%m-%Y"),
            "admission_fee": int(admission_fee),
            "tuition_fee": int(tuition_fee),
            "annual_fee": int(annual_fee),
            "exam_fee": int(exam_fee),
            "late_fee": int(late_fee),
            "discount": int(discount),
            "notes": notes,
        }

        values["total_amount"] = total_fee(values)

        st.metric(
            "Total Amount",
            money(values["total_amount"])
        )

        pdf_bytes = create_voucher_pdf(values)

        if st.button("Save Record"):

            try:
                save_voucher(values)
                st.success("Voucher Saved")

            except sqlite3.IntegrityError:
                st.error("Voucher Number Already Exists")

        st.download_button(
            "Download PDF",
            data=pdf_bytes,
            file_name=f"{voucher_no}.pdf",
            mime="application/pdf"
        )


# =========================
# HISTORY
# =========================

with history_tab:

    st.subheader("Voucher History")

    query = st.text_input("Search")

    records = get_vouchers(query)

    if records:

        rows = [
            {
                "Voucher No": row[0],
                "Student": row[3],
                "Class": row[5],
                "Month": row[7],
                "Total": money(row[9]),
            }
            for row in records
        ]

        st.dataframe(
            rows,
            use_container_width=True
        )

    else:
        st.info("No Records Found")