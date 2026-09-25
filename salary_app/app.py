from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pandas as pd
from calendar import monthrange
from datetime import date
from io import BytesIO
import os
import zipfile
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, HRFlowable
)
from reportlab.lib.units import mm

app = Flask(__name__)
app.secret_key = "change-this-secret-key"

UPLOAD_FOLDER = "uploads"
STATIC_FOLDER = "static"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(STATIC_FOLDER, exist_ok=True)

EMPLOYEE_FILE = os.path.join(UPLOAD_FOLDER, "employees.xlsx")

# Company details
COMPANY_NAME = "Six Sense Media"
COMPANY_ADDRESS = (
    "Office no. 403, Jyoti building, Sheri.10,\n"
    "Mavdi Main Road, Sri Nath Society,\n"
    "Rajkot, Gujarat 360004"
)
COMPANY_WEB = "www.sixsensemedia.com"
COMPANY_EMAIL = "info@sixsensemedia.com"
COMPANY_PHONE = "+91 91046 8495"
LOGO_CANDIDATES = [
    os.path.join(STATIC_FOLDER, "logo.png"),
    os.path.join(STATIC_FOLDER, "logo.jpg"),
    os.path.join(STATIC_FOLDER, "2.jpeg"),
    os.path.join(STATIC_FOLDER, "2.jpg"),
]


def normalize_columns(df):
    rename = {}
    for col in df.columns:
        key = str(col).strip().lower().replace(" ", "_")
        key = key.replace("-", "_")
        rename[col] = key
    df = df.rename(columns=rename)

    aliases = {
        "employee": "employee_name",
        "name": "employee_name",
        "emp_name": "employee_name",
        "emp_id": "employee_id",
        "present": "present_days",
        "present_day": "present_days",
        "paid_leaves": "paid_leave",
        "leave_paid": "paid_leave",
        "basic": "basic_salary",
        "salary_basic": "basic_salary",
        "allowances": "allowance",
        "salary_allowance": "allowance",
        "deductions": "deduction",
        "designation": "designation",
        "dept": "department",
        "joining": "joining_date",
        "date_of_joining": "joining_date",
        "doj": "joining_date",
        "pan": "pan_number",
        "pan_no": "pan_number",
        "pan_card": "pan_number",
        "bank": "bank_name",
        "bank_account_name": "bank_name",
        "account_name": "bank_name",
        "account_number": "bank_account_number",
        "bank_account": "bank_account_number",
        "account_no": "bank_account_number",
        "a/c_no": "bank_account_number",
        "ifsc": "ifsc_code",
        "ifsc_code": "ifsc_code",
    }

    for old, new in aliases.items():
        if old in df.columns and new not in df.columns:
            df = df.rename(columns={old: new})

    return df


def load_employees():
    if not os.path.exists(EMPLOYEE_FILE):
        return pd.DataFrame()

    df = pd.read_excel(EMPLOYEE_FILE)
    df = normalize_columns(df)

    required_numeric = [
        "present_days", "paid_leave",
        "basic_salary", "allowance", "deduction",
    ]
    required_text = [
        "employee_name", "designation", "department",
        "joining_date", "pan_number", "bank_name",
        "bank_account_number", "ifsc_code",
    ]

    for col in required_numeric:
        if col not in df.columns:
            df[col] = 0
    for col in required_text:
        if col not in df.columns:
            df[col] = ""

    if "employee_id" not in df.columns:
        df["employee_id"] = range(1, len(df) + 1)

    for col in required_numeric:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    for col in required_text + ["employee_id"]:
        df[col] = df[col].fillna("").astype(str)

    # Normalize joining_date display
    def fmt_date(v):
        if not v or str(v).strip() in ("", "nan", "NaT", "None"):
            return ""
        try:
            return pd.to_datetime(v).strftime("%d-%m-%Y")
        except Exception:
            return str(v)

    df["joining_date"] = df["joining_date"].apply(fmt_date)

    return df


def working_days_without_sundays(year, month):
    total_days = monthrange(year, month)[1]
    sundays = sum(
        1 for d in range(1, total_days + 1)
        if date(year, month, d).weekday() == 6
    )
    return total_days, sundays, total_days - sundays


def calculate_salary(row, year, month):
    total_calendar_days, sundays, working_days = working_days_without_sundays(year, month)

    present = float(row.get("present_days", 0))
    paid_leave = float(row.get("paid_leave", 0))
    basic = float(row.get("basic_salary", 0))
    allowance = float(row.get("allowance", 0))
    fixed_deduction = float(row.get("deduction", 0))

    paid_days = present + paid_leave
    paid_days = min(paid_days, working_days)
    unpaid_leave = max(working_days - paid_days, 0)

    gross_monthly = basic + allowance
    per_day_salary = gross_monthly / working_days if working_days else 0
    earned_gross = per_day_salary * paid_days
    unpaid_leave_deduction = per_day_salary * unpaid_leave
    total_deduction = fixed_deduction + unpaid_leave_deduction
    net_salary = max(earned_gross - fixed_deduction, 0)

    return {
        "total_calendar_days": total_calendar_days,
        "sundays": sundays,
        "working_days": working_days,
        "present_days": present,
        "paid_leave": paid_leave,
        "paid_days": paid_days,
        "unpaid_leave": unpaid_leave,
        "basic_salary": basic,
        "allowance": allowance,
        "gross_monthly": gross_monthly,
        "per_day_salary": per_day_salary,
        "earned_gross": earned_gross,
        "unpaid_leave_deduction": unpaid_leave_deduction,
        "fixed_deduction": fixed_deduction,
        "total_deduction": total_deduction,
        "net_salary": net_salary,
    }


def money(value):
    return f"₹ {value:,.2f}"


def _find_logo():
    for path in LOGO_CANDIDATES:
        if os.path.exists(path):
            return path
    return None


def generate_pdf(employee, calc, year, month):
    month_name = date(year, month, 1).strftime("%B %Y")
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=28,
        leftMargin=28,
        topMargin=24,
        bottomMargin=24,
    )

    styles = getSampleStyleSheet()
    company_name_style = ParagraphStyle(
        "CompanyName",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=14,
        textColor=colors.HexColor("#1f4e79"),
        spaceAfter=2,
    )
    small = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#333333"),
    )
    small_center = ParagraphStyle(
        "SmallCenter",
        parent=small,
        alignment=TA_CENTER,
    )
    title = ParagraphStyle(
        "TitleCustom",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=14,
        textColor=colors.HexColor("#1f4e79"),
        spaceAfter=2,
        spaceBefore=4,
    )
    right = ParagraphStyle(
        "Right",
        parent=styles["Normal"],
        alignment=TA_RIGHT,
        fontSize=9,
    )
    section = ParagraphStyle(
        "Section",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        textColor=colors.HexColor("#1f4e79"),
        spaceBefore=8,
        spaceAfter=4,
    )

    story = []

    # ----- Company header -----
    logo_path = _find_logo()
    addr_html = COMPANY_ADDRESS.replace("\n", "<br/>")
    contact_line = f"{COMPANY_WEB}  |  {COMPANY_EMAIL}  |  {COMPANY_PHONE}"

    left_content = []
    if logo_path:
        try:
            img = Image(logo_path, width=18 * mm, height=18 * mm)
            left_content.append(img)
        except Exception:
            pass

    info_paras = [
        Paragraph(COMPANY_NAME, company_name_style),
        Paragraph(addr_html, small),
        Paragraph(contact_line, small),
    ]

    if left_content:
        header_table = Table(
            [[left_content[0], info_paras]],
            colWidths=[22 * mm, 150 * mm],
        )
        header_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(header_table)
    else:
        for p in info_paras:
            story.append(p)

    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1f4e79")))
    story.append(Paragraph("SALARY SLIP", title))
    story.append(Paragraph(f"Monthly Salary Statement — {month_name}", small_center))
    story.append(Spacer(1, 10))

    # ----- Employee info -----
    emp_rows = [
        ["Employee ID", str(employee.get("employee_id", "")),
         "Salary Month", month_name],
        ["Employee Name", str(employee.get("employee_name", "")),
         "Designation", str(employee.get("designation", ""))],
        ["Department", str(employee.get("department", "")),
         "Joining Date", str(employee.get("joining_date", ""))],
        ["PAN Number", str(employee.get("pan_number", "")),
         "Working Days", str(calc["working_days"])],
        ["Bank Name", str(employee.get("bank_name", "")),
         "Account No.", str(employee.get("bank_account_number", ""))],
        ["IFSC Code", str(employee.get("ifsc_code", "")),
         "", ""],
    ]

    t = Table(emp_rows, colWidths=[95, 160, 95, 130])
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f0f4f8")),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#f0f4f8")),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("PADDING", (0, 0), (-1, -1), 5),
        ("SPAN", (1, 5), (3, 5)),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))

    # ----- Attendance -----
    story.append(Paragraph("Attendance", section))
    attendance = [
        ["Particulars", "Days"],
        ["Total Calendar Days", str(calc["total_calendar_days"])],
        ["Sunday / Weekly Off", str(calc["sundays"])],
        ["Total Working Days", str(calc["working_days"])],
        ["Present Days", str(calc["present_days"])],
        ["Paid Leave", str(calc["paid_leave"])],
        ["Unpaid Leave", str(calc["unpaid_leave"])],
        ["Paid Days", str(calc["paid_days"])],
    ]
    at = Table(attendance, colWidths=[340, 140])
    at.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e9eef5")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("PADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(at)
    story.append(Spacer(1, 8))

    # ----- Salary -----
    story.append(Paragraph("Earnings / Deductions", section))
    salary = [
        ["Particulars", "Amount"],
        ["Basic Salary", money(calc["basic_salary"])],
        ["Allowance", money(calc["allowance"])],
        ["Monthly Gross Salary", money(calc["gross_monthly"])],
        ["Per Working Day", money(calc["per_day_salary"])],
        ["Earned Gross Salary", money(calc["earned_gross"])],
        ["Unpaid Leave Deduction", money(calc["unpaid_leave_deduction"])],
        ["Other / Fixed Deduction", money(calc["fixed_deduction"])],
        ["Total Deduction", money(calc["total_deduction"])],
        ["NET SALARY", money(calc["net_salary"])],
    ]
    st = Table(salary, colWidths=[340, 140])
    st.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e9eef5")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#dce8f7")),
        ("PADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(st)
    story.append(Spacer(1, 24))

    sig = Table(
        [["Employee Signature", "Authorized Signature"]],
        colWidths=[240, 240],
    )
    sig.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (0, 0), "LEFT"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 20),
    ]))
    story.append(sig)
    story.append(Spacer(1, 12))
    story.append(Paragraph(
        "This is a computer-generated salary slip.",
        small_center,
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer


@app.route("/")
def index():
    df = load_employees()
    employees = df.to_dict("records") if not df.empty else []
    return render_template("index.html", employees=employees)


@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("excel_file")

    if not file or file.filename == "":
        flash("Please select an Excel file.", "error")
        return redirect(url_for("index"))

    if not file.filename.lower().endswith((".xlsx", ".xls")):
        flash("Please upload an Excel file (.xlsx or .xls).", "error")
        return redirect(url_for("index"))

    file.save(EMPLOYEE_FILE)
    flash("Excel file uploaded successfully.", "success")
    return redirect(url_for("index"))


@app.route("/generate", methods=["POST"])
def generate():
    df = load_employees()

    if df.empty:
        flash("Please upload your employee Excel file first.", "error")
        return redirect(url_for("index"))

    employee_id = str(request.form.get("employee_id", ""))
    year = int(request.form.get("year"))
    month = int(request.form.get("month"))

    matches = df[df["employee_id"].astype(str) == employee_id]
    if matches.empty:
        flash("Employee not found.", "error")
        return redirect(url_for("index"))

    employee = matches.iloc[0].to_dict()
    calc = calculate_salary(employee, year, month)

    return render_template(
        "salary_slip.html",
        employee=employee,
        calc=calc,
        year=year,
        month=month,
        month_name=date(year, month, 1).strftime("%B %Y"),
        company_name=COMPANY_NAME,
        company_address=COMPANY_ADDRESS,
        company_web=COMPANY_WEB,
        company_email=COMPANY_EMAIL,
        company_phone=COMPANY_PHONE,
    )


@app.route("/download", methods=["POST"])
def download():
    df = load_employees()

    employee_id = str(request.form.get("employee_id", ""))
    year = int(request.form.get("year"))
    month = int(request.form.get("month"))

    matches = df[df["employee_id"].astype(str) == employee_id]
    if matches.empty:
        flash("Employee not found.", "error")
        return redirect(url_for("index"))

    employee = matches.iloc[0].to_dict()
    calc = calculate_salary(employee, year, month)

    pdf = generate_pdf(employee, calc, year, month)
    safe_name = str(employee.get("employee_name", "Employee")).replace(" ", "_")
    filename = f"SalarySlip_{safe_name}_{year}_{month:02d}.pdf"

    return send_file(
        pdf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf",
    )


@app.route("/download_all", methods=["POST"])
def download_all():
    """Generate salary slips for ALL employees and return as a ZIP."""
    df = load_employees()

    if df.empty:
        flash("Please upload your employee Excel file first.", "error")
        return redirect(url_for("index"))

    year = int(request.form.get("year"))
    month = int(request.form.get("month"))
    month_name = date(year, month, 1).strftime("%B_%Y")

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, row in df.iterrows():
            employee = row.to_dict()
            calc = calculate_salary(employee, year, month)
            pdf = generate_pdf(employee, calc, year, month)
            safe_name = str(employee.get("employee_name", "Employee")).replace(" ", "_")
            emp_id = str(employee.get("employee_id", ""))
            filename = f"SalarySlip_{emp_id}_{safe_name}_{year}_{month:02d}.pdf"
            zf.writestr(filename, pdf.read())

    zip_buffer.seek(0)
    zip_name = f"SalarySlips_All_{month_name}.zip"

    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip",
    )


if __name__ == "__main__":
    app.run(debug=True)
