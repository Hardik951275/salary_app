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
# REMOVE: from weasyprint import HTML   ← delete this line


app = Flask(__name__)
app.secret_key = "change-this-secret-key"

UPLOAD_FOLDER = "uploads"
STATIC_FOLDER = "static"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(STATIC_FOLDER, exist_ok=True)

EMPLOYEE_FILE = os.path.join(UPLOAD_FOLDER, "employees.xlsx")

# ... keep all your existing helper functions exactly as they are
# (normalize_columns, load_employees, working_days_without_sundays,
#  calculate_salary, money, _find_logo, generate_pdf)



# Company details
COMPANY_NAME = "Six Sense Media"
COMPANY_ADDRESS = (
    "Office no. 404, Jyoti building,\n"
    "Mavdi Main Road, Shree Nath Society,\n"
    "Rajkot, Gujarat 360004"
)
COMPANY_WEB = "www.sixsensemedia.com"
COMPANY_EMAIL = "info@sixsensemedia.com"
COMPANY_PHONE = "+91 91046 84954"
COMPANY_GST = "24AFNFS7501L1ZM"

LOGO_CANDIDATES = [
    os.path.join(STATIC_FOLDER, "SixSense.jpeg"),
    os.path.join(STATIC_FOLDER, "SixSense.jpg"),
    os.path.join(STATIC_FOLDER, "logo.png"),
    os.path.join(STATIC_FOLDER, "logo.jpg"),
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
    return f"Rs. {value:,.2f}"


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
        rightMargin=18,
        leftMargin=18,
        topMargin=15,
        bottomMargin=15,
    )

    styles = getSampleStyleSheet()

    # Styles
    company_name_style = ParagraphStyle(
        "CompanyName", parent=styles["Normal"],
        fontName="Helvetica-Bold", fontSize=14,
        textColor=colors.HexColor("#1a202c"), spaceAfter=1
    )
    small = ParagraphStyle(
        "Small", parent=styles["Normal"],
        fontName="Helvetica", fontSize=7.5, leading=9.5,
        textColor=colors.HexColor("#2d3748")
    )
    small_center = ParagraphStyle(
        "SmallCenter", parent=small, alignment=TA_CENTER
    )
    title_style = ParagraphStyle(
        "TitleCustom", parent=styles["Normal"],
        fontName="Helvetica-Bold", fontSize=13,
        textColor=colors.HexColor("#1a202c"),
        alignment=TA_CENTER, spaceBefore=4, spaceAfter=1
    )
    section_style = ParagraphStyle(
        "Section", parent=styles["Normal"],
        fontName="Helvetica-Bold", fontSize=9.5,
        textColor=colors.HexColor("#2b6cb0"),
        spaceBefore=8, spaceAfter=3
    )
    label_style = ParagraphStyle(
        "Label", parent=styles["Normal"],
        fontName="Helvetica", fontSize=7.5,
        textColor=colors.HexColor("#4a5568")
    )
    value_style = ParagraphStyle(
        "Value", parent=styles["Normal"],
        fontName="Helvetica-Bold", fontSize=8,
        textColor=colors.HexColor("#1a202c")
    )
    net_style = ParagraphStyle(
        "Net", parent=styles["Normal"],
        fontName="Helvetica-Bold", fontSize=12,
        textColor=colors.HexColor("#22543d"), alignment=TA_CENTER
    )

    story = []

    # ========== HEADER ==========
    logo_path = _find_logo()
    left_content = []

    if logo_path:
        try:
            img = Image(logo_path, width=20*mm, height=20*mm)
            left_content.append(img)
        except:
            pass

    company_info = [
        Paragraph(COMPANY_NAME, company_name_style),
        Paragraph(COMPANY_ADDRESS.replace("\n", "<br/>"), small),
        Paragraph(f"{COMPANY_WEB}  |  {COMPANY_EMAIL}  |  {COMPANY_PHONE}", small),
        Paragraph(f"GST No.- {COMPANY_GST}", small),
    ]

    if left_content:
        header_data = [[left_content[0], company_info]]
        header_table = Table(header_data, colWidths=[25*mm, 155*mm])
    else:
        header_data = [[company_info]]
        header_table = Table(header_data, colWidths=[180*mm])

    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(header_table)

    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#2b6cb0")))
    
    # Title
    story.append(Paragraph("SALARY SLIP", title_style))
    story.append(Paragraph(f"<b>{month_name}</b>", small_center))
    story.append(Spacer(1, 8))

    # ========== EMPLOYEE INFO (Grid like your HTML) ==========
    emp_fields = [
        ("Employee ID", str(employee.get("employee_id", ""))),
        ("Employee Name", str(employee.get("employee_name", ""))),
        ("Designation", str(employee.get("designation") or "—")),
        ("Department", str(employee.get("department") or "—")),
        ("Joining Date", str(employee.get("joining_date") or "—")),
        ("PAN Number", str(employee.get("pan_number") or "—")),
        ("Bank Name", str(employee.get("bank_name") or "—")),
        ("Account Number", str(employee.get("bank_account_number") or "—")),
        ("IFSC Code", str(employee.get("ifsc_code") or "—")),
        ("Working Days", str(calc["working_days"])),
        ("Total Calendar Days", str(calc["total_calendar_days"])),
        ("Salary Month", month_name),
    ]

    # Create 3 columns grid
    emp_rows = []
    for i in range(0, 12, 3):
        row = []
        for j in range(3):
            if i + j < len(emp_fields):
                label, value = emp_fields[i + j]
                cell = [
                    Paragraph(label, label_style),
                    Paragraph(value, value_style)
                ]
                row.append(cell)
            else:
                row.append("")
        emp_rows.append(row)

    emp_table = Table(emp_rows, colWidths=[60*mm, 60*mm, 60*mm])
    emp_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#cbd5e0")),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#e2e8f0")),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f7fafc")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(emp_table)
    story.append(Spacer(1, 8))

    # ========== ATTENDANCE ==========
    story.append(Paragraph("Attendance", section_style))

    att_data = [
        [Paragraph("<b>Particulars</b>", value_style), Paragraph("<b>Days</b>", value_style)],
        ["Sunday / Weekly Off", str(calc["sundays"])],
        ["Present Days", str(calc["present_days"])],
        ["Paid Leave", str(calc["paid_leave"])],
        ["Unpaid Leave", str(calc["unpaid_leave"])],
        [Paragraph("<b>Paid Days</b>", value_style), Paragraph(f"<b>{calc['paid_days']}</b>", value_style)],
    ]

    att_table = Table(att_data, colWidths=[130*mm, 50*mm])
    att_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#a0aec0")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2b6cb0")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#ebf8ff")),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(att_table)
    story.append(Spacer(1, 8))

    # ========== SALARY DETAILS ==========
    story.append(Paragraph("Salary Details", section_style))

    sal_data = [
        [Paragraph("<b>Particulars</b>", value_style), Paragraph("<b>Amount</b>", value_style)],
        ["Basic Salary", money(calc["basic_salary"])],
        ["Allowance", money(calc["allowance"])],
        ["Monthly Gross Salary", money(calc["gross_monthly"])],
        ["Per Working Day", money(calc["per_day_salary"])],
        ["Earned Gross Salary", money(calc["earned_gross"])],
        ["Unpaid Leave Deduction", money(calc["unpaid_leave_deduction"])],
        ["Other / Fixed Deduction", money(calc["fixed_deduction"])],
        ["Total Deduction", money(calc["total_deduction"])],
    ]

    sal_table = Table(sal_data, colWidths=[130*mm, 50*mm])
    sal_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#a0aec0")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2b6cb0")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(sal_table)
    story.append(Spacer(1, 8))

    # ========== NET SALARY ==========
    net_data = [[
        Paragraph("NET SALARY", net_style),
        Paragraph(money(calc["net_salary"]), net_style)
    ]]
    net_table = Table(net_data, colWidths=[90*mm, 90*mm])
    net_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#c6f6d5")),
        ("BOX", (0, 0), (-1, -1), 1.5, colors.HexColor("#38a169")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(net_table)

    story.append(Spacer(1, 22))

    # ========== SIGNATURE ==========
    sig_data = [
        ["____________________________", "____________________________"],
        ["Employee Signature", "Authorized Signature"]
    ]
    sig_table = Table(sig_data, colWidths=[90*mm, 90*mm])
    sig_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, 0), 12),
        ("BOTTOMPADDING", (0, 1), (-1, 1), 2),
    ]))
    story.append(sig_table)

    story.append(Spacer(1, 10))
    story.append(Paragraph(
        "This is a computer-generated salary slip and does not require any physical signature",
        small_center
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

    pdf_buffer = generate_pdf(employee, calc, year, month)

    safe_name = str(employee.get("employee_name", "Employee")).replace(" ", "_")
    filename = f"SalarySlip_{safe_name}_{year}_{month:02d}.pdf"

    return send_file(
        pdf_buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf"
    )
    
@app.route("/download_all", methods=["POST"])
def download_all():
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
