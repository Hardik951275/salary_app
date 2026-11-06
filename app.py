from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pandas as pd
from calendar import monthrange
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
import os
import zipfile
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from xml.sax.saxutils import escape


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
    "Office no. 404, Jyoti building,\n"
    "Mavdi Main Road, Shree Nath Society,\n"
    "Rajkot, Gujarat 360004"
)
COMPANY_WEB = "www.sixsensemedia.com"
COMPANY_EMAIL = "info@sixsensemedia.com"
COMPANY_PHONE = "+91 91046 84954"
COMPANY_GST = "24AFNFS7501L1ZM"

LOGO_CANDIDATES = [
    os.path.join(STATIC_FOLDER, "SixSense-letterhead.jpeg"),
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
            return pd.to_datetime(v, dayfirst=True).strftime("%d-%m-%Y")
        except Exception:
            return str(v)

    df["joining_date"] = df["joining_date"].apply(fmt_date)

    return df


def calendar_days_in_month(year, month):
    total_days = monthrange(year, month)[1]
    sundays = sum(
        1 for d in range(1, total_days + 1)
        if date(year, month, d).weekday() == 6
    )
    return total_days, sundays


def _amount(value):
    try:
        amount = Decimal(str(value or 0))
        if not amount.is_finite():
            return Decimal("0")
        return amount
    except (InvalidOperation, ValueError, TypeError):
        return Decimal("0")


def _rounded(value):
    return float(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def calculate_salary(row, year, month):
    total_calendar_days, sundays = calendar_days_in_month(year, month)

    present = max(_amount(row.get("present_days", 0)), Decimal("0"))
    paid_leave = max(_amount(row.get("paid_leave", 0)), Decimal("0"))
    basic = max(_amount(row.get("basic_salary", 0)), Decimal("0"))
    allowance = max(_amount(row.get("allowance", 0)), Decimal("0"))
    fixed_deduction = max(_amount(row.get("deduction", 0)), Decimal("0"))

    # Present Days in the uploaded sheet already includes Sundays and excludes
    # leave. Add paid leave separately; unpaid leave is inferred from the month.
    paid_days = min(present + paid_leave, Decimal(total_calendar_days))
    unpaid_leave = max(Decimal(total_calendar_days) - paid_days, Decimal("0"))
    gross_monthly = basic + allowance
    per_day_salary = gross_monthly / Decimal(total_calendar_days)
    gross_monthly = Decimal(str(_rounded(gross_monthly)))
    unpaid_leave_deduction = Decimal(str(_rounded(per_day_salary * unpaid_leave)))
    earned_gross = max(gross_monthly - unpaid_leave_deduction, Decimal("0"))
    total_deduction = unpaid_leave_deduction + fixed_deduction
    net_salary = max(gross_monthly - total_deduction, Decimal("0"))

    return {
        "total_calendar_days": total_calendar_days,
        "sundays": sundays,
        "payroll_days": total_calendar_days,
        "present_days": _rounded(present),
        "paid_leave": _rounded(paid_leave),
        "paid_days": _rounded(paid_days),
        "unpaid_leave": _rounded(unpaid_leave),
        "basic_salary": _rounded(basic),
        "allowance": _rounded(allowance),
        "gross_monthly": _rounded(gross_monthly),
        "per_day_salary": _rounded(per_day_salary),
        "earned_gross": _rounded(earned_gross),
        "unpaid_leave_deduction": _rounded(unpaid_leave_deduction),
        "fixed_deduction": _rounded(fixed_deduction),
        "total_deduction": _rounded(total_deduction),
        "net_salary": _rounded(net_salary),
    }


def requested_period():
    try:
        year = int(request.form.get("year", ""))
        month = int(request.form.get("month", ""))
        if not 2000 <= year <= 2100 or not 1 <= month <= 12:
            raise ValueError
        return year, month
    except (TypeError, ValueError):
        flash("Please select a valid month and year.", "error")
        return None


def _register_pdf_fonts():
    font_pairs = [
        (
            "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        ),
        (
            "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
            "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf",
        ),
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ),
    ]
    for regular_path, bold_path in font_pairs:
        if os.path.exists(regular_path) and os.path.exists(bold_path):
            regular_name = "SalarySlipSans"
            bold_name = "SalarySlipSans-Bold"
            if regular_name not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(regular_name, regular_path))
                pdfmetrics.registerFont(TTFont(bold_name, bold_path))
                pdfmetrics.registerFontFamily(
                    "SalarySlipSans",
                    normal=regular_name,
                    bold=bold_name,
                    italic=regular_name,
                    boldItalic=bold_name,
                )
            break
    else:
        regular_name, bold_name = "Helvetica", "Helvetica-Bold"

    currency_font = None
    for currency_path in (
        "/System/Library/Fonts/Geneva.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ):
        if os.path.exists(currency_path):
            currency_font = "SalarySlipCurrency"
            if currency_font not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(currency_font, currency_path))
            break
    return regular_name, bold_name, currency_font


PDF_FONT, PDF_FONT_BOLD, PDF_CURRENCY_FONT = _register_pdf_fonts()


def money(value):
    if PDF_CURRENCY_FONT:
        return f'<font name="{PDF_CURRENCY_FONT}">₹</font> {value:.2f}'
    return f"Rs. {value:.2f}"


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
        rightMargin=25,
        leftMargin=25,
        topMargin=156,
        bottomMargin=24,
    )
    styles = getSampleStyleSheet()
    navy = colors.HexColor("#234e70")
    blue = colors.HexColor("#285985")
    text = colors.HexColor("#252525")
    muted = colors.HexColor("#777777")
    pale = colors.HexColor("#f3f7fc")
    stripe = colors.HexColor("#f7f9fb")
    line = colors.HexColor("#dce5ef")

    label_style = ParagraphStyle(
        "EmployeeLabel", parent=styles["Normal"], fontName=PDF_FONT,
        fontSize=8.5, leading=12, textColor=muted,
    )
    value_style = ParagraphStyle(
        "EmployeeValue", parent=styles["Normal"], fontName=PDF_FONT_BOLD,
        fontSize=10.5, leading=15, textColor=text,
    )
    row_label_style = ParagraphStyle(
        "RowLabel", parent=styles["Normal"], fontName=PDF_FONT,
        fontSize=10.5, leading=14, textColor=colors.HexColor("#454545"),
    )
    row_value_style = ParagraphStyle(
        "RowValue", parent=styles["Normal"], fontName=PDF_FONT_BOLD,
        fontSize=10.5, leading=14, textColor=colors.HexColor("#111111"),
        alignment=TA_RIGHT,
    )
    section_style = ParagraphStyle(
        "SlipSection", parent=styles["Normal"], fontName=PDF_FONT_BOLD,
        fontSize=12, leading=15, textColor=blue, spaceBefore=4, spaceAfter=3,
    )
    net_label_style = ParagraphStyle(
        "NetLabel", parent=styles["Normal"], fontName=PDF_FONT_BOLD,
        fontSize=12.5, leading=16, textColor=blue,
    )
    net_value_style = ParagraphStyle(
        "NetValue", parent=net_label_style, alignment=TA_RIGHT,
    )
    footer_style = ParagraphStyle(
        "SlipFooter", parent=styles["Normal"], fontName=PDF_FONT,
        fontSize=8.5, leading=11, alignment=TA_CENTER, textColor=muted,
    )
    story = []
    story.append(Spacer(1, 2))

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
        ("Payroll Days (incl. Sundays)", str(calc["total_calendar_days"])),
        ("Salary Month", month_name),
    ]
    emp_rows = []
    for i in range(0, len(emp_fields), 3):
        row = []
        for j in range(3):
            if i + j < len(emp_fields):
                label, value = emp_fields[i + j]
                row.append([
                    Paragraph(escape(label.upper()), label_style),
                    Paragraph(escape(value), value_style),
                ])
            else:
                row.append("")
        emp_rows.append(row)

    available_width = A4[0] - 50
    emp_table = Table(
        emp_rows,
        colWidths=[available_width / 3] * 3,
        cornerRadii=[5, 5, 5, 5],
    )
    emp_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.8, line),
        ("BACKGROUND", (0, 0), (-1, -1), pale),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 11),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(emp_table)
    story.append(Spacer(1, 8))

    story.append(Paragraph("Attendance", section_style))
    story.append(HRFlowable(width="100%", thickness=0.7, color=line, spaceAfter=7))
    attendance_rows = [
        ("Sundays in Month (included in Present Days)", calc["sundays"]),
        ("Present Days", calc["present_days"]),
        ("Paid Leave", calc["paid_leave"]),
        ("Unpaid Leave", calc["unpaid_leave"]),
        ("Paid Days", calc["paid_days"]),
    ]
    attendance_data = [
        [
            Paragraph(escape(str(label)), row_label_style),
            Paragraph(str(value) if "<font" in str(value) else escape(str(value)), row_value_style),
        ]
        for label, value in attendance_rows
    ]
    att_table = Table(attendance_data, colWidths=[available_width - 75, 75])
    att_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, stripe]),
        ("LINEBELOW", (0, 0), (-1, -1), 0.35, colors.HexColor("#edf1f5")),
    ]))
    story.append(att_table)
    story.append(Spacer(1, 6))

    story.append(Paragraph("Salary Details", section_style))
    story.append(HRFlowable(width="100%", thickness=0.7, color=line, spaceAfter=7))
    salary_rows = [
        ("Basic Salary", money(calc["basic_salary"])),
        ("Allowance", money(calc["allowance"])),
        ("Monthly Gross Salary", money(calc["gross_monthly"])),
        ("Per Calendar Day", money(calc["per_day_salary"])),
        ("Earned Gross Salary", money(calc["earned_gross"])),
        ("Unpaid Leave Deduction", money(calc["unpaid_leave_deduction"])),
        ("Other / Fixed Deduction", money(calc["fixed_deduction"])),
        ("Total Deduction", money(calc["total_deduction"])),
    ]
    salary_data = [
        [
            Paragraph(escape(str(label)), row_label_style),
            Paragraph(str(value) if "<font" in str(value) else escape(str(value)), row_value_style),
        ]
        for label, value in salary_rows
    ]
    salary_table = Table(salary_data, colWidths=[available_width - 115, 115])
    salary_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, stripe]),
        ("LINEBELOW", (0, 0), (-1, -1), 0.35, colors.HexColor("#edf1f5")),
    ]))
    story.append(salary_table)
    story.append(Spacer(1, 14))

    net_table = Table(
        [[
            Paragraph("NET SALARY", net_label_style),
            Paragraph(money(calc["net_salary"]), net_value_style),
        ]],
        colWidths=[available_width / 2, available_width / 2],
        cornerRadii=[5, 5, 5, 5],
    )
    net_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#dce8f7")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(net_table)
    story.append(Spacer(1, 15))
    story.append(Paragraph(
        "This is a computer-generated salary slip and does not require any physical signature",
        footer_style,
    ))

    def draw_page_header(canvas, document):
        canvas.saveState()
        page_width, page_height = A4
        canvas.setFillColor(navy)
        canvas.rect(0, page_height - 3, page_width, 3, stroke=0, fill=1)

        logo_path = _find_logo()
        if logo_path:
            canvas.drawImage(
                logo_path, 25, page_height - 64,
                width=92, height=38, preserveAspectRatio=True, mask="auto",
            )

        canvas.setFillColor(muted)
        canvas.setFont(PDF_FONT, 11)
        canvas.drawRightString(page_width - 25, page_height - 29, "Salary Slip")
        canvas.setFillColor(blue)
        canvas.setFont(PDF_FONT_BOLD, 13)
        canvas.drawRightString(page_width - 25, page_height - 51, month_name)

        canvas.setFillColor(colors.HexColor("#4b4b4b"))
        canvas.setFont(PDF_FONT, 9.5)
        header_y = page_height - 77
        for line_text in COMPANY_ADDRESS.splitlines():
            canvas.drawString(25, header_y, line_text)
            header_y -= 13
        canvas.setFillColor(blue)
        canvas.drawString(25, header_y - 3, COMPANY_WEB)
        web_width = pdfmetrics.stringWidth(COMPANY_WEB, PDF_FONT, 9.5)
        canvas.setFillColor(muted)
        contact_x = 25 + web_width + 8
        canvas.drawString(contact_x, header_y - 3, "|  " + COMPANY_EMAIL + "  |  " + COMPANY_PHONE)
        canvas.setFillColor(text)
        canvas.setFont(PDF_FONT, 12)
        canvas.drawString(25, header_y - 21, f"GST No.- {COMPANY_GST}")

        canvas.setStrokeColor(blue)
        canvas.setLineWidth(1.5)
        canvas.line(25, page_height - 149, page_width - 25, page_height - 149)
        canvas.restoreState()

    doc.build(story, onFirstPage=draw_page_header)
    buffer.seek(0)
    return buffer


@app.route("/")
def index():
    df = load_employees()
    employees = df.to_dict("records") if not df.empty else []
    return render_template("index.html", employees=employees, current_month=date.today().month,
                           current_year=date.today().year)


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
    period = requested_period()
    if period is None:
        return redirect(url_for("index"))
    year, month = period

    matches = df[df["employee_id"].astype(str) == employee_id]
    if matches.empty:
        flash("Employee not found.", "error")
        return redirect(url_for("index"))

    employee = matches.iloc[0].to_dict()
    return render_template(
        "salary_slip.html",
        employee=employee,
        year=year,
        month=month,
        month_name=date(year, month, 1).strftime("%B %Y"),
        preview_pdf_url=url_for(
            "preview_pdf", employee_id=employee_id, year=year, month=month
        ),
    )


@app.route("/preview_pdf/<employee_id>/<int:year>/<int:month>")
def preview_pdf(employee_id, year, month):
    if not 2000 <= year <= 2100 or not 1 <= month <= 12:
        flash("Please select a valid month and year.", "error")
        return redirect(url_for("index"))

    df = load_employees()
    matches = df[df["employee_id"].astype(str) == str(employee_id)]
    if matches.empty:
        flash("Employee not found.", "error")
        return redirect(url_for("index"))

    employee = matches.iloc[0].to_dict()
    calc = calculate_salary(employee, year, month)
    pdf_buffer = generate_pdf(employee, calc, year, month)
    safe_name = re.sub(
        r"[^A-Za-z0-9_-]", "_", str(employee.get("employee_name", "Employee"))
    )
    return send_file(
        pdf_buffer,
        as_attachment=False,
        download_name=f"SalarySlip_{safe_name}_{year}_{month:02d}.pdf",
        mimetype="application/pdf",
    )


@app.route("/download", methods=["POST"])
def download():
    df = load_employees()
    if df.empty:
        flash("Please upload your employee Excel file first.", "error")
        return redirect(url_for("index"))

    employee_id = str(request.form.get("employee_id", ""))
    period = requested_period()
    if period is None:
        return redirect(url_for("index"))
    year, month = period

    matches = df[df["employee_id"].astype(str) == employee_id]
    if matches.empty:
        flash("Employee not found.", "error")
        return redirect(url_for("index"))

    employee = matches.iloc[0].to_dict()
    calc = calculate_salary(employee, year, month)

    pdf_buffer = generate_pdf(employee, calc, year, month)
    safe_name = re.sub(r"[^A-Za-z0-9_-]", "_", str(employee.get("employee_name", "Employee")))
    filename = f"SalarySlip_{safe_name}_{year}_{month:02d}.pdf"
    return send_file(
        pdf_buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf",
    )


@app.route("/download_all", methods=["POST"])
def download_all():
    df = load_employees()

    if df.empty:
        flash("Please upload your employee Excel file first.", "error")
        return redirect(url_for("index"))

    period = requested_period()
    if period is None:
        return redirect(url_for("index"))
    year, month = period
    month_name = date(year, month, 1).strftime("%B_%Y")

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, row in df.iterrows():
            employee = row.to_dict()
            calc = calculate_salary(employee, year, month)
            pdf = generate_pdf(employee, calc, year, month)
            safe_name = re.sub(r"[^A-Za-z0-9_-]", "_", str(employee.get("employee_name", "Employee")))
            emp_id = re.sub(r"[^A-Za-z0-9_-]", "_", str(employee.get("employee_id", "")))
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
