import base64
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from fpdf import FPDF


st.set_page_config(page_title="Flat Quotation Generator", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
MASTER_FILE = BASE_DIR / "master_cost_sheet.xlsx"
USERS_FILE = BASE_DIR / "users.xlsx"
LOGO_FILE = BASE_DIR / "logo.png"

BRAND_RED = (179, 32, 42)
IST = ZoneInfo("Asia/Kolkata")


def now_ist():
    return datetime.now(IST)


def indian_format(value):
    try:
        if pd.isna(value):
            return ""
        value = round(float(str(value).replace(",", "")))
        s = str(abs(value))
        last3 = s[-3:]
        rest = s[:-3]
        if rest:
            parts = []
            while len(rest) > 2:
                parts.insert(0, rest[-2:])
                rest = rest[:-2]
            if rest:
                parts.insert(0, rest)
            out = ",".join(parts + [last3])
        else:
            out = last3
        return "-" + out if value < 0 else out
    except Exception:
        return str(value)


def to_num(value):
    try:
        if pd.isna(value):
            return 0
        return float(str(value).replace(",", "").strip())
    except Exception:
        return 0


def clean_mobile(value):
    return "".join(filter(str.isdigit, str(value)))[:10]


def normalize_col(col):
    return "".join(ch.lower() for ch in str(col).strip() if ch.isalnum())


def load_master():
    df = pd.read_excel(MASTER_FILE)
    df.columns = [str(c).strip() for c in df.columns]

    alias = {
        "TOWER": ["tower"],
        "FLOOR": ["floor"],
        "FLAT NO.": ["flatno", "flatnumber", "flatno."],
        "UNIT TYPE": ["unittype"],
        "SUPER BUILT UP AREA (SQ.FT.)": ["superbuiltupareasqft", "sbasqft"],
        "RERA CARPET AREA (SQ.FT.)": ["reracarpetareasqft"],
        "BASE RATE": ["baserate"],
        "PLC": ["plc"],
        "FLOOR RISE": ["floorrise"],
        "TOTAL BASE AMOUNT PER SQ.FT.": ["totalbaseamountpersqft", "baseamount"],
        "CAR PARK": ["carpark", "parking"],
        "CLUB HOUSE CHARGES": ["clubhousecharges"],
        "INFRA CHARGES ( INCLUDING POWER, WATER, GENERATOR, STP CHARGES ETC.) PER SQ.FT": [
            "infrachargesincludingpowerwatergeneratorstpchargesetcpersqft",
            "infrachargespersqft",
            "commoninfracharges",
        ],
        "TOTAL BASE PRICE": ["totalbaseprice"],
        "GST ON THE TOTAL BASE PRICE": ["gstonthetotalbaseprice"],
        "LEGAL & DOCUMENTATION CHARGES": ["legaldocumentationcharges"],
        "ADVANCE MAINTENANCE FOR 1 YEAR": ["advancemaintenancefor1year"],
        "GST FOR LEGAL & DOCUMENTATION CHARGES AND ADVANCE MAINTENANCE CHARGES": [
            "gstforlegaldocumentationchargesandadvancemaintenancecharges",
            "gstforlegaldocumentationchargesandadvancemaintenance",
        ],
        "SINKING FUND ( INTEREST FREE - PASS THROUGH TO OWNER'S ASSOCIATION )": [
            "sinkingfundinterestfreepassthroughtoownersassociation",
            "sinkingfund",
            "sinkingfundpassroughtoownersassociation",
            "sinkingfundinterestfree",
        ],
        "TOTAL PRICE": ["totalprice", "totalamountwithgst"],
    }

    rename = {}
    actual = {normalize_col(c): c for c in df.columns}

    for standard, options in alias.items():
        for opt in options:
            if opt in actual:
                rename[actual[opt]] = standard
                break

    # Extra fallback for sinking fund column
    for c in df.columns:
        nc = normalize_col(c)
        if "sinkingfund" in nc:
            rename[c] = "SINKING FUND ( INTEREST FREE - PASS THROUGH TO OWNER'S ASSOCIATION )"

    df = df.rename(columns=rename)

    for r in ["TOWER", "FLOOR", "FLAT NO."]:
        if r not in df.columns:
            st.error(f"Missing column in master_cost_sheet.xlsx: {r}")
            st.stop()

    df["TOWER"] = df["TOWER"].astype(str).str.strip()
    df["FLOOR"] = df["FLOOR"].astype(str).str.strip()
    df["FLAT NO."] = df["FLAT NO."].astype(str).str.strip()

    return df


def load_users():
    df = pd.read_excel(USERS_FILE)
    df.columns = [str(c).strip() for c in df.columns]

    for col in ["USERNAME", "PASSWORD", "SALES PERSON NAME"]:
        if col not in df.columns:
            st.error(f"Missing column in users.xlsx: {col}")
            st.stop()

    df["USERNAME"] = df["USERNAME"].astype(str).str.strip()
    df["PASSWORD"] = df["PASSWORD"].astype(str).str.strip()
    df["SALES PERSON NAME"] = df["SALES PERSON NAME"].astype(str).str.strip()
    return df


def get_value(row, col):
    if col in row.index:
        return row[col]

    target = normalize_col(col)

    for actual_col in row.index:
        if normalize_col(actual_col) == target:
            return row[actual_col]

    if "sinkingfund" in target:
        for actual_col in row.index:
            if "sinkingfund" in normalize_col(actual_col):
                return row[actual_col]

    return ""


def get_number(row, col):
    return to_num(get_value(row, col))


def get_logo_base64():
    if not LOGO_FILE.exists():
        return ""
    with open(LOGO_FILE, "rb") as f:
        return base64.b64encode(f.read()).decode()


st.markdown(
    """
<style>
.block-container {
    max-width: 1180px;
    padding-top: 2rem;
}
.logo-box {
    text-align: center;
}
.main-title {
    text-align: center;
    font-size: 44px;
    font-weight: 800;
    color: #262730 !important;
    margin-top: 18px;
    margin-bottom: 30px;
}
.login-card {
    background: #ffffff;
    padding: 22px 30px 18px 30px;
    border-radius: 22px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.08);
    margin-bottom: 22px;
}
.login-heading {
    text-align: center;
    font-size: 24px;
    font-weight: 800;
    color: #262730 !important;
}
.login-subtitle {
    text-align: center;
    font-size: 14px;
    color: #777777 !important;
}
.welcome-box {
    background: #E8F5E9;
    padding: 12px 18px;
    border-radius: 10px;
    color: #1B5E20 !important;
    font-weight: 600;
}
.section-heading {
    font-size: 28px;
    font-weight: 800;
    color: #262730 !important;
    margin-top: 18px;
}
.stButton > button, .stDownloadButton > button {
    background-color: #B3202A;
    color: white !important;
    border-radius: 12px;
    height: 50px;
    font-size: 17px;
    font-weight: 700;
    border: none;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    background-color: #931923;
    color: white !important;
}
.preview-table {
    width: 100%;
    border-collapse: collapse;
    background: #ffffff !important;
}
.preview-table td {
    border: 1px solid #d5d5d5;
    padding: 9px 12px;
    color: #111111 !important;
    background: #ffffff !important;
    font-weight: 500;
}
.preview-section td {
    background: #f3f3f3 !important;
    color: #111111 !important;
    font-weight: 800 !important;
}
.preview-total td {
    background: #B3202A !important;
    color: #ffffff !important;
    font-weight: 800 !important;
}
</style>
""",
    unsafe_allow_html=True,
)


def show_logo(width=280):
    logo64 = get_logo_base64()
    if logo64:
        st.markdown(
            f"""
            <div class="logo-box">
                <img src="data:image/png;base64,{logo64}" width="{width}">
            </div>
            """,
            unsafe_allow_html=True,
        )


def preview_rows(row):
    return [
        ("section", "UNIT DETAILS", ""),
        ("normal", "TOWER", get_value(row, "TOWER")),
        ("normal", "FLOOR", get_value(row, "FLOOR")),
        ("normal", "FLAT NO.", get_value(row, "FLAT NO.")),
        ("normal", "UNIT TYPE", get_value(row, "UNIT TYPE")),
        ("normal", "SUPER BUILT UP AREA (SQ.FT.)", indian_format(get_value(row, "SUPER BUILT UP AREA (SQ.FT.)"))),
        ("normal", "RERA CARPET AREA (SQ.FT.)", indian_format(get_value(row, "RERA CARPET AREA (SQ.FT.)"))),
        ("section", "BASE COST DETAILS", ""),
        ("normal", "BASE RATE", indian_format(get_value(row, "BASE RATE"))),
        ("normal", "PLC", indian_format(get_value(row, "PLC"))),
        ("normal", "FLOOR RISE", indian_format(get_value(row, "FLOOR RISE"))),
        ("normal", "TOTAL BASE AMOUNT PER SQ.FT.", indian_format(get_value(row, "TOTAL BASE AMOUNT PER SQ.FT."))),
        ("section", "OTHER CHARGES", ""),
        ("normal", "CAR PARK", indian_format(get_value(row, "CAR PARK"))),
        ("normal", "CLUB HOUSE CHARGES", indian_format(get_value(row, "CLUB HOUSE CHARGES"))),
        ("normal", "INFRA CHARGES ( INCLUDING POWER, WATER, GENERATOR, STP CHARGES ETC.) PER SQ.FT", indian_format(get_value(row, "INFRA CHARGES ( INCLUDING POWER, WATER, GENERATOR, STP CHARGES ETC.) PER SQ.FT"))),
        ("total", "TOTAL BASE PRICE", indian_format(get_value(row, "TOTAL BASE PRICE"))),
        ("normal", "GST ON THE TOTAL BASE PRICE", indian_format(get_value(row, "GST ON THE TOTAL BASE PRICE"))),
        ("normal", "LEGAL & DOCUMENTATION CHARGES", indian_format(get_value(row, "LEGAL & DOCUMENTATION CHARGES"))),
        ("normal", "ADVANCE MAINTENANCE FOR 1 YEAR", indian_format(get_value(row, "ADVANCE MAINTENANCE FOR 1 YEAR"))),
        ("normal", "GST FOR LEGAL & DOCUMENTATION CHARGES AND ADVANCE MAINTENANCE CHARGES", indian_format(get_value(row, "GST FOR LEGAL & DOCUMENTATION CHARGES AND ADVANCE MAINTENANCE CHARGES"))),
        ("normal", "SINKING FUND ( INTEREST FREE - PASS THROUGH TO OWNER'S ASSOCIATION )", indian_format(get_value(row, "SINKING FUND ( INTEREST FREE - PASS THROUGH TO OWNER'S ASSOCIATION )"))),
        ("section", "FINAL PAYABLE AMOUNT", ""),
        ("total", "TOTAL PRICE", indian_format(get_value(row, "TOTAL PRICE"))),
    ]


def preview_html(row):
    html = '<table class="preview-table">'
    for typ, label, value in preview_rows(row):
        if typ == "section":
            html += f'<tr class="preview-section"><td colspan="2">{label}</td></tr>'
        elif typ == "total":
            html += f'<tr class="preview-total"><td>{label}</td><td style="text-align:right;">{value}</td></tr>'
        else:
            html += f'<tr><td>{label}</td><td style="text-align:right;">{value}</td></tr>'
    html += "</table>"
    return html


class PDF(FPDF):
    def __init__(self, generated_on, generated_by):
        super().__init__("P", "mm", "A4")
        self.generated_on = generated_on
        self.generated_by = generated_by
        self.set_auto_page_break(False)
        self.set_margins(12, 8, 12)

    def footer(self):
        self.set_y(-14)
        self.set_font("Arial", "I", 8)
        self.set_text_color(90, 90, 90)
        self.cell(90, 6, f"Generated on: {self.generated_on}", 0, 0, "L")
        self.cell(90, 6, f"Generated By: {self.generated_by}", 0, 0, "R")


def wrap_lines(pdf, text, width):
    text = "" if text is None else str(text)
    words = text.split()
    lines, line = [], ""
    for w in words:
        test = w if not line else line + " " + w
        if pdf.get_string_width(test) <= width:
            line = test
        else:
            if line:
                lines.append(line)
            line = w
    if line:
        lines.append(line)
    return lines or [""]


def box_text(pdf, x, y, w, h, text, align="L", style="", size=9, color=(0, 0, 0), pad=2):
    pdf.set_font("Arial", style, size)
    pdf.set_text_color(*color)
    lines = wrap_lines(pdf, text, w - 2 * pad)
    line_h = 4.5
    start_y = y + max(0, (h - len(lines) * line_h) / 2)
    for i, line in enumerate(lines):
        pdf.set_xy(x + pad, start_y + i * line_h)
        pdf.cell(w - 2 * pad, line_h, line, 0, 0, align)


def add_logo_pdf(pdf):
    if LOGO_FILE.exists():
        pdf.image(str(LOGO_FILE), x=78, y=7, w=54)


def red_title(pdf, title, y=26):
    pdf.set_fill_color(*BRAND_RED)
    pdf.rect(12, y, 186, 11, "DF")
    pdf.set_font("Arial", "B", 16)
    pdf.set_text_color(255, 255, 255)
    pdf.set_xy(12, y + 2.5)
    pdf.cell(186, 5, title, 0, 0, "C")
    pdf.set_text_color(0, 0, 0)
    pdf.set_y(y + 11)


def add_date_row(pdf, date_text):
    y = pdf.get_y()
    pdf.rect(12, y, 186, 8)
    box_text(pdf, 12, y, 186, 8, f"Date: {date_text}", "R", "", 10)
    pdf.set_y(y + 8)


def section_row(pdf, title):
    y = pdf.get_y()
    pdf.set_fill_color(245, 245, 245)
    pdf.rect(12, y, 186, 7, "DF")
    box_text(pdf, 12, y, 186, 7, title, "L", "B", 9.5, (0, 0, 0))
    pdf.set_y(y + 7)


def cost_row(pdf, label, value, highlight=False):
    left_w = 141
    right_w = 45
    x = 12
    y = pdf.get_y()

    pdf.set_font("Arial", "", 8.5)
    left_lines = wrap_lines(pdf, label, left_w - 4)
    right_lines = wrap_lines(pdf, value, right_w - 4)
    h = max(7, max(len(left_lines), len(right_lines)) * 4.5 + 2)

    if highlight:
        pdf.set_fill_color(*BRAND_RED)
        fill = "DF"
        color = (255, 255, 255)
        style = "B"
    else:
        fill = ""
        color = (0, 0, 0)
        style = ""

    pdf.rect(x, y, left_w, h, fill)
    pdf.rect(x + left_w, y, right_w, h, fill)
    box_text(pdf, x, y, left_w, h, label, "L", style, 8.8, color)
    box_text(pdf, x + left_w, y, right_w, h, value, "R", style, 8.8, color)
    pdf.set_y(y + h)


def make_payment_schedule(row):
    total_base = get_number(row, "TOTAL BASE PRICE")
    legal = get_number(row, "LEGAL & DOCUMENTATION CHARGES")
    adv = get_number(row, "ADVANCE MAINTENANCE FOR 1 YEAR")
    gst_other = get_number(row, "GST FOR LEGAL & DOCUMENTATION CHARGES AND ADVANCE MAINTENANCE CHARGES")
    sinking = get_number(row, "SINKING FUND ( INTEREST FREE - PASS THROUGH TO OWNER'S ASSOCIATION )")
    total_price = get_number(row, "TOTAL PRICE")

    cgst_rate = 0.025
    sgst_rate = 0.025

    initial = 500000
    first_10 = round(total_base * 0.10)
    balance = max(0, first_10 - initial)

    base_rows = [
        (1, "Initial booking amount", "", initial),
        (2, "Balance booking amount (to be paid within 15 days of Booking)", "", balance),
        (3, "On the execution of Agreement to Sale (Payable within 30 days from booking)", "10%", total_base * 0.10),
        (4, "On Completion of Mass Excavation of the respective wing", "10%", total_base * 0.10),
        (5, "On completion of Foundation of the respective wing", "8%", total_base * 0.08),
        (6, "On Completion of Basement -1 Slab", "7%", total_base * 0.07),
        (7, "On Completion of 01st Floor Slab", "4%", total_base * 0.04),
        (8, "On Completion of 02nd Floor Slab", "4%", total_base * 0.04),
        (9, "On Completion of 04th Floor Slab", "4%", total_base * 0.04),
        (10, "On Completion of 06th Floor Slab", "4%", total_base * 0.04),
        (11, "On Completion of 08th Floor Slab", "3%", total_base * 0.03),
        (12, "On Completion of 10th Floor Slab", "3%", total_base * 0.03),
        (13, "On Completion of 12th Floor Slab", "3%", total_base * 0.03),
        (14, "On Completion of 14th Floor Slab", "3%", total_base * 0.03),
        (15, "On Completion of 16th Floor Slab", "3%", total_base * 0.03),
        (16, "On Completion of 18th Floor Slab", "3%", total_base * 0.03),
        (17, "On Completion of 20th Floor Slab", "3%", total_base * 0.03),
        (18, "On Completion of Terrace Slab", "3%", total_base * 0.03),
        (19, "On Completion of internal Flooring of the said Apartment", "5%", total_base * 0.05),
        (20, "On Completion of First coat of Internal Paint of the said Apartment", "5%", total_base * 0.05),
        (21, "On Intimation of Possession", "5%", total_base * 0.05),
    ]

    rows = []
    for sr, milestone, pct, basic in base_rows:
        basic = round(basic)
        cgst = round(basic * cgst_rate)
        sgst = round(basic * sgst_rate)
        rows.append({
            "sr": sr, "milestone": milestone, "pct": pct,
            "basic": basic, "cgst": cgst, "sgst": sgst,
            "total": basic + cgst + sgst
        })

    other_basic = round(legal + adv)
    other_cgst = round(gst_other / 2)
    other_sgst = round(gst_other / 2)
    rows.append({
        "sr": 22,
        "milestone": "On Intimation of Possession - Other Charges (Legal & Documentation and Maintenance Charges)",
        "pct": "",
        "basic": other_basic,
        "cgst": other_cgst,
        "sgst": other_sgst,
        "total": other_basic + other_cgst + other_sgst
    })

    rows.append({
        "sr": 23,
        "milestone": "On Intimation of Possession - Other Charges (Sinking Fund)",
        "pct": "",
        "basic": round(sinking),
        "cgst": 0,
        "sgst": 0,
        "total": round(sinking)
    })

    totals = {
        "basic": sum(r["basic"] for r in rows),
        "cgst": sum(r["cgst"] for r in rows),
        "sgst": sum(r["sgst"] for r in rows),
        "total": round(total_price) if total_price else sum(r["total"] for r in rows)
    }

    return rows, totals


def payment_row_height(pdf, row, widths):
    pdf.set_font("Arial", "", 7)
    lines = wrap_lines(pdf, row["milestone"], widths[1] - 3)
    return max(5.8, len(lines) * 3.5 + 1.5)


def pay_cell(pdf, x, y, w, h, text, align="C", style="", size=7, color=(0, 0, 0)):
    box_text(pdf, x, y, w, h, text, align, style, size, color, pad=1.2)


def draw_payment_table(pdf, row):
    rows, totals = make_payment_schedule(row)

    pdf.add_page()
    red_title(pdf, "PAYMENT SCHEDULE", y=18)

    y = pdf.get_y()
    pdf.rect(12, y, 93, 8)
    pdf.rect(105, y, 93, 8)
    box_text(pdf, 12, y, 93, 8, f"Flat No.: {get_value(row, 'FLAT NO.')}", "L", "", 9.5)
    box_text(pdf, 105, y, 93, 8, f"Date: {now_ist().strftime('%d-%m-%Y')}", "R", "", 9.5)
    pdf.set_y(y + 11)

    widths = [10, 84, 13, 25, 18, 18, 18]
    headers = ["Sr.", "Milestone", "%", "Basic", "CGST", "SGST", "Total"]

    x = 12
    y = pdf.get_y()
    pdf.set_fill_color(*BRAND_RED)
    for i, htxt in enumerate(headers):
        pdf.rect(x, y, widths[i], 7, "DF")
        pay_cell(pdf, x, y, widths[i], 7, htxt, "C", "B", 8, (255, 255, 255))
        x += widths[i]
    pdf.set_y(y + 7)

    r1, r2 = rows[0], rows[1]
    h1 = payment_row_height(pdf, r1, widths)
    h2 = payment_row_height(pdf, r2, widths)
    y = pdf.get_y()

    sr_w, ms_w, pct_w, basic_w, cgst_w, sgst_w, total_w = widths

    x = 12
    pdf.rect(x, y, sr_w, h1); pay_cell(pdf, x, y, sr_w, h1, r1["sr"]); x += sr_w
    pdf.rect(x, y, ms_w, h1); pay_cell(pdf, x, y, ms_w, h1, r1["milestone"], "L"); x += ms_w
    pdf.rect(x, y, pct_w, h1 + h2); pay_cell(pdf, x, y, pct_w, h1 + h2, "10%"); x += pct_w

    for key, w in zip(["basic", "cgst", "sgst", "total"], widths[3:]):
        pdf.rect(x, y, w, h1)
        pay_cell(pdf, x, y, w, h1, indian_format(r1[key]), "R")
        x += w

    y2 = y + h1
    x = 12
    pdf.rect(x, y2, sr_w, h2); pay_cell(pdf, x, y2, sr_w, h2, r2["sr"]); x += sr_w
    pdf.rect(x, y2, ms_w, h2); pay_cell(pdf, x, y2, ms_w, h2, r2["milestone"], "L"); x += ms_w + pct_w

    for key, w in zip(["basic", "cgst", "sgst", "total"], widths[3:]):
        pdf.rect(x, y2, w, h2)
        pay_cell(pdf, x, y2, w, h2, indian_format(r2[key]), "R")
        x += w

    pdf.set_y(y + h1 + h2)

    for r in rows[2:]:
        y = pdf.get_y()
        h = payment_row_height(pdf, r, widths)
        x = 12
        vals = [
            r["sr"], r["milestone"], r["pct"],
            indian_format(r["basic"]),
            "-" if r["cgst"] == 0 else indian_format(r["cgst"]),
            "-" if r["sgst"] == 0 else indian_format(r["sgst"]),
            indian_format(r["total"]),
        ]
        aligns = ["C", "L", "C", "R", "R", "R", "R"]
        for val, w, al in zip(vals, widths, aligns):
            pdf.rect(x, y, w, h)
            pay_cell(pdf, x, y, w, h, val, al)
            x += w
        pdf.set_y(y + h)

    y = pdf.get_y()
    h = 8
    x = 12
    pdf.set_fill_color(*BRAND_RED)
    merged = widths[0] + widths[1]
    pdf.rect(x, y, merged, h, "DF")
    pay_cell(pdf, x, y, merged, h, "Total", "C", "B", 8, (255, 255, 255))
    x += merged

    vals = ["100%", indian_format(totals["basic"]), indian_format(totals["cgst"]), indian_format(totals["sgst"]), indian_format(totals["total"])]
    for val, w in zip(vals, widths[2:]):
        pdf.rect(x, y, w, h, "DF")
        pay_cell(pdf, x, y, w, h, val, "R" if val != "100%" else "C", "B", 8, (255, 255, 255))
        x += w


def generate_pdf(row, customer_name, customer_mobile, salesperson):
    generated_on = now_ist().strftime("%d-%m-%Y %I:%M %p")
    pdf = PDF(generated_on=generated_on, generated_by=salesperson)

    pdf.add_page()
    add_logo_pdf(pdf)
    red_title(pdf, "COST SHEET & PAYMENT PLAN", y=25)
    add_date_row(pdf, now_ist().strftime("%d-%m-%Y"))
    pdf.set_y(48)

    if customer_name:
        y = pdf.get_y()
        pdf.rect(12, y, 186, 7)
        box_text(pdf, 12, y, 186, 7, f"Name: {customer_name}", "L", "", 9.5)
        pdf.set_y(y + 7)

    if customer_mobile:
        y = pdf.get_y()
        pdf.rect(12, y, 186, 7)
        box_text(pdf, 12, y, 186, 7, f"Mobile: {customer_mobile}", "L", "", 9.5)
        pdf.set_y(y + 7)

    for typ, label, value in preview_rows(row):
        if typ == "section":
            section_row(pdf, label)
        elif typ == "total":
            cost_row(pdf, label, value, highlight=True)
        else:
            cost_row(pdf, label, value, highlight=False)

    if pdf.get_y() < 246:
        pdf.set_y(246)
        pdf.set_fill_color(*BRAND_RED)
        pdf.rect(12, pdf.get_y(), 186, 9, "DF")
        box_text(pdf, 12, pdf.get_y(), 186, 9, "Important Note", "L", "B", 9.5, (255, 255, 255))
        pdf.set_y(pdf.get_y() + 9)
        pdf.rect(12, pdf.get_y(), 186, 10)
        box_text(
            pdf,
            12,
            pdf.get_y(),
            186,
            10,
            "This quotation is indicative and subject to final approval. Prices, GST and other charges are subject to change as applicable.",
            "L",
            "I",
            8.8
        )

    draw_payment_table(pdf, row)

    out = pdf.output(dest="S")
    if isinstance(out, str):
        return out.encode("latin-1")
    return bytes(out)


if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "salesperson" not in st.session_state:
    st.session_state.salesperson = ""

if not st.session_state.logged_in:
    show_logo(260)
    st.markdown('<div class="main-title">Flat Quotation Generator</div>', unsafe_allow_html=True)

    users_df = load_users()

    c1, c2, c3 = st.columns([1.3, 1, 1.3])
    with c2:
        st.markdown(
            """
            <div class="login-card">
                <div class="login-heading">Sales Login</div>
                <div class="login-subtitle">Enter your credentials to continue</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        username = st.text_input("Username", placeholder="Enter username")
        password = st.text_input("Password", type="password", placeholder="Enter password")

        if st.button("Login", use_container_width=True):
            match = users_df[
                (users_df["USERNAME"] == str(username).strip()) &
                (users_df["PASSWORD"] == str(password).strip())
            ]
            if match.empty:
                st.error("Invalid username or password.")
            else:
                st.session_state.logged_in = True
                st.session_state.salesperson = str(match.iloc[0]["SALES PERSON NAME"]).strip()
                st.rerun()

    st.stop()


df = load_master()

top_left, top_right = st.columns([8, 1])
with top_right:
    if st.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.salesperson = ""
        st.rerun()

show_logo(280)
st.markdown('<div class="main-title">Flat Quotation Generator</div>', unsafe_allow_html=True)

st.markdown(f'<div class="welcome-box">Welcome, {st.session_state.salesperson}</div>', unsafe_allow_html=True)

st.markdown('<div class="section-heading">Select Flat Details</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)

with col1:
    tower = st.selectbox("Select Tower", sorted(df["TOWER"].dropna().astype(str).unique()))

tower_df = df[df["TOWER"].astype(str) == str(tower)]

with col2:
    floors = sorted(
        tower_df["FLOOR"].dropna().astype(str).unique(),
        key=lambda x: float(x) if str(x).replace(".", "", 1).isdigit() else str(x)
    )
    floor = st.selectbox("Select Floor", floors)

floor_df = tower_df[tower_df["FLOOR"].astype(str) == str(floor)]

with col3:
    flats = sorted(
        floor_df["FLAT NO."].dropna().astype(str).unique(),
        key=lambda x: float(x) if str(x).replace(".", "", 1).isdigit() else str(x)
    )
    flat = st.selectbox("Select Flat No.", flats)

if "show_quote" not in st.session_state:
    st.session_state.show_quote = False

b1, b2 = st.columns([3, 1])
with b1:
    if st.button("Generate Quotation", use_container_width=True):
        st.session_state.show_quote = True

with b2:
    if st.button("Clear All", use_container_width=True):
        st.session_state.show_quote = False
        st.rerun()

if st.session_state.show_quote:
    selected_df = floor_df[floor_df["FLAT NO."].astype(str) == str(flat)]

    if selected_df.empty:
        st.error("Selected flat not found.")
        st.stop()

    row = selected_df.iloc[0]

    st.markdown('<div class="section-heading">Quotation Preview</div>', unsafe_allow_html=True)
    st.markdown(preview_html(row), unsafe_allow_html=True)

    st.markdown('<div class="section-heading">Optional Customer Details for PDF</div>', unsafe_allow_html=True)

    n1, n2 = st.columns(2)
    with n1:
        customer_name = st.text_input("Name", placeholder="Enter customer name")
    with n2:
        raw_mobile = st.text_input("Mobile", placeholder="Enter 10-digit mobile")
        customer_mobile = clean_mobile(raw_mobile)
        if raw_mobile and raw_mobile != customer_mobile:
            st.caption("Only digits allowed. Maximum 10 digits.")

    pdf_bytes = generate_pdf(row, customer_name.strip(), customer_mobile.strip(), st.session_state.salesperson)

    st.download_button(
        label="Download Quotation PDF",
        data=pdf_bytes,
        file_name=f"Quotation_{tower}_{flat}.pdf",
        mime="application/pdf",
        use_container_width=True
    )
