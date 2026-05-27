import streamlit as st
import pandas as pd
from fpdf import FPDF
import base64
from datetime import datetime, timezone, timedelta

st.set_page_config(page_title="Quotation Generator", layout="wide")

EXCEL_FILE = "master_cost_sheet.xlsx"
USERS_FILE = "users.xlsx"
LOGO_FILE = "logo.png"
BRAND_HEX = "#B3202A"

IST = timezone(timedelta(hours=5, minutes=30))


def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


BRAND_RED = hex_to_rgb(BRAND_HEX)


def indian_format(value):
    try:
        if pd.isna(value):
            return ""
        value = float(value)
        if value.is_integer():
            value = int(value)

        s = str(value)
        whole, decimal = s.split(".") if "." in s else (s, "")
        last_three = whole[-3:]
        other = whole[:-3]

        if other:
            parts = []
            while len(other) > 2:
                parts.insert(0, other[-2:])
                other = other[:-2]
            if other:
                parts.insert(0, other)
            result = ",".join(parts) + "," + last_three
        else:
            result = last_three

        return result + ("." + decimal if decimal else "")
    except:
        return str(value)


def clean_mobile(value):
    return "".join(filter(str.isdigit, value))[:10]


def should_highlight(desc):
    desc = str(desc).upper().strip()
    if "GST" in desc:
        return False
    return desc in ["TOTAL BASE PRICE", "TOTAL PRICE", "TOTAL AMOUNT", "TOTAL AMOUNT WITH GST"]


def section_name(desc):
    desc = str(desc).upper().strip()

    if desc in [
        "TOWER", "FLOOR", "FLAT NO.", "UNIT TYPE",
        "SUPER BUILT UP AREA (SQ.FT.)", "RERA CARPET AREA ( SQ.FT.)"
    ]:
        return "UNIT DETAILS"

    if desc in ["BASE RATE", "PLC", "FLOOR RISE", "TOTAL BASE AMOUNT PER SQ.FT."]:
        return "BASE COST DETAILS"

    if desc in [
        "CAR PARK", "CLUB HOUSE CHARGES",
        "INFRA CHARGES ( INCLUDING POWER, WATER, GENERATOR, STP CHARGES ETC.) PER SQ.FT",
        "LEGAL & DOCUMENTATION CHARGES",
        "ADVANCE MAINTENANCE FOR 1 YEAR",
        "GST FOR LEGAL & DOCUMENTATION CHARGES AND ADVANCE MAINTENANCE CHARGES",
        "SINKING FUND ( INTEREST FREE - PASS THROUGH TO OWNER'S ASSOCIATION )"
    ]:
        return "OTHER CHARGES"

    if desc == "TOTAL PRICE":
        return "FINAL PAYABLE AMOUNT"

    return ""


def add_section_header(pdf, title):
    pdf.set_fill_color(245, 245, 245)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "B", 9)
    pdf.cell(190, 7, title, border=1, ln=True, align="L", fill=True)


def add_wrapped_row(pdf, desc, data, desc_w=145, data_w=45, line_h=5, min_h=7, highlight=False):
    x = pdf.get_x()
    y = pdf.get_y()

    desc_lines = pdf.multi_cell(desc_w - 4, line_h, desc, border=0, split_only=True)
    data_lines = pdf.multi_cell(data_w - 4, line_h, data, border=0, split_only=True)

    row_h = max(min_h, len(desc_lines) * line_h, len(data_lines) * line_h)

    if highlight:
        pdf.set_fill_color(*BRAND_RED)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", "B", 9)
    else:
        pdf.set_fill_color(255, 255, 255)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "", 8.5)

    pdf.rect(x, y, desc_w, row_h, "DF")
    pdf.rect(x + desc_w, y, data_w, row_h, "DF")

    desc_y = y + ((row_h - len(desc_lines) * line_h) / 2)
    data_y = y + ((row_h - len(data_lines) * line_h) / 2)

    pdf.set_xy(x + 2, desc_y)
    pdf.multi_cell(desc_w - 4, line_h, desc, border=0, align="L")

    pdf.set_xy(x + desc_w + 2, data_y)
    pdf.multi_cell(data_w - 4, line_h, data, border=0, align="R")

    pdf.set_xy(x, y + row_h)


def show_center_logo(image_path, width=300):
    with open(image_path, "rb") as img_file:
        encoded = base64.b64encode(img_file.read()).decode()

    st.markdown(
        f"""
        <div class="logo-box">
            <img src="data:image/png;base64,{encoded}" width="{width}">
        </div>
        """,
        unsafe_allow_html=True
    )


st.markdown("""
<style>
.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    max-width: 1200px;
}

.logo-box {
    text-align: center;
    margin-top: 20px;
}

.main-title {
    text-align: center;
    font-size: 44px;
    font-weight: 800;
    color: #2d2f3a;
    margin-top: 20px;
    margin-bottom: 35px;
}

.login-card {
    background: #ffffff;
    padding: 22px 30px 18px 30px;
    border-radius: 22px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.08);
    margin-bottom: 25px;
}

.login-heading {
    text-align: center;
    font-size: 24px;
    font-weight: 800;
    color: #2d2f3a;
    margin-bottom: 8px;
}

.login-subtitle {
    text-align: center;
    font-size: 14px;
    color: #777777;
    margin-bottom: 0px;
}

.welcome-box {
    background: #E8F5E9;
    padding: 12px 18px;
    border-radius: 10px;
    color: #1B5E20;
    font-weight: 600;
    margin-bottom: 20px;
}

.section-card {
    background: #ffffff;
    padding: 30px;
    border-radius: 20px;
    box-shadow: 0 8px 28px rgba(0,0,0,0.08);
    margin-top: 20px;
    margin-bottom: 30px;
}

.section-heading {
    font-size: 28px;
    font-weight: 800;
    color: #2d2f3a;
    margin-bottom: 20px;
}

.stButton > button, .stDownloadButton > button {
    background-color: #B3202A;
    color: white;
    border-radius: 12px;
    height: 52px;
    font-size: 18px;
    font-weight: 700;
    border: none;
}

.stButton > button:hover, .stDownloadButton > button:hover {
    background-color: #931923;
    color: white;
}

div[data-testid="stTextInput"] {
    margin-top: 10px;
}
</style>
""", unsafe_allow_html=True)


if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "salesperson_name" not in st.session_state:
    st.session_state.salesperson_name = ""
if "username" not in st.session_state:
    st.session_state.username = ""
if "quotation_generated" not in st.session_state:
    st.session_state.quotation_generated = False
if "selected_data" not in st.session_state:
    st.session_state.selected_data = None
if "quotation_time" not in st.session_state:
    st.session_state.quotation_time = None


if not st.session_state.logged_in:
    try:
        users_df = pd.read_excel(USERS_FILE)
        users_df.columns = users_df.columns.str.strip()
    except:
        st.error("users.xlsx file not found or not readable.")
        st.stop()

    show_center_logo(LOGO_FILE, width=250)
    st.markdown('<div class="main-title">Flat Quotation Generator</div>', unsafe_allow_html=True)

    left, middle, right = st.columns([1.25, 1, 1.25])

    with middle:
        st.markdown(
            """
            <div class="login-card">
                <div class="login-heading">Sales Login</div>
                <div class="login-subtitle">Enter your credentials to continue</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        username_input = st.text_input("Username", placeholder="Enter username")
        password_input = st.text_input("Password", type="password", placeholder="Enter password")

        if st.button("Login", use_container_width=True):
            match = users_df[
                (users_df["USERNAME"].astype(str).str.strip() == username_input.strip()) &
                (users_df["PASSWORD"].astype(str).str.strip() == password_input.strip())
            ]

            if not match.empty:
                st.session_state.logged_in = True
                st.session_state.username = username_input.strip()
                st.session_state.salesperson_name = str(match.iloc[0]["SALES PERSON NAME"]).strip()
                st.rerun()
            else:
                st.error("Invalid username or password.")

    st.stop()


df = pd.read_excel(EXCEL_FILE)
df.columns = df.columns.str.strip()

show_center_logo(LOGO_FILE, width=300)
st.markdown('<div class="main-title">Flat Quotation Generator</div>', unsafe_allow_html=True)

header_left, header_right = st.columns([8, 1])

with header_left:
    st.markdown(
        f"""
        <div class="welcome-box">
            Welcome, {st.session_state.salesperson_name}
        </div>
        """,
        unsafe_allow_html=True
    )

with header_right:
    if st.button("Logout", key="logout_btn"):
        st.session_state.logged_in = False
        st.session_state.salesperson_name = ""
        st.session_state.username = ""
        st.session_state.quotation_generated = False
        st.session_state.selected_data = None
        st.session_state.quotation_time = None
        st.rerun()


st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-heading">Select Flat Details</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)

with c1:
    tower = st.selectbox(
        "Select Tower",
        sorted(df["TOWER"].dropna().unique()),
        disabled=st.session_state.quotation_generated
    )

tower_df = df[df["TOWER"] == tower]

with c2:
    floor = st.selectbox(
        "Select Floor",
        sorted(tower_df["FLOOR"].dropna().unique()),
        disabled=st.session_state.quotation_generated
    )

floor_df = tower_df[tower_df["FLOOR"] == floor]

with c3:
    flat = st.selectbox(
        "Select Flat No.",
        sorted(floor_df["FLAT NO."].dropna().unique()),
        disabled=st.session_state.quotation_generated
    )

if not st.session_state.quotation_generated:
    if st.button("Generate Quotation", use_container_width=True):
        selected = floor_df[floor_df["FLAT NO."] == flat].iloc[0]
        st.session_state.selected_data = selected
        st.session_state.quotation_generated = True
        st.session_state.quotation_time = datetime.now(IST)
        st.rerun()
else:
    if st.button("Clear All Data / New Quotation", use_container_width=True):
        st.session_state.quotation_generated = False
        st.session_state.selected_data = None
        st.session_state.quotation_time = None
        st.rerun()

st.markdown('</div>', unsafe_allow_html=True)


if st.session_state.quotation_generated and st.session_state.selected_data is not None:
    selected = st.session_state.selected_data
    quotation_date = st.session_state.quotation_time.strftime("%d-%m-%Y")
    quotation_datetime = st.session_state.quotation_time.strftime("%d-%m-%Y %I:%M %p")

    quotation_df = pd.DataFrame({
        "Description": selected.index,
        "Data": selected.values
    })

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-heading">Quotation Preview</div>', unsafe_allow_html=True)

    st.caption(f"Quotation Date: {quotation_date}")
    st.caption(f"Generated By: {st.session_state.salesperson_name}")

    preview_df = quotation_df.copy()
    preview_df["Data"] = preview_df["Data"].apply(indian_format)
    st.table(preview_df)

    st.markdown("### Customer Details for PDF Optional")

    c4, c5 = st.columns(2)

    with c4:
        customer_name = st.text_input("Name")

    with c5:
        customer_mobile_input = st.text_input("Mobile", max_chars=10)

    customer_mobile = clean_mobile(customer_mobile_input)

    if customer_mobile_input != customer_mobile:
        st.warning("Mobile number should contain digits only and maximum 10 digits.")

    pdf = FPDF()
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()

    try:
        pdf.image(LOGO_FILE, x=75, y=8, w=60)
    except:
        pass

    pdf.ln(28)

    pdf.set_fill_color(*BRAND_RED)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 13)
    pdf.cell(190, 9, "COST SHEET & PAYMENT PLAN", border=1, ln=True, align="C", fill=True)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "", 10)
    pdf.cell(190, 7, f"Date: {quotation_date}", border=1, ln=True, align="R")

    if customer_name.strip():
        pdf.cell(190, 7, f"Name: {customer_name.strip()}", border=1, ln=True, align="L")

    if customer_mobile.strip():
        pdf.cell(190, 7, f"Mobile: {customer_mobile.strip()}", border=1, ln=True, align="L")

    pdf.ln(3)

    last_section = ""

    for _, row in quotation_df.iterrows():
        desc = str(row["Description"]).strip()
        data = indian_format(row["Data"])

        current_section = section_name(desc)

        if current_section and current_section != last_section:
            add_section_header(pdf, current_section)
            last_section = current_section

        add_wrapped_row(pdf, desc, data, highlight=should_highlight(desc))

    note_y = max(pdf.get_y() + 6, 250)
    if note_y > 258:
        note_y = pdf.get_y() + 5

    pdf.set_y(note_y)

    pdf.set_fill_color(*BRAND_RED)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 10)
    pdf.cell(190, 8, "Important Note", border=1, ln=True, fill=True)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "I", 8)

    note_text = "This quotation is indicative and subject to final approval. Prices, GST and other charges are subject to change as applicable."
    pdf.multi_cell(190, 6, note_text, border=1)

    pdf.set_y(282)
    pdf.set_font("Arial", "I", 7)
    pdf.set_text_color(90, 90, 90)

    pdf.cell(95, 8, f"Generated on: {quotation_datetime}", align="L")
    pdf.cell(95, 8, f"Generated By: {st.session_state.salesperson_name}", align="R")

    pdf_output = bytes(pdf.output(dest="S"))

    st.download_button(
        label="Download Quotation PDF",
        data=pdf_output,
        file_name=f"Quotation_{selected['TOWER']}_{selected['FLAT NO.']}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

    st.markdown('</div>', unsafe_allow_html=True)
