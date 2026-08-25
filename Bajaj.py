import streamlit as st
import pandas as pd
import re
from io import BytesIO
import pdfplumber
from datetime import datetime

# --- Streamlit Config ---

st.set_page_config(page_title="PDF to Excel - Bajaj Policy Extractor", layout="centered")
st.title("📄 Bajaj PDF Policy Extractor → Excel")
st.write("Upload Bajaj General Insurance policy PDFs to extract key details into Excel.")

# --- Sidebar for Direct Accessory Value Input ---

st.sidebar.header("🔧 Direct Accessory Adjustment")
accessory_value = st.sidebar.number_input(
    "Total Value of Non-Electronic Accessories (₹)",
    min_value=0.0,
    step=100.0,
    value=0.0,
    help="Enter the total charges/values for non-electronic accessories (e.g., roof racks, mats). This will be directly added to the Sum Insured / IDV for all extracted policies."
)

# --- File Upload ---

uploaded_files = st.file_uploader("Upload Policy PDFs", type=["pdf"], accept_multiple_files=True)

# --- Output Columns ---

columns = [
    "Customer Id", "Customer Name", "Policy No", "Effective Date", "Expiry Date",
    "Product Name", "Sum Insured / IDV", "Premium Paid (Incl. GST)", "Intermediary Name",
    "Customer Number", "cust_email", "Fuel Type", "Vehicle No / Registration Number",
    "CHASSIS NUM", "ENGINE NUM", "VEHICLE INFO", "Payment Mode", "File Name"
]

# --- Helper Function ---

def find(pattern, text, flags=re.IGNORECASE | re.DOTALL):
    match = re.search(pattern, text, flags)
    if not match:
        return "N/A"
    return match.group(1).strip() if match.lastindex else match.group(0).strip()

# --- Date Formatter ---

def format_date(date_str):
    try:
        clean_date = date_str.split()[0]
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d %b %Y", "%d %b '%y"):
            try:
                d = datetime.strptime(clean_date, fmt)
                return d.strftime("%d %b '%y")
            except:
                continue
        return date_str
    except:
        return date_str

# --- Extraction Function Tailored for Bajaj ---

def extract_policy_details(text, file_name):
    t = re.sub(r'\s+', ' ', text.replace("\n", " "))

    # --- Policy Number ---
    policy_no = find(r"(?:Policy\s*(?:Number|No\.?)|policy\s*number)\s*[:\-'\s]*([0-9]{2}\-[0-9]{4}\-[0-9]{10}\-[0-9]{2}|OG\-\d{2}\-\d{4}\-\d{4}\-\d{8}|\d{6,16})", t)

    # --- Effective / Expiry Date ---
    eff_date = exp_date = "N/A"
    date_match = re.search(
        r"From[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}|\d{1,2}\s+[A-Za-z]{3}\s+'?\d{2,4})(?:.*?\bTo[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}|\d{1,2}\s+[A-Za-z]{3}\s+'?\d{2,4}|Midnight))?",
        t, re.IGNORECASE
    )
    if date_match:
        eff_date = format_date(date_match.group(1))
        if date_match.group(2) and date_match.group(2).lower() != "midnight":
            exp_date = format_date(date_match.group(2))

    # Fallback date extraction from text blocks
    if eff_date == "N/A":
        all_dates = re.findall(r"\b(\d{2}[\/\-]\d{2}[\/\-]\d{4})\b", t)
        if len(all_dates) >= 2:
            eff_date, exp_date = format_date(all_dates[0]), format_date(all_dates[1])
        elif len(all_dates) == 1:
            eff_date = format_date(all_dates[0])

    # --- Customer ID ---
    cust_id = find(r"Customer\s*ID\s*[:\-]?\s*([0-9A-Z]+)", t)

    # --- Customer Name ---
    cust_name = find(r"(?:Insured\s*Name|Name\s*\(Proposer\)|Received\s*with\s*thanks\s*from)\s*[:\-]?\s*([A-Za-z\s]+?)(?=\s*(?:Name|Address|Customer|Policy|GSTIN|PAN|a\s*total|$))", t)
    if cust_name == "N/A":
        cust_name = find(r"Dear\s+([A-Za-z\s]+?),", t)
    cust_name = cust_name.strip().title() if cust_name != "N/A" else "N/A"

     # ============================================================
# --- PRODUCT NAME ---
# ============================================================

    product = "N/A"

# 1. Product Name
    product = find(
    r"(?:Product\s*Name|Product)\s*[:\-]?\s*"
    r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
    r"(?=\s*(?:UIN|Policy|Customer|Insured|Vehicle|$))",
    t,
    re.IGNORECASE
)

# 2. OR - Policy Type
    if product == "N/A":
     product = find(
        r"Policy\s*Type\s*[:\-]?\s*"
        r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
        r"(?=\s*(?:UIN|Policy|Customer|Insured|Vehicle|$))",
        t,
        re.IGNORECASE
    )

# 3. OR - Plan Name
    if product == "N/A":
       product = find(
        r"Plan\s*Name\s*[:\-]?\s*"
        r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
        r"(?=\s*(?:UIN|Policy|Customer|Insured|Vehicle|$))",
        t,
        re.IGNORECASE
    )

# 4. OR - Product / Plan
    if product == "N/A":
       product = find(
        r"Product\s*\/\s*Plan\s*[:\-]?\s*"
        r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
        r"(?=\s*(?:UIN|Policy|Customer|Insured|Vehicle|$))",
        t,
        re.IGNORECASE
    )

# 5. OR - Type of Policy
    if product == "N/A":
       product = find(
        r"Type\s*of\s*Policy\s*[:\-]?\s*"
        r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
        r"(?=\s*(?:UIN|Policy|Customer|Insured|Vehicle|$))",
        t,
        re.IGNORECASE
    )

# 6. OR - TRANSCRIPT OF PROPOSAL FOR
    if product == "N/A":
       product = find(
        r"TRANSCRIPT\s*OF\s*PROPOSAL\s*FOR\s*"
        r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
        r"(?=\s*(?:UIN|Customer|Insured|Vehicle|$))",
        t,
        re.IGNORECASE
    )
    
    if product == "N/A":
       product = find(
        r"([A-Za-z0-9\s\-\&\/\(\)]+?)\s*-\s*POLICY\s*SCHEDULE",
        t,
        re.IGNORECASE
    )

# 7. Known policy wording fallback
    if product == "N/A":

        if "liability only" in t.lower():
           product = "Liability Only Policy For Private Car"

        elif "package policy" in t.lower():
           product = "Private Car Package Policy"
        elif "policy" in t.lower():
                      product = "two wheeler Policy"
        elif "comprehensive" in t.lower():
           product = "Comprehensive Motor Insurance Policy"

        elif "third party" in t.lower():
           product = "Third Party Motor Insurance Policy"

# 8. Final fallback
        if product == "N/A":
           product = "Motor Insurance Policy"

# Clean result
    product = re.sub(r"\s+", " ", product).strip()


    # --- Premium ---
    premium = find(r"(?:Final\s*Premium|Total\s*Amount|Gross\s*Premium)\s*[:\-]?\s*Rs\.?\s*([\d,\.]+)", t)
    if premium == "N/A":
        premium = find(r"Final\s*Premium\s*([\d,]+)", t)
    if premium == "N/A":
        premium = find(r"Total\s*Amount\s*[:\-]?\s*([\d,\.]+)", t)
    if premium != "N/A":
        try:
            premium = f"{float(premium.replace(',', '')):,.0f}"
        except:
            pass

    # ============================================================
# --- INTERMEDIARY NAME ---
# ============================================================

    intermediary = "N/A"

    intermediary = find(
    r"(?:Intermediary\s*Name|Intermediary|Agency\s*Name|"
    r"Agent\s*Name|Broker\s*Name)"
    r"\s*[:\|]?\s*"
    r"([A-Za-z0-9\s\.\&]+?)"
    r"(?=\s*(?:Email|Contact|Mobile|Phone|Sub|SP|$))",
    t,
    re.IGNORECASE
)
    if intermediary != "N/A":
       intermediary = intermediary.strip().title()

    # --- Mobile (Customer Number) ---
    mobile = find(r"Mobile\s*(?:Number|No\.?)\s*[:\|]?\s*([6-9]\d{9})", t)
    if mobile == "N/A":
        mobile = find(r"\b[6-9]\d{9}\b", t)

    # --- Customer Email ---
    email = find(r"Email\s*(?:ID)?\s*[:\|]?\s*([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})", t)
    if email == "N/A":
        all_emails = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", t, re.IGNORECASE)
        email = next((e for e in all_emails if not re.search(r"bajaj|care|support|info|admin", e, re.IGNORECASE)), "N/A")

       # ============================================================
    # --- FUEL TYPE ---
    # ============================================================

    fuel = "N/A"

    # 1. Direct Fuel Type label
    fuel_match = re.search(
        r"Fuel\s*Type\s*[:\|]?\s*"
        r"(PETROL(?:\s*\([A-Z]+\))?|DIESEL(?:\s*\([A-Z]+\))?|"
        r"CNG(?:\s*\([A-Z]+\))?|LPG(?:\s*\([A-Z]+\))?|"
        r"ELECTRIC|HYBRID)",
        t,
        re.IGNORECASE
    )

    if fuel_match:
        fuel = fuel_match.group(1).strip().upper()

    # 2. Bajaj vehicle-table fallback
    # Example:
    # GJ-33-F-0282 MARUTI DZIRE VXI 1197 61 PETROL(P)
    if fuel == "N/A":
        fuel_match = re.search(
            r"\b(PETROL\s*\([A-Z]+\)|"
            r"DIESEL\s*\([A-Z]+\)|"
            r"CNG\s*\([A-Z]+\)|"
            r"LPG\s*\([A-Z]+\)|"
            r"ELECTRIC|HYBRID)\b",
            t,
            re.IGNORECASE
        )

        if fuel_match:
            fuel = fuel_match.group(1).strip().upper()

    # 3. Final fallback
    if fuel == "N/A":
        fuel_match = re.search(
            r"\b(PETROL|DIESEL|CNG|LPG|ELECTRIC|HYBRID)\b",
            t,
            re.IGNORECASE
        )

        if fuel_match:
            fuel = fuel_match.group(1).strip().upper()
        # ============================================================
    # --- Registration Number (Vehicle No) ---
    # ============================================================

    reg_no = find(
        r"Registration\s*Number\s*[:\|]?\s*"
        r"([A-Z]{2}[\-\s]?\d{2}[\-\s]?[A-Z]{1,2}[\-\s]?\d{4})",
        t
    )

    reg_no = find(
    r"(?:Registration\s*Number|Registration\s*No\.?|"
    r"Regn\s*Number|Regn\s*No\.?|Vehicle\s*Number|Vehicle\s*No\.?)"
    r"\s*[:\|]?\s*"
    r"([A-Z]{2}[\-\s]?\d{2}[\-\s]?[A-Z]{1,3}[\-\s]?\d{4})",
    t
)

    # Fallback: find registration number anywhere in the PDF
    if reg_no == "N/A":
        reg_no = find(
            r"\b([A-Z]{2}[\-\s]?\d{2}[\-\s]?[A-Z]{1,2}[\-\s]?\d{4})\b",
            t
        )

    # Remove spaces/hyphens if required
    if reg_no != "N/A":
        reg_no = re.sub(r"\s+", "-", reg_no)
   # ============================================================
# --- ENGINE NUMBER ---
# ============================================================

    engine = "N/A"
    match_eng = None

# 1. Direct Engine Number label
    engine = find(
    r"(?:Engine\s*Number|Engine\s*No\.?|Engine\s*#|Engine\s*Number)"
    r"\s*[:\|]?\s*([A-Z0-9]{8,25})",
    t
)

# 2. Bajaj vehicle schedule fallback
    if engine == "N/A":
       match_eng = re.search(
        r"\b([A-HJ-NPR-Z0-9]{17})\s+"
        r"([A-Z0-9]{8,25})\s+"
        r"(?:0|[\d,]+(?:\.\d+)?)\b",
        t,
        re.IGNORECASE
    )

    if match_eng:
        engine = match_eng.group(2).strip().upper()

# 3. Final fallback - engine after chassis
    if engine == "N/A":
        match_eng = re.search(
        r"(?:Chassis\s*Number|Chassis\s*No\.?|Chassis\s*#)"
        r".{0,300}?"
        r"\b([A-HJ-NPR-Z0-9]{17})\b"
        r"\s+"
        r"([A-Z0-9]{8,25})\b",
        t,
        re.IGNORECASE | re.DOTALL
    )

    if match_eng:
        engine = match_eng.group(2).strip().upper()


# ============================================================
# --- CHASSIS NUMBER ---
# ============================================================

    chassis = "N/A"
    match_chas = None

# 1. Direct Chassis Number label
    chassis = find(
    r"(?:Chassis\s*Number|Chassis\s*No\.?|Chassis\s*#)"
    r"\s*[:\|]?\s*([A-Z0-9]{11,25})",
    t
)

# 2. Standard 17-character VIN / Chassis fallback
    if chassis == "N/A":
       match_chas = re.search(
        r"\b([A-HJ-NPR-Z0-9]{17})\b",
        t,
        re.IGNORECASE
    )

    if match_chas:
        chassis = match_chas.group(1).strip().upper()

# 3. Bajaj chassis + engine sequence fallback
    if chassis == "N/A":
       match_chas = re.search(
        r"\b([A-HJ-NPR-Z0-9]{17})\b"
        r"\s+"
        r"[A-Z0-9]{8,25}"
        r"\s+"
        r"(?:0|[\d,]+(?:\.\d+)?)\b",
        t,
        re.IGNORECASE
    )

    if match_chas:
        chassis = match_chas.group(1).strip().upper()
    # ============================================================
    # --- VEHICLE INFO ---
    # ============================================================

    vehicle_info = "N/A"

    # Bajaj structure:
    # GJ-33-F-0282 MARUTI DZIRE VXI 1197 61 PETROL(P)

    vehicle_match = re.search(
        r"\b[A-Z]{2}[\-\s]?\d{2}[\-\s]?[A-Z]{1,2}[\-\s]?\d{4}\s+"
        r"(.+?)\s+"
        r"\d{2,5}\s+\d+\s+"
        r"(?:PETROL|DIESEL|CNG|ELECTRIC|HYBRID)",
        t,
        re.IGNORECASE
    )

    if vehicle_match:
        vehicle_info = vehicle_match.group(1).strip()
 
    # Clean vehicle info
    if vehicle_info != "N/A":
        vehicle_info = re.sub(r"\s+", " ", vehicle_info)
    # --- Payment Mode ---
    if "online payment" in t.lower():
        pay_mode = "Online Payment"
    elif "cheque" in t.lower():
        pay_mode = "Cheque"
    else:
        pay_mode = find(r"Instrument\s*Type\s*[:\|]?\s*([A-Za-z\s]+)", t)

    return {
        "Customer Id": cust_id,
        "Customer Name": cust_name,
        "Policy No": policy_no,
        "Effective Date": eff_date,
        "Expiry Date": exp_date,
        "Product Name": product,
        "Sum Insured / IDV": idv,
        "Premium Paid (Incl. GST)": premium,
        "Intermediary Name": intermediary,
        "Customer Number": mobile,
        "cust_email": email,
        "Fuel Type": fuel,
        "Vehicle No / Registration Number": reg_no,
        "CHASSIS NUM": chassis,
        "ENGINE NUM": engine,
        "VEHICLE INFO": vehicle_info,
        "Payment Mode": pay_mode
    }

# --- Main Processing ---

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        with pdfplumber.open(file) as pdf:
            text = " ".join(page.extract_text() or "" for page in pdf.pages)

        data = extract_policy_details(text, file.name)
        data["File Name"] = file.name
        all_data.append(data)

    df = pd.DataFrame(all_data, columns=columns).fillna("N/A")

    # --- Add accessory value directly to IDV ---
    if accessory_value > 0:
        def update_idv(idv_str):
            if idv_str == "N/A":
                return f"{accessory_value:,.2f}"
            try:
                base_value = float(idv_str.replace(',', ''))
                updated_value = base_value + accessory_value
                return f"{updated_value:,.2f}"
            except:
                return idv_str

        df["Sum Insured / IDV"] = df["Sum Insured / IDV"].apply(update_idv)
        st.sidebar.success(f"✅ Sum Insured updated by ₹{accessory_value:,.2f} for accessories!")

    st.success("✅ Extraction complete! Review below:")
    st.dataframe(df)

    # --- Download Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Policy Details")

    st.download_button(
        label="📥 Download Extracted Policy Data (Excel)",
        data=output.getvalue(),
        file_name="bajaj_policy_extracted_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")
st.caption("Built with 💙 Streamlit + pdfplumber + Regex | Tailored specifically for Bajaj General Insurance Policy PDFs")