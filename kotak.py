import streamlit as st
import pandas as pd
import re
from io import BytesIO
from PyPDF2 import PdfReader
from datetime import datetime

# --- Streamlit Config ---
st.set_page_config(page_title="PDF to Excel - Policy Extractor", layout="centered")
st.title("📄 PDF Policy Extractor → Excel")
st.write("Upload one or more insurance policy PDFs (Tata AIG, Zurich Kotak, Royal Sundaram, ICICI Lombard, Reliance, etc.) to extract key details into Excel.")

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
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%d %b '%y"):
            try:
                d = datetime.strptime(date_str, fmt)
                return d.strftime("%d %b '%y")
            except:
                continue
        return date_str
    except:
        return date_str

# --- Extraction Function ---
def extract_policy_details(text, file_name):
    t = re.sub(r'\s+', ' ', text.replace("\n", " "))

    # --- Policy Number ---
    policy_no = find(r"Policy\s*(?:No\.?|Number)\s*[:\-]?\s*(\d{6,15})", t)

    # --- Effective / Expiry Date ---
    eff_date = exp_date = "N/A"
    match = re.search(
        r"(?:Period\s*of\s*Insurance|Policy\s*Period).*?(?:From|Valid\s*from)[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}).*?(?:To|Till|Up\s*to)[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})",
        t, re.IGNORECASE)
    if match:
        eff_date, exp_date = format_date(match.group(1)), format_date(match.group(2))

    # --- Customer ID ---
    cust_id = find(r"Customer\s*ID\s*[:\-]?\s*([0-9A-Z]+)", t)

    # --- Customer Name ---
    cust_name = find(r"Name\s*[:\-]?\s*([A-Za-z\s\.]+Agarwal)", t)
    if cust_name == "N/A":
        cust_name = find(r"(?:Insured|Customer)\s*Name\s*[:\-]?\s*([A-Za-z\s\.]+)", t)
    cust_name = cust_name.strip().title()

    # --- Product Name ---
    if "car secure" in t.lower():
        product = "Private Car Package Policy (Car Secure)"
    elif "private car" in t.lower():
        product = "Private Car Package Policy"
    else:
        product = find(r"(?:Product\s*Name|Policy\s*Type|Cover\s*Type)\s*[:\-]?\s*([A-Za-z\s]+)", t)

    # --- IDV / Sum Insured ---
    idv = find(r"Total\s*Value\s*of\s*the\s*Vehicle[^\d]*([\d,]+)", t)
    if idv != "N/A":
        try:
            idv = f"{float(idv.replace(',', '')):,.2f}"
        except:
            pass

    # --- Premium ---
    premium = find(r"Total\s*Premium\s*\(in\s*₹\s*\)[^\d]*([\d,]+)", t)
    if premium == "N/A":
        premium = find(r"(?:Total\s*Premium|Premium\s*Amount|Premium\s*Paid)[^\d]*([\d,\.]+)", t)
    if premium != "N/A":
        try:
            premium = f"{float(premium.replace(',', '')):,.0f}"
        except:
            pass

    # --- Intermediary Name ---
    intermediary = find(r"Intermediary\s*Name\s*([A-Za-z\s\.]+)", t)
    intermediary = re.sub(r"\bIntermediary\b", "", intermediary).strip().title()

    # --- Mobile (Customer Number) ---
    mobile = find(r"(?:Mobile|Phone|Contact\s*No\.?)\s*[:\-]?\s*([6-9]\d{9})", t)
    if mobile == "N/A":
        match = re.search(r"(?:Mobile|Phone|Contact)\s*[:\-]?\s*(\d{2,3}X+\d{2,3})", t, re.IGNORECASE)
        if match:
            mobile = match.group(1)
    if mobile == "N/A":
        mobile = find(r"\b[6-9]\d{9}\b", t)
    
    # --- INSURED DETAILS BLOCK ---
    insured_block = find(r"INSURED\s*DETAILS(.*?)(?:POLICY\s*DETAILS|INTERMEDIARY\s*DETAILS|VEHICLE\s*DETAILS)", t)
     # --- Customer Email (from INSURED DETAILS first) ---
    email = find(r"(?:Email(?:\s*ID)?|E[\-\s]?mail)\s*[:\-]?\s*([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})", insured_block)
    if email == "N/A":
        # try global search but skip company emails
        all_emails = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", t, re.IGNORECASE)
        email = next((e for e in all_emails if not re.search(r"zurichkotak|care|support|info|admin", e, re.IGNORECASE)), "N/A")

    # --- Fuel Type ---
    fuel = find(r"\b(PETROL|DIESEL|CNG|ELECTRIC|HYBRID)\b", t)

    # --- Registration Number (Vehicle No) ---
    reg_no = find(r"(?:Registration\s*No\.?|Vehicle\s*No\.?|Regn\s*No\.?|Registration\s*Number)\s*[:\-]?\s*([A-Z]{2}[\s\-]?\d{2}[\s\-]?[A-Z]{1,2}[\s\-]?\d{4})", t)
    if reg_no == "N/A":
        reg_no = find(r"([A-Z]{2}[\s\-]?\d{2}[\s\-]?[A-Z]{1,2}[\s\-]?\d{4})", t)

    # --- Chassis (Vehicle Chassis No.) ---
    chassis = find(r"Vehicle\s*Chassis\s*(?:No\.?)?\s*[:\-]?\s*([A-Z0-9\s]{6,20})", t)
    if chassis == "N/A":
        match = re.search(
            r"HONDA[\/]?\s*CITY.*?(\d{4})\s+[A-Z]+\s+[A-Z0-9]+\s+\d+([A-Z0-9\s]{6,20})\s+([A-Z0-9\s]{8,20})",
            t, re.IGNORECASE)
        if match:
            chassis = match.group(2).strip()  # Chassis is 2nd group

    # --- Engine (Engine Number) ---
    engine = find(r"Engine\s*No\.?\s*([A-Z0-9\s]{8,20})", t)
    if engine == "N/A":
        match = re.search(
            r"(\d{4})\s+[A-Z]+\s+[A-Z0-9]+\s+\d+([A-Z0-9\s]{6,20})\s+([A-Z0-9]{6,20})(?:\s+(?:PETROL|DIESEL|CNG|ELECTRIC))?",
            t, re.IGNORECASE
        )
        if match:
            engine = match.group(3).strip()

    # --- Vehicle Info ---
    vehicle_info = find(r"(HONDA[\/]?\s*CITY\s+[A-Za-z0-9\s\-\(\)\.]+?)(?=\d{4}|\s+[A-Z]{2,}|\s+Insured|$)", t)
    if vehicle_info == "N/A":
        vehicle_info = find(r"(?:Make\s*\/\s*Model|Manufacturer\s*Model)\s*[:\-]?\s*([A-Za-z0-9\s\-\(\)\/]+)", t)
    vehicle_info = vehicle_info.strip()

    # --- Payment Mode ---
    if "payment aggregator" in t.lower():
        pay_mode = "PAYMENT AGGREGATOR"
    elif "online" in t.lower():
        pay_mode = "Online Payment"
    else:
        pay_mode = find(r"(?:Payment\s*Mode|Mode\s*of\s*Payment)\s*[:\-]?\s*([A-Za-z\s]+)", t)

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
        reader = PdfReader(file)
        text = " ".join(page.extract_text() or "" for page in reader.pages)
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
        file_name="policy_extracted_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")
st.caption("Built with 💙 Streamlit + PyPDF2 + Regex | Supports Tata AIG, Zurich Kotak, Royal Sundaram, ICICI Lombard, Reliance & more | Fixed Zurich Kotak Engine/Chassis mapping")
