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

# --- File Upload ---
uploaded_files = st.file_uploader("Upload Policy PDFs", type=["pdf"], accept_multiple_files=True)

# --- Output Columns ---
columns = [
    "Customer Id", "Customer Name", "Policy No", "Effective Date", "Expiry Date",
    "Product Name", "Sum Insured / IDV", "Premium Paid (Incl. GST)", "Intermediary Name",
    "CUST_MOBILE_NUMBER", "CUST_EMAIL", "Fuel Type", "Vehicle No / Registration Number",
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
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%d %b '%y", "%d.%m.%Y", "%d-%b-%Y", "%d %B %Y", "%d-%B-%Y"):
            try:
                d = datetime.strptime(date_str.strip(), fmt)
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
        r"Period\s*of\s*Insurance\s*[:\-]?\s*From\s*\d{1,2}:\d{2}\s*Hrs\s*on\s*(\d{1,2}[-/\s]?[A-Za-z]{3,9}[-/\s]?\d{2,4})\s*to\s*Midnight\s*of\s*(\d{1,2}[-/\s]?[A-Za-z]{3,9}[-/\s]?\d{2,4})",
        t, re.IGNORECASE
    )
    if match:
        eff_date, exp_date = format_date(match.group(1)), format_date(match.group(2))
    else:
        match = re.search(
            r"Period\s*of\s*Insurance\s*[:\-]?\s*From[^\d]*(\d{1,2}[-/\s]?[A-Za-z]{3,9}[-/\s]?\d{2,4}).*?to[^\d]*(\d{1,2}[-/\s]?[A-Za-z]{3,9}[-/\s]?\d{2,4})",
            t, re.IGNORECASE
        )
        if match:
            eff_date, exp_date = format_date(match.group(1)), format_date(match.group(2))
        else:
            match = re.search(
                r"Period\s*of\s*Insurance\s*[:\-]?\s*(\d{1,2}[-/\s]?[A-Za-z]{3,9}[-/\s]?\d{2,4})\s*(?:to|-)\s*(\d{1,2}[-/\s]?[A-Za-z]{3,9}[-/\s]?\d{2,4})",
                t, re.IGNORECASE
            )
            if match:
                eff_date, exp_date = format_date(match.group(1)), format_date(match.group(2))

    # --- Customer ID ---
    cust_id = find(r"Customer\s*ID\s*[:\-]?\s*([0-9A-Z]+)", t)

    # --- Customer Name ---
    cust_name = find(r"Insured\s*Name\s*[:\-]?\s*((?:Mr\.?|Mrs\.?|Ms\.?|M/s\.?)\s*[A-Za-z\s\.]+?)(?=\s*Period|\s*Policy|\s*$)", t)
    if cust_name == "N/A":
        cust_name = find(r"(?:Customer|Policy\s*Holder)\s*Name\s*[:\-]?\s*((?:Mr\.?|Mrs\.?|Ms\.?|M/s\.?)\s*[A-Za-z\s\.]+)", t)
    cust_name = cust_name.strip().title()

    # --- Product Name ---
    product = find(r"Reliance\s+[A-Za-z0-9\s\-\(\)]+\s*Package\s*Policy\s*-\s*Policy\s*Schedule", t)
    if product == "N/A":
        product = find(r"(?:Product\s*Name|Policy\s*Type|Cover\s*Type|Plan\s*Name|Policy\s*Schedule)\s*[:\-]?\s*([A-Za-z0-9\s\-\(\)]+?)(?=\s*Sum|\s*Premium|\s*Intermediary|$)", t)
        if product == "N/A":
            if "car secure" in t.lower():
                product = "Private Car Package Policy (Car Secure)"
            elif "private car" in t.lower():
                product = "Private Car Package Policy"
            else:
                product = "Policy Schedule"

    # --- Financial Details ---
    idv = find(r"(?:IDV|Sum\s*Insured|Liability\s*Limit)[^\d]*([\d,.]+)", t)
    premium = find(
        r"(?:Total\s*Premium|Premium\s*Paid|Gross\s*Premium|Total\s*Amount\s*Payable|Net\s*Premium\s*\+?\s*GST)[^\d]*([\d,\.]+)",
        t
    )

    # --- Intermediary Name ---
    intermediary = find(r"Intermediary\s*Name\s*[:\-]?\s*([A-Za-z\s\.]+)", t)
    if intermediary == "N/A":
        intermediary = find(r"Agent\s*Name\s*[:\-]?\s*([A-Za-z\s\.]+)", t)
    intermediary = re.sub(r"\bIntermediary\b|\s*Code\s*$", "", intermediary).strip().title()

    # --- Customer Mobile ---
    mobile = find(r"(?:Mobile\s*No\.?|Customer\s*contact\s*number)\s*[:\-]?\s*([\d\*\s]+)", t)
    if mobile == "N/A":
        mobile = find(r"\b[6-9]\d{9}\b", t)
    if mobile != "N/A":
        mobile = mobile.replace(" ", "").replace("*", "X")

    # --- Customer Email ---
    email = find(r"Email[\s\-]*ID\s*[:\-]?\s*([A-Za-z0-9._%+\-*]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}|NA)", t)
    if email.upper() == "NA" or not re.search(r"@", email):
        email = "N/A"

    # --- Fuel Type ---
    fuel = find(r"Fuel\s*Type\s*[:\-]?\s*([A-Za-z]+)", t)
    if fuel == "N/A":
        fuel = find(r"(PETROL|DIESEL|CNG|ELECTRIC|HYBRID)", t)

    # --- Vehicle Info (Make / Model / Modal / Variant / Make and Model) ---
    vehicle_info = find(
        r"(?:Make\s*(?:\/|and)\s*(?:Model|Modal)\s*&?\s*Variant)\s*[:\-]?\s*([A-Za-z0-9\s\-\(\)\/]+?)(?=\s*Engine|\s*Chassis|$)",
        t
    )
    if vehicle_info == "N/A":
        vehicle_info = find(
            r"(?:Make\s*(?:\/|and)\s*(?:Model|Modal))\s*[:\-]?\s*([A-Za-z0-9\s\-\(\)\/]+?)(?=\s*Engine|\s*Chassis|$)",
            t
        )
    if vehicle_info == "N/A":
        vehicle_info = find(
            r"(?:Vehicle\s*Description|Model\s*Details)\s*[:\-]?\s*([A-Za-z0-9\s\-\(\)\/]+)",
            t
        )
    vehicle_info = vehicle_info.strip()

    # --- Registration Number ---
    reg_no = find(
        r"(?:Registration\s*No\.?|Vehicle\s*No\.?|Regn\s*No\.?|Registration\s*Number)\s*[:\-]?\s*([A-Z]{2}\s*\d{2}\s*[A-Z]{1,2}\s*\d{4})",
        t
    )

    # --- Engine / Chassis ---
    combined_ec = find(r"Engine\s*No\.?\s*\/\s*Chassis\s*No\.?\s*[:\-]?\s*([A-Z0-9\s\/\-]+)", t)
    if combined_ec != "N/A":
        parts = re.split(r"[\/\s\-]+", combined_ec)
        engine = parts[0] if len(parts) > 0 else "N/A"
        chassis = parts[1] if len(parts) > 1 else "N/A"
    else:
        engine = find(r"Engine\s*(?:No\.?|Number)\s*[:\-]?\s*([A-Z0-9]{6,})", t)
        chassis = find(r"Chassis\s*(?:No\.?|Number)\s*[:\-]?\s*([A-Z0-9]{6,})", t)

    # --- Payment Mode ---
    if "payment aggregator" in t.lower():
        pay_mode = "PAYMENT AGGREGATOR"
    elif "online" in t.lower():
        pay_mode = "Online Payment"
    elif "cheque" in t.lower():
        pay_mode = "Cheque"
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
        "CUST_MOBILE_NUMBER": mobile,
        "CUST_EMAIL": email,
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
st.caption("Built with 💙 Streamlit + PyPDF2 + Regex | Supports Tata AIG, Zurich Kotak, Royal Sundaram, ICICI Lombard, Reliance & more")
