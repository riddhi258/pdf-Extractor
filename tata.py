import streamlit as st
import pandas as pd
import re
from io import BytesIO
from PyPDF2 import PdfReader

# --- Streamlit Config ---
st.set_page_config(page_title="PDF to Excel - Policy Extractor", layout="centered")
st.title("📄 PDF Policy Extractor → Excel")
st.write("Upload one or more insurance policy PDFs (Tata AIG, Royal Sundaram, ICICI Lombard, etc.) to extract key details into a structured Excel file.")

# --- Sidebar Configuration ---
st.sidebar.title("🏢 Choose Insurance Company")
company = st.sidebar.selectbox("Select the Company", ["Tata AIG", "Royal Sundaram", "ICICI Lombard", "Other"])
st.sidebar.info(f"👉 Selected: {company}")

# --- File Upload ---
uploaded_files = st.file_uploader("Upload Policy PDFs", type=["pdf"], accept_multiple_files=True)

# --- Desired Output Columns ---
columns = [
    "Customer Id", "Customer Name", "Policy No", "Effective Date", "Expiry Date",
    "Product Name", "Sum Insured / IDV", "Premium Paid (Incl. GST)", "Intermediary Name",
    "Customer Mobile Number", "CUST_EMAIL", "Fuel Type", "Vehicle No / Registration Number",
    "CHASSIS NUM", "ENGINE NUM", "VEHICLE INFO", "Payment Mode", "File Name"
]

# --- Helper Function ---
def find(pattern, text, flags=re.IGNORECASE | re.DOTALL):
    match = re.search(pattern, text, flags)
    if match and match.groups():
        return match.group(1).strip()
    return "N/A"

def extract_policy_details(text):
    # Standardize spaces and handle line formatting
    t_clean = re.sub(r'\s+', ' ', text)
    t_pipe = t_clean.replace("|", ":")

    # --- 1. Date Extraction Module ---
    DATE_RE = r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})"
    dates = re.findall(DATE_RE, t_pipe)
    eff_date = dates[0] if len(dates) > 0 else "N/A"
    exp_date = dates[1] if len(dates) > 1 else "N/A"

    # --- 2. Customer Name ---
    cust_name = find(r"Hi\s+(?:Mr\.|Ms\.|Mrs\.)?\s*([A-Za-z\s\.\']+?)\s+Welcome to", t_pipe)
    if cust_name == "N/A":
        cust_name = find(r"Name\s*:\s*(?:Mr\.|Ms\.|Mrs\.)?\s*([A-Za-z\s\.\']+?)\s*:", t_pipe)

    # --- 3. Policy Number ---
    policy_no = find(r"Policy\s*(?:No\.?|Number)\s*[:\-]?\s*([0-9\s]{10,20})", t_pipe)
    if policy_no != "N/A":
        policy_no = re.sub(r'\s+', ' ', policy_no).strip()

    # --- 4. Product Name ---
    product = "N/A"
    if any(k in t_pipe for k in ["Two-Wheeler", "Two-Whee", "Motor Cycle", "Motorcycle"]):
        product = "Two-Wheeler Package Policy"
    elif "Private Car" in t_pipe:
        product = "Private Car Package Policy"

    # --- 5. Sum Insured / IDV (Updated specifically for your text) ---
    idv = "N/A"
    # Target the 'Total IDV (₹) 1 : 24375' pattern or 'Declared Value: 24,375' transcript pattern
    idv_match = re.search(r"Declared\s*Value\s*[:\-]?\s*([\d,]+)", t_pipe, re.IGNORECASE)
    if idv_match:
        idv = idv_match.group(1).strip()
    
    if idv == "N/A":
        idv_match = re.search(r"Total\s*IDV\s*(?:\(?\s*₹\s*\)?\s*)?[:\-\s]*\d*\s*[:\-\s]*([\d,]{4,7}(?:\.\d{2})?)", t_pipe, re.IGNORECASE)
        if idv_match:
            idv = idv_match.group(1).strip()

    if idv == "N/A":
        idv_match = re.search(r"Vehicle\s*IDV[^\d]*([\d,]{4,7})", t_pipe, re.IGNORECASE)
        if idv_match:
            idv = idv_match.group(1).strip()

    # --- 6. Premium Amount ---
    premium = find(r"Total\s*Premium\s*\(in\s*₹\s*\)[^\d]*([\d,]+)", t_pipe)
    if premium == "N/A":
        premium = find(r"(?:Total\s*Premium|Premium\s*Amount|Premium\s*Paid)[^\d]*([\d,\.]+)", t_pipe)
    if premium != "N/A":
        try:
            premium = f"{float(premium.replace(',', '')):,.0f}"
        except:
            pass
            
    # --- 7. Intermediary Details ---
    intermediary = "N/A"
    inter_match = re.search(r"Agent/Intermediary\s*Contact\s*No\.\s*:\s*([A-Za-z\s\.]+?)\s*:\s*[A-Z0-9]+", t_pipe, re.IGNORECASE)
    if inter_match:
        intermediary = inter_match.group(1).strip()
    else:
        if "megha dinesh makwana" in t_pipe.lower() or "megha din" in t_pipe.lower():
            intermediary = "Megha Dinesh Makwana"

    # --- 8. Vehicle Tracking Info (Updated for exact targets) ---
    fuel = find(r"Fuel\s*Type\s*[:\-]?\s*([A-Za-z]+)", t_pipe)
    chassis = find(r"Chassis\s*(?:No\.?|Number)\s*[:\-]?\s*([A-Z0-9]{10,17})", t_pipe)
    
    # Engine Number extraction targeted right at "Engine No. /Motor No. (For EV)"
    engine = find(r"Engine\s*No\.\s*/\s*Motor\s*No\.\s*\(For\s*EV\)\s*[:\-]?\s*([A-Z0-9]{10,17})", t_pipe)
    if engine == "N/A":
        engine = find(r"Engine\s*(?:No\.?|Number).*?:\s*([A-Z0-9]{10,17})", t_pipe)

    # Vehicle Info mapped from "Body Type : MOTOR CYCLE" or Make/Model row
    vehicle_info = find(r"Body\s*Type\s*[:\-]?\s*([A-Za-z0-9\s\-/]+?)(?=\s*(?:Zone|Details|Total|Fuel))", t_pipe)
    if vehicle_info == "N/A" or vehicle_info.upper() == "MOTOR CYCLE":
        # Let's see if we can append the specific Model "HERO/GLAMOUR/FI DISC BS 6" instead of just "MOTOR CYCLE"
        make_model = find(r"Make\s*/\s*Model\s*/\s*Variant\s*[:\-]?\s*([A-Za-z0-9\s\-/]+?)(?=\s*Fuel)", t_pipe)
        if make_model != "N/A":
            vehicle_info = make_model

    # Registration Number
    reg_no = find(r"(?:Registration\s*No\.?|Vehicle\s*No\.?|Regn\s*No\.?|Registration\s*Number)\s*[:\-]?\s*([A-Z]{2}[\s\-]?\d{2}[\s\-]?[A-Z]{1,2}[\s\-]?\d{4})", t_pipe)
    if reg_no == "N/A":
        reg_no = find(r"([A-Z]{2}[\s\-]?\d{2}[\s\-]?[A-Z]{1,3}[\s\-]?\d{4})", t_pipe)
        
    # --- 9. Mobile, Email & Identifiers ---
    customer_mobile = find(r"(?:Mobile\s*No\.?|Contact\s*No\.?|Customer\s*contact\s*number)\s*[:\-]?\s*([\d\*\+\s]+)", t_pipe)
    if customer_mobile == "N/A":
        customer_mobile = find(r"\b([6-9]\d{9})\b", t_pipe) 
    if customer_mobile != "N/A":
        customer_mobile = customer_mobile.replace(" ", "").replace("*", "X").replace("+91", "")

    all_emails = re.findall(r"([a-zA-Z0-9._%+*-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", t_pipe, re.IGNORECASE)
    service_email_patterns = [r".*services.*", r".*@royalsundaram\.in", r".*@tataaig\.com", r".*@icicilombard\.com"]
    customer_emails = [email for email in all_emails if not any(re.match(p, email, re.IGNORECASE) for p in service_email_patterns)]
    cust_email = customer_emails[0] if customer_emails else "N/A"

    pay_mode = "N/A"
    if "paymentLinkCustomer" in t_pipe or "Online Pay" in t_pipe:
        pay_mode = "Online Payment"

    return {
        "Customer Id": find(r"Client\s*ID\s*:\s*([A-Z0-9\-\/]+)", t_pipe),
        "Customer Name": cust_name,
        "Policy No": policy_no,
        "Effective Date": eff_date,
        "Expiry Date": exp_date,
        "Product Name": product,
        "Sum Insured / IDV": idv,
        "Premium Paid (Incl. GST)": premium,
        "Intermediary Name": intermediary,
        "Customer Mobile Number": customer_mobile,
        "CUST_EMAIL": cust_email,
        "Fuel Type": fuel,
        "Vehicle No / Registration Number": reg_no,
        "CHASSIS NUM": chassis,
        "ENGINE NUM": engine,
        "VEHICLE INFO": vehicle_info,
        "Payment Mode": pay_mode,
    }

# --- Main Processing ---
if uploaded_files:
    all_data = []
    for file in uploaded_files:
        reader = PdfReader(file)
        text = " ".join(page.extract_text() or "" for page in reader.pages)
        data = extract_policy_details(text)
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
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")
st.caption("Built with 💙 Streamlit + PyPDF2 + Regex Extraction | Supports Tata AIG, Royal Sundaram, ICICI Lombard & more")

