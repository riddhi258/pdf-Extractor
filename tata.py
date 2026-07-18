import streamlit as st
import pandas as pd
import re
from io import BytesIO
from PyPDF2 import PdfReader

# --- Streamlit Config ---
st.set_page_config(page_title="PDF to Excel - Policy Extractor", layout="centered")
st.title("📄 PDF Policy Extractor → Excel")
st.write("Upload one or more insurance policy PDFs (Tata AIG, Royal Sundaram, ICICI Lombard, etc.) to extract key details into a structured Excel file.")

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
    return match.group(1).strip() if match else "N/A"

def extract_policy_details(text):
    # Standardize spaces and clean layout components
    t = re.sub(r'\s+', ' ', text)
    t_clean = t.replace("|", ":")

    # --- Date Parsing Engine ---
    DATE_RE = r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})"
    dates = re.findall(DATE_RE, t_clean)
    
    eff_date = "N/A"
    exp_date = "N/A"
    if len(dates) >= 2:
        # Avoid picking up vehicle registration date (which usually follows later in text)
        eff_date = dates[0]
        exp_date = dates[1]

    # --- Policy Number ---
    policy_no = find(r"Policy\s*(?:No\.?|Number)\s*[:\-]?\s*([0-9\s]{10,18})", t_clean)
    if policy_no != "N/A":
        policy_no = re.sub(r'\s+', ' ', policy_no).strip()

    # --- Customer Name ---
    # Captures name while rejecting trailing table elements
    cust_name = find(r"Name\s*:\s*(?:Mr\.|Ms\.|Mrs\.)?\s*([A-Za-z\s\.\']+?)\s*(?:Address|:)", t_clean)

    # --- Financial Details ---
    # Focuses on explicit numeric totals
    idv = find(r"Total\s*IDV\s*:\s*([\d,.]+)", t_clean)
    if idv == "N/A":
        idv = find(r"Vehicle\s*IDV\s*\(?\s*₹\s*\)?\s*:\s*([\d,.]+)", t_clean)
        
    premium = find(r"Total\s*Policy\s*Premium\s*:\s*(?:[₹Rs\.\s])*([\d,.]+)", t_clean)
    if premium == "N/A":
        premium = find(r"Premium\s*Amount\s*\(Including\s*GST\)\s*:\s*(?:[₹Rs\.\s])*([\d,.]+)", t_clean)

    # --- Intermediary Name ---
    # Tata AIG places names in a secondary line beneath the row headers
    intermediary = "N/A"
    inter_match = re.search(r"Agent/Intermediary\s*Contact\s*No\.\s*:\s*(.*?)\s*:\s*[A-Z0-9]+", t_clean, re.IGNORECASE)
    if inter_match:
        # Grab text directly following the headers structure
        raw_segment = inter_match.group(1).strip()
        # Clean out lingering headers if any
        intermediary = raw_segment.split(':')[-1].strip()

    # --- Customer Mobile Number ---
    customer_mobile = find(r"Contact\s*No\.\s*:\s*([\+\d\*\s]+)", t_clean)
    if customer_mobile != "N/A":
        customer_mobile = customer_mobile.replace(" ", "").replace("*", "X").replace("+91", "")

    # --- Email Extraction ---
    all_emails = re.findall(r"([a-zA-Z0-9._%+*-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", t_clean, re.IGNORECASE)
    service_email_patterns = [r".*services.*", r".*@royalsundaram\.in", r".*@tataaig\.com", r".*@icicilombard\.com"]
    customer_emails = [email for email in all_emails if not any(re.match(p, email, re.IGNORECASE) for p in service_email_patterns)]
    cust_email = customer_emails[0] if customer_emails else "N/A"

    # --- Vehicle Info Engine ---
    fuel = find(r"Fuel\s*Type\s*:\s*([A-Za-z]+)", t_clean)
    
    reg_no = find(r"Registration\s*No\.\s*:\s*([A-Z]{2}\s*[0-9]{1,2}\s*[A-Z0-9\s]{2,8})", t_clean)
    if reg_no != "N/A": 
        reg_no = reg_no.strip()
        
    chassis = find(r"Chassis\s*(?:No\.?|Number)\s*:\s*([A-Z0-9]{5,})", t_clean)
    engine = find(r"Engine\s*(?:No\.?|Number)\s*(?:\/Motor\s*No\.\s*\(For\s*EV\))?\s*:\s*([A-Z0-9]{5,})", t_clean)

    # --- Product Name ---
    product = "N/A"
    if "Two-Wheeler Package Policy" in t_clean:
        product = "Two-Wheeler Package Policy"
    elif "Private Car" in t_clean:
        product = "Private Car Package Policy"

    # --- Vehicle Make / Model ---
    vehicle_info = find(r"Make\s*/\s*Model\s*/\s*Variant\s*:\s*([A-Za-z0-9\s\-/]+?)\s*:\s*Fuel", t_clean)

    # --- Payment Mode ---
    pay_mode = "N/A"
    if "paymentLinkCustomer" in t_clean:
        pay_mode = "Online Payment"
    else:
        pay_mode = find(r"(?:Payment\s*Mode|Mode\s*of\s*Payment)\s*[:\-]?\s*([A-Za-z\s]+)", t_clean)

    return {
        "Customer Id": find(r"Client\s*ID\s*:\s*([A-Z0-9\-\/]+)", t_clean),
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
