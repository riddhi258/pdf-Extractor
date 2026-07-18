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
    # Clean up sequential spaces while retaining standard formatting structure
    t = re.sub(r'\s+', ' ', text)
    t_pipe = t.replace("\n", " ")

    # --- 1. Date Extraction Module ---
    # Captures dates inside the target period format safely
    DATE_RE = r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})"
    dates = re.findall(DATE_RE, t_pipe)
    eff_date = dates[0] if len(dates) > 0 else "N/A"
    exp_date = dates[1] if len(dates) > 1 else "N/A"

    # --- 2. Customer Name ---
    # Extracts the full name right from the opening welcome salute text
    cust_name = find(r"Hi\s+(Mr\.|Ms\.|Mrs\.)?\s*([A-Za-z\s\.\']+?)\s+Welcome to", t_pipe)
    if cust_name == "N/A":
        cust_name = find(r"Name\s*:\s*(?:Mr\.|Ms\.|Mrs\.)?\s*([A-Za-z\s\.\']+?)\s*:", t_pipe)

    # --- 3. Policy Number ---
    policy_no = find(r"Policy\s*(?:No\.?|Number)\s*[:\-]?\s*(\d{6,15})", t)

    # --- 4. Product Name ---
    product = "N/A"
    if "Two-Wheeler Package Policy" in t_pipe:
        product = "Two-Wheeler Package Policy"
    elif "Private Car" in t_pipe:
        product = "Private Car Package Policy"

    # --- 5. Sum Insured / IDV ---
    # Pulls the total IDV from the numbered insurance schedule array structure
    idv = find(r"Total\s*IDV\s*(?:\(?\s*₹\s*\)?\s*)?:\s*([\d,.]+)", t_pipe)
    if idv == "N/A":
        idv = find(r"1\s*:\s*(\d{4,7})\s*:\s*0\s*:\s*0", t_pipe)  # Looks for the raw row values sequence

    # --- 6. Premium Amount ---
    premium = find(r"Total\s*Policy\s*Premium\s*:\s*(?:[₹Rs\.\s])*([\d,.]+)", t_pipe)
    if premium == "N/A":
        premium = find(r"Premium\s*Amount\s*\(Including\s*GST\)\s*:\s*(?:[₹Rs\.\s])*([\d,.]+)", t_pipe)

    # --- 7. Intermediary Details ---
    # Targets the separate string chunk generated beneath the combined structural headers
    intermediary = "N/A"
    inter_match = re.search(r"Agent/Intermediary\s*Contact\s*No\..*?\|\s*([A-Za-z\s\.]+?)\s*\|\s*[A-Z0-9]+", t)
    if inter_match:
        intermediary = inter_match.group(1).strip()
    else:
        # Fallback tracking if pipe structures are stripped out early
        intermediary = find(r"Agent/Intermediary\s*Contact\s*No\.\s*:\s*([A-Za-z\s\.]+?)\s*:\s*[A-Z0-9]+", t_pipe)

    # --- 8. Vehicle Tracking Info --
    fuel = find(r"Fuel\s*Type\s*[:\-]?\s*([A-Za-z]+)", t_pipe)
    chassis = find(r"Chassis\s*(?:No\.?|Number)\s*[:\-]?\s*([A-Z0-9]{10,17})", t_pipe)
    engine = find(r"Engine\s*(?:No\.?|Number).*?:\s*([A-Z0-9]{10,17})", t_pipe)
    
    # Extract the full Make / Model line block directly
  # --- Vehicle Info (Make / Model / Modal / Variant / Make and Model) ---
    vehicle_info = find(r"Make\s*/\s*Model\s*/\s*Variant\s*:\s*([A-Za-z0-9\s\-/]+?)\s*:\s*Fuel", t_pipe)

    # --- Registration Number ---
    reg_no = find(
        r"(?:Registration\s*No\.?|Vehicle\s*No\.?|Regn\s*No\.?|Registration\s*Number)\s*[:\-]?\s*([A-Z]{2}\s*\d{2}\s*[A-Z]{1,2}\s*\d{4})",
        t
    )

    # --- 9. Mobile, Email & Identifiers ---
      # --- Customer Mobile ---
    mobile = find(r"(?:Mobile\s*No\.?|Customer\s*contact\s*number)\s*[:\-]?\s*([\d\*\s]+)", t)
    if mobile == "N/A":
        mobile = find(r"\b[6-9]\d{9}\b", t)
    if mobile != "N/A":
        mobile = mobile.replace(" ", "").replace("*", "X")

    all_emails = re.findall(r"([a-zA-Z0-9._%+*-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", t_pipe, re.IGNORECASE)
    service_email_patterns = [r".*services.*", r".*@royalsundaram\.in", r".*@tataaig\.com", r".*@icicilombard\.com"]
    customer_emails = [email for email in all_emails if not any(re.match(p, email, re.IGNORECASE) for p in service_email_patterns)]
    cust_email = customer_emails[0] if customer_emails else "N/A"

    pay_mode = "N/A"
    if "paymentLinkCustomer" in t_pipe:
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
