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

# --- Extraction Logic ---
def extract_policy_details(text):
    # Clean whitespace
    t = re.sub(r'\s+', ' ', text.replace("\n", " "))

    # --- DATE pattern ---
    DATE = r"(\d{1,2}\s+[A-Za-z]{3}\s+'?\d{2,4}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})"

     # --- Policy Number ---
    policy_no = find(r"Policy\s*(?:No\.?|Number|No\s*&\s*Certificate\s*No)\s*[:\-]?\s*([A-Z0-9\/\-]{5,})", t)


    # --- Effective & Expiry Dates ---
    eff_date = exp_date = "N/A"
    match = re.search(r"(?:OD\s*Cover\s*Period|Period\s*of\s*Insurance).*?" + DATE + r".*?(?:to|till)\s*" + DATE, t, re.IGNORECASE)
    if match:
        eff_date, exp_date = match.group(1).strip(), match.group(2).strip()

    # --- Customer Name ---
    cust_name = find(r"(?:Insured\s*Name|Customer\s*Name|Policyholder\s*Name)\s*[:\-]?\s*([A-Za-z\s\.\']+?)(?:\s+Address|\s*$)", t)

    # --- Financial Details ---
    idv = find(r"(?:IDV|Sum\s*Insured|Liability\s*Limit)[^\d]*([\d,.]+)", t)
    premium = find(r"(?:Total\s*Premium|Premium\s*Amount|Gross\s*Premium|Total\s*Payable)\s*[₹Rs\.:\s]*([\d,\.]+)", t)

    # --- Intermediary ---
    intermediary = find(
        r"(?:Intermediary\s*Name|Agent\s*Name)\s*[:\-]?\s*([A-Za-z\s\.,]+?)(?:\s+Agent\s+License|\s+Code|\s+Private|\s+Ltd|\s*$)", t
    )

    # --- Customer Mobile Number (handles Tata AIG format) ---
    customer_mobile = find(r"Customer\s*contact\s*number\s*[:\-]?\s*([\d\*\s]+)", t)
    if customer_mobile == "N/A":
        customer_mobile = find(r"(?:Mobile\s*No\.?|Contact\s*No\.?|Phone\s*No\.?)\s*[:\-]?\s*(\b[6-9]\d{9}\b)", t)
    if customer_mobile == "N/A":
        customer_mobile = find(r"(\b[6-9]\d{9}\b)", t)
    if customer_mobile != "N/A":
        customer_mobile = customer_mobile.replace(" ", "").replace("*", "X")

    # --- Email Extraction ---
    all_emails = re.findall(r"([a-zA-Z0-9._%+*-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", t, re.IGNORECASE)
    service_email_patterns = [r".*services.*", r".*@royalsundaram\.in", r".*@tataaig\.com", r".*@icicilombard\.com"]
    customer_emails = [email for email in all_emails if not any(re.match(p, email, re.IGNORECASE) for p in service_email_patterns)]
    cust_email = customer_emails[0] if customer_emails else "N/A"

    # --- Vehicle Info ---
    fuel = find(r"Fuel\s*Type\s*[:\-]?\s*([A-Za-z]+)", t)
    reg_no = find(r"(?:Registration\s*No|Vehicle\s*No|Regn\s*No)\s*[:\-]?\s*([A-Z0-9\s]{5,}?)(?:\s+Registration\s+Authority|\s*$)", t)
    chassis = find(r"Chassis\s*(?:No\.?|Number)\s*[:\-]?\s*([A-Z0-9]{5,})", t)
    engine = find(r"(?:Engine\s*(?:No\.?|Number)|Battery\s*Number)\s*[:\-]?\s*([A-Z0-9]{5,})", t)

    # --- Product ---
    product = find(r"(?:Product\s*Name|Policy\s*Type|Cover\s*Type)\s*[:\-]?\s*([A-Za-z\s]+Policy)", t)
    if product == "N/A":
        if "Private Car" in t:
            product = "Private Car Package Policy"
        elif "Goods Carrying Vehicle" in t:
            product = "Goods Carrying Vehicle Policy"

    # --- Vehicle Make / Model ---
    vehicle_info = find(
        r"(?:Make\s*/\s*Model|Make\s*and\s*Model|Vehicle\s*Make)\s*[:\-]?\s*([A-Za-z0-9\s\-/]+?)(?:\s+Fuel\s+Type|\s*$)",
        t,
    )
    if vehicle_info == "N/A":
        vehicle_info = find(r"([A-Za-z]+\s+Ltd\.?\s+[A-Za-z0-9\s\-]+BSVI?)", t)

    # --- Payment Mode ---
    pay_mode = find(r"(?:Payment\s*Mode|Mode\s*of\s*Payment)\s*[:\-]?\s*([A-Za-z\s]+)", t)
    if "paymentLinkCustomer" in t:
        pay_mode = "Online Payment"
    elif pay_mode == "N/A" and "cheque" in t.lower():
        pay_mode = "Cheque"

    return {
        "Customer Id": find(r"(?:Customer\s*ID|Client\s*ID)\s*[:\-]?\s*([A-Z0-9\-\/]+)", t),
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
