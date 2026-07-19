import streamlit as st

import pandas as pd

import re

from io import BytesIO

from PyPDF2 import PdfReader

from datetime import datetime



# --- Streamlit Config ---

st.set_page_config(page_title="PDF to Excel - Policy Extractor", layout="centered")

st.title("📄 PDF Policy Extractor → Excel")

st.write("Upload insurance policy PDFs (Tata AIG, Zurich Kotak, Royal Sundaram, ICICI Lombard, Reliance, National, etc.) to extract details into Excel.")

allow_underscore_names = st.sidebar.checkbox(
    "Allow names starting with underscore (_)", 
    value=False,
    help="Enable this if the customer names in your PDFs are prefixed with an underscore (e.g., _John Doe)."
)

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

    return match.group(1).strip() if match and match.lastindex else "N/A"



# --- Date Formatter ---

def format_date(date_str):

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d %b %Y", "%d %B %Y"):

        try:

            return datetime.strptime(date_str.strip(), fmt).strftime("%d %b '%y")

        except:

            continue

    return date_str



# --- Extraction Function ---

def extract_policy_details(text, file_name, allow_underscore_names):
    
    t = re.sub(r'\s+', ' ', text.replace("\n", " "))



    # Detect National Insurance

    is_national = "national insurance" in t.lower()



    # --- Policy Number ---

    policy_no = find(r"Policy\s*(?:No\.?|Number)\s*[:\-]?\s*([A-Z0-9\/\-]{6,})", t)



    # --- Effective / Expiry Date ---

    eff_date = exp_date = "N/A"

    if is_national:

        match = re.search(

            r"Policy\s*Effective\s*from.*?on\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}).*?to\s*midnight\s*of\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})",

            t, re.IGNORECASE)

        if match:

            eff_date, exp_date = format_date(match.group(1)), format_date(match.group(2))

    else:

        match = re.search(

            r"(?:Period\s*of\s*Insurance|Policy\s*Period).*?(?:From|Valid\s*from)[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}).*?(?:To|Till|Up\s*to)[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})",

            t, re.IGNORECASE)

        if match:

            eff_date, exp_date = format_date(match.group(1)), format_date(match.group(2))



    # --- Customer ID ---

    cust_id = find(r"(?:Customer\s*ID|Client\s*ID)\s*[:\-]?\s*([0-9A-Z\/\-]+)", t)



    # --- Customer Name Extraction ---
    if allow_underscore_names:
        cust_name = find(r"(?:Insured|Customer|Policyholder|Customer Name)\s*Name\s*[:\-]?\s*([_A-Za-z\s\.\']+)", t)
        if cust_name != "N/A":
            cust_name = cust_name.strip()
            if cust_name.startswith("_"):
                cust_name = "_" + cust_name[1:].strip().title()
            else:
                cust_name = cust_name.title()
    else:
        cust_name = find(r"(?:Insured|Customer|Policyholder|Customer Name)\s*Name\s*[:\-]?\s*([A-Za-z\s\.\']+)", t)
        cust_name = cust_name.strip().title() if cust_name != "N/A" else "N/A"



    # --- Product Name ---

    if is_national:

        product = find(r"Class\s*of\s*Vehicle\s*[:\-]?\s*([A-Za-z\s\/,]+)", t)

    else:

        product = find(r"(?:Product\s*Name|Policy\s*Type|Cover\s*Type)\s*[:\-]?\s*([A-Za-z\s\-\(\)\/]+)", t)

    if product == "N/A" and "private car" in t.lower():

        product = "Private Car Package Policy"

    else: product = find(r"(?:Product\s*Name|Policy\s*Type|Cover\s*Type)\s*[:\-]?\s*([A-Za-z\s]+)", t)



    # --- Sum Insured / IDV ---

         # --- Sum Insured / IDV ---

    # Handles bilingual "वाहन का आई.डी.वी/Vehicle IDV" or "Vehicle IDV" or "Total Value"

    idv = find(

        r"(?:वाहन\s*का\s*आई\.डी\.वी\/Vehicle\s*IDV|Vehicle\s*IDV|Insured\s*Declared\s*Value|Sum\s*Insured)\s*[`₹:\-]?\s*([\d,\.]+)",

        t

    )



    if idv == "N/A":

        idv = find(r"Total\s*Value\s*[₹:\-\s]*([\d,\.]+)", t)



    # Clean up formatting

    if idv != "N/A":

        try:

            idv = f"{float(idv.replace(',', '').strip()):,.2f}"

        except:

            pass



    # --- Premium (Premium Paid Incl. GST) ---

         # --- Premium Paid (Incl. GST) ---

            # --- Premium Paid (Incl. GST) ---

    # Handles "कुल राशि Total Amount" in Hindi-English mix with any spacing or hidden characters

    premium = find(

        r"क\s*ु\s*ल\s*र\s*ा\s*श\s*ि.*?Total\s*Amount\s*[₹`:\-\s]*([\d,.,]+)",

        t,

        flags=re.IGNORECASE

    )



    # If still not found, try a simpler English-only fallback

    if premium == "N/A":

        premium = find(

            r"Total\s*Amount\s*[₹`:\-\s]*([\d,.,]+)",

            t,

            flags=re.IGNORECASE

        )



    # Format cleanly

    if premium != "N/A":

        try:

            premium = f"{float(premium.replace(',', '').strip()):,.2f}"

        except:

            pass







    # --- Intermediary ---

    if is_national:

        intermediary = find(r"\bName\s*[:\-]?\s*([A-Za-z\s\.\']+)", t)

        if cust_name in intermediary:

            intermediary = "N/A"

    else:

        intermediary = find(r"(?:Intermediary\s*Name|Agent\s*Name)\s*[:\-]?\s*([A-Za-z\s\.,]+)", t)

    intermediary = re.sub(r"\b(Intermediary|Code)\b", "", intermediary).strip().title()



    # --- Customer Mobile Number ---

    mobile = find(r"(?:Phone|Cell|Mobile\s*No\.?)\s*[:\-]?\s*([0-9\*\s]{8,15})", t)

    if mobile != "N/A":

        mobile = mobile.replace(" ", "").replace("*", "X")



    # --- Customer Email (E-Mail:) ---

        # --- Customer Email (E-Mail:) ---

    # Capture only the email right after "E-Mail" or "ई-मेल"

    cust_email_match = re.search(

        r"(?:ई[-\s]*मेल|E[-\s]*Mail)\s*[:\-]?\s*([A-Za-z0-9.*_%+/-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})",

        t

    )



    cust_email = cust_email_match.group(1).strip() if cust_email_match else "N/A"



    # Step 2: Exclude known company/service domains even if matched

    exclude_domains = [

        "royalsundaram.in", "tataaig.com", "reliancegeneral.co.in",

        "icicilombard.com", "nationalinsurance.nic.co.in", "nic.co.in",

        "kotak.com", "hdfcergo.com", "tvs.in"

    ]

    if any(domain in cust_email.lower() for domain in exclude_domains):

        # Fallback – look for a personal Gmail/Yahoo/Hotmail address elsewhere

        all_emails = re.findall(

            r"[A-Za-z0-9.*_%+/-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",

            t

        )

        for email in all_emails:

            if not any(domain in email.lower() for domain in exclude_domains):

                cust_email = email

                break

        else:

            cust_email = "N/A"



    cust_email = cust_email.replace(" ", "")





    # --- Fuel Type ---

    fuel = find(r"(?:Type\s*of\s*Fuel|Fuel\s*Type)\s*[:\-]?\s*([A-Za-z]+)", t)



    # --- Vehicle Info ---

    vehicle_info = find(r"(?:Make|Manufacturer)\s*[:\-]?\s*([A-Za-z0-9\s&\.\-]+)", t)



    # --- Registration Number ---

    reg_no = find(r"(?:Regn\.?\s*Number|Registration\s*No\.?)\s*[:\-]?\s*([A-Z]{2}[\s\-]?\d{2}[\s\-]?[A-Z]{1,2}[\s\-]?\d{4})", t)



    # --- Engine / Chassis ---

    engine = find(r"(?:Engine\s*or\s*M\/c\s*No\.?|Engine\s*Number)\s*[:\-]?\s*([A-Z0-9\s]+)", t)

    chassis = find(r"(?:Chassis\s*Number|Chassis\s*No\.?)\s*[:\-]?\s*([A-Z0-9\s]+)", t)



    # --- Payment Mode ---

    if "online" in t.lower():

        pay_mode = "Online Payment"

    elif "aggregator" in t.lower():

        pay_mode = "Payment Aggregator"

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

        "CUST_EMAIL": cust_email,

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

        data = extract_policy_details(text, file.name, allow_underscore_names)

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

st.caption("Built with 💙 Streamlit + PyPDF2 + Regex | Supports Tata AIG, Reliance, Zurich Kotak, Royal Sundaram, ICICI Lombard & National Insurance") add exctration name option where name start with under score 

