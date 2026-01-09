import streamlit as st
import pandas as pd
import re
from io import BytesIO
from PyPDF2 import PdfReader

# Configure the Streamlit page
st.set_page_config(page_title="PDF to Excel - Policy Extractor", layout="centered")

# --- UI Setup ---
st.title("📄 PDF Policy Extractor → Excel")
st.write("Upload one or more insurance policy PDFs to extract key details into a structured Excel format.")

# File uploader widget
uploaded_files = st.file_uploader(
    "Upload Policy PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

# Define columns structure globally for consistent error handling and output order
desired_columns = [
    "Customer Id", "Customer Name", "Policy No", "Effective Date", "Expiry Date",
    "Product Name", "Sum Insured / IDV", "Premium Paid (Incl. GST)", "Intermediary Name",
    "CUST_MOBILE_NUMBER", "CUST_EMAIL", "Fuel Type", "Vehicle No / Registration Number", "CHASSIS NUM", "ENGINE NUM",
    "VEHICLE INFO", "Payment Mode", "File Name"
]

# --- Function to safely extract fields using Regex ---
def find(pattern, text, flags=re.IGNORECASE | re.DOTALL):
    """
    Safely searches for a pattern in text and returns the content of the
    first capturing group. Returns an empty string if no match is found.
    """
    try:
        match = re.search(pattern, text, flags)
        if match and match.lastindex:
            return match.group(1).strip() 
        elif match:
            # If no capturing group is used, return the whole match
            return match.group(0).strip()
    except Exception as e:
        # Log the regex error for debugging complex cases
        print(f"Regex error for pattern {pattern}: {e}")
        return ""
    return ""

# --- Core Extraction Logic (Maximum Robustness) ---
def extract_policy_details(text):
    """
    Extracts structured data points from the raw text content of a policy PDF.
    Refined for better separation of Customer Name and Address/ID, and improved mappings for VEHICLE INFO and GVW.
    """
    # Normalize text: replace newlines and reduce multiple spaces
    text_clean = text.replace("\n", " ").replace("\r", " ").strip()
    text_clean = re.sub(r'\s+', ' ', text_clean).strip()
    
    # Common Date Pattern (DD/MM/YYYY, DD-Month-YYYY, DD-MMM-YYYY)
    DATE_REGEX = r"([0-3]?\d[\s\/\-][A-Za-z\d]{1,}[\s\/\-]\d{2,4})"
    
    # --- 1. Identify Policy Number for contextual search ---
    # Non-greedy capture of the policy number, cleaned to remove extra "Policy" at the end
    policy_no_raw = find(r"(?:Policy\s*N(?:o\.?|umber)?|Certificate\s*No)\s*[:\-\s]*([A-Z0-9\/\-]{4,})", text_clean)
    policy_no = re.sub(r'Policy$', '', policy_no_raw).strip() if policy_no_raw else "N/A"

    # --- 2. Aggressive Date Extraction (Fallback 1: Adjacent Dates) ---
    dates = []
    if policy_no and policy_no != "N/A":
        # Search for the policy number followed by two dates
        date_search_pattern = re.compile(
            re.escape(policy_no) + r".{0,100}?" + DATE_REGEX + r".{0,50}?" + DATE_REGEX, 
            re.IGNORECASE | re.DOTALL
        )
        match = date_search_pattern.search(text_clean)
        if match:
            # Group 1 is the Effective Date, Group 2 is the Expiry Date
            dates = [match.group(1).strip(), match.group(2).strip()]

    customer_name_raw = find(
        r"(?:Insured|Customer|Policyholder)\s*Name?\s*[:\-\s]*(.*?)(?=\s*(?:Policy\s*N|VGC|D\d+|Address|Pin\s*Code|City|State|Effective\s*Date|ID|Mobile|Vehicle|Premium|\d{10}))", 
        text_clean
    ).strip()
    
    if customer_name_raw:
        customer_name_clean = customer_name_raw.split(',')[0].strip()  # Split on first comma and take first part
        customer_name_clean = re.sub(r'[\d\s]+$', '', customer_name_clean).strip()  # Remove trailing numbers/spaces
        # Remove common titles
        customer_name_clean = re.sub(r'^(Mr\.|Mrs\.|Ms\.|Dr\.|M/s\.)\s*', '', customer_name_clean, flags=re.IGNORECASE).strip()
    else:
        customer_name_clean = "N/A"

         # --- Special handling for CUST_EMAIL: Find all emails, exclude service/company emails ---
    all_emails = re.findall(r"([a-zA-Z0-9._%+*-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", text_clean, re.IGNORECASE)
    # Exclude emails that look like service emails (e.g., containing 'services' or from 'royalsundaram.in')
    service_email_patterns = [r".*services.*", r".*@royalsundaram\.in"]
    customer_emails = [email for email in all_emails if not any(re.match(pattern, email, re.IGNORECASE) for pattern in service_email_patterns)]
    cust_email = customer_emails[0] if customer_emails else "N/A"
    
    details = {
        # --- Customer Details ---
        # 1. Customer ID: Stronger patterns for common IDs
        "Customer Id": find(
            r"(?:Customer|Client|Insured|Agent)\s*(?:ID|Code|No\.?)\s*[:\-\s]*([A-Z0-9\/\-]+)", 
            text_clean
        ) or "N/A",
        
        # 2. Customer Name: Now with cleanup
        "Customer Name": customer_name_clean,
        
        # --- Policy Details ---
        "Policy No": policy_no, 
        
        # Effective Date - Primary search by label, then use adjacent date fallback
        "Effective Date": find(r"(?:Effective\s*Date|from\s*Date|Date\s*of\s*Issue|Period\s*from)\s*[:\-\s]*" + DATE_REGEX, text_clean)
                          or (dates[0] if dates else "N/A"),
        
        # Expiry Date - Primary search by label, then use adjacent date fallback
        "Expiry Date": find(r"(?:Expiry\s*Date|to\s*Date|Valid\s*until|Period\s*to)\s*[:\-\s]*" + DATE_REGEX, text_clean)
                       or (dates[1] if len(dates) > 1 else "N/A"),
        
        # Product Name - Enhanced to capture specific policy types directly or via labels
        "Product Name": find(r"(?:Product\s*Name|Policy\s*Type|Plan\s*Name|Cover\s*Type)\s*[:\-\s]*(.*?)(?=\s*(?:Sum\s*Insured|Premium|Policy\s*N|Effective\s*Date|\d{1,3},\d{3}|Intermediary|Payment|Vehicle|Fuel|IDV|Customer|Insured))", text_clean) 
                        or find(r"(Digit\s+Private\s+Car\s+Stand-alone\s+Own\s+Damage\s+Policy)", text_clean) 
                        or find(r"(Goods\s+Carrying\s+Vehicle\s+Policy)", text_clean) 
                        or find(r"([A-Za-z\s\-]+(?:Policy|Plan))", text_clean)  # Broad capture for any "X Policy" or "X Plan"
                        or "N/A",
        
        # --- Financial Details ---
        "Sum Insured / IDV": find(
        r"(?:IDV|Sum\s*Insured|Liability\s*Limit)[^\d]*([\d,.]+)", 
        text_clean
        ) or "N/A",
        "Premium Paid (Incl. GST)": find(r"(?:Total\s*Premium|Premium\s*Paid|Gross\s*Premium|Total\s*Amount\s*Payable)[^\d]*([\d,\.]+)\s*(?:Rs\.|USD|INR|\b)", text_clean) or "N/A",
        
        # --- Intermediary/Payment Details ---
        "Intermediary Name": find(r"Intermediary\s*Name\s*[:\-]?\s*([A-Za-z\s\.,]+PRIVATE\s+LIMITED)", text_clean) or "N/A",
        "Payment Mode": find(r"Payment\s*Mode\s*[:\-\s]*([A-Za-z\s]+)", text_clean) 
                        or find(r"(?:Mode\s*of\s*Payment|Payment\s*Method|Paid\s*by)\s*[:\-\s]*([A-Za-z\s]+)", text_clean) 
                        or find(r"(?:Cash|Cheque|Online|Credit\s*Card|Debit\s*Card|Net\s*Banking)", text_clean)  # Common payment modes
                        or "N/A",
        
        # --- Contact Details ---
        "CUST_MOBILE_NUMBER": find(r"(?:Mobile|Phone|Contact)\s*N(?:o\.?|umber)?\s*[:\-\s]*([\+x\d]{10,15})", text_clean) or "N/A",
         "CUST_EMAIL": cust_email,

        # --- Vehicle Details ---
        "Fuel Type": find(r"Fuel\s*Type\s*[:\-]?\s*([A-Za-z]+)", text_clean) or "N/A",
         "Vehicle No / Registration Number": find(r"(?:Vehicle\s*N(?:o\.?|umber)|Regn\.?\s*No\.?|Registration\s*Number|Reg\s*No|Plate\s*No|Vehicle\s*Registration\s*No)\s*[:\-\s]*([A-Z0-9\s]{4,}?)(?=\s*Type\s*of\s*Body|Fuel\s*Type|\b)", text_clean) or "N/A",
    
        
        "CHASSIS NUM": find(r"Chassis\s*No\.?\s*[:\-\s]*([A-Z0-9]{5,})", text_clean) or "N/A",
        "ENGINE NUM": find(r"Engine\s*No\.?\s*[:\-\s]*([A-Z0-9]{5,})", text_clean) or "N/A",
        
        # VEHICLE INFO - Prioritize "Make of the Vehicle", then fall back to existing patterns
        "VEHICLE INFO": find(r"(?:Make\s*of\s*the\s*Vehicle)\s*[:\-\s]*(.*?)(?=\s*(?:Fuel\s*Type|Chassis\s*No|Engine\s*No|Vehicle\s*N|Registration|CC|GVW|Type\s*of\s*Body))", text_clean)
                        or find(r"(?:Make\s*and\s*Model|Vehicle\s*Make\s*and\s*Model)\s*[:\-\s]*(.*?)(?=\s*(?:Fuel\s*Type|Chassis\s*No|Engine\s*No|Vehicle\s*N|Registration|CC|GVW|Type\s*of\s*Body))", text_clean) 
                        or find(r"(?:VOLKSWAGEN\s+VIRTUS|Ashok\s+Leyland\s+Ltd\.\s+MJ\d+.*?(?:T\s*\d+|TIPPER).*?BSVI)", text_clean)  # Specific for known models
                        or find(r"([A-Z][a-z]+\s+[A-Z][a-z]+\s+(?:Ltd\.|Inc\.|Corp\.)?\s*[A-Z0-9\s\-]+(?:BSVI|BSIV|etc\.))", text_clean)  # Broad for vehicle descriptions
                        or "N/A",
    }
    
    # Final cleanup and replace empty values with "N/A"
    for key, value in details.items():
        if isinstance(value, str):
            cleaned_value = re.sub(r'^[\s\:\-]+|[\s\:\-]+$', '', value)
            details[key] = cleaned_value if cleaned_value else "N/A"
        elif not value:
            details[key] = "N/A"
    
    return details

# --- Main Processing Block ---
if uploaded_files:
    st.info(f"Processing {len(uploaded_files)} PDF(s)... please wait ⏳")
    all_data = []

    for file in uploaded_files:
        try:
            file.seek(0)
            reader = PdfReader(file)
            text = ""
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + " "
            
            extracted = extract_policy_details(text)
            extracted["File Name"] = file.name
            all_data.append(extracted)
        
        except Exception as e:
            st.error(f"Failed to process file {file.name}: {e}")
            # Append an error record to the data frame
            error_record = {col: "N/A" for col in desired_columns}
            error_record["Policy No"] = f"ERROR: See console for {file.name}"
            error_record["File Name"] = file.name
            all_data.append(error_record)

    
    df = pd.DataFrame(all_data)
    # Ensure the columns are in the desired order and fill any remaining NaNs
    output_df = df.reindex(columns=desired_columns).fillna("N/A")
    output_df = output_df.applymap(lambda x: "N/A" if isinstance(x, str) and not x.strip() else x)
    
    st.success("✅ Extraction complete! Review the data below.")
    st.dataframe(output_df)
    
    if not output_df.empty:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            output_df.to_excel(writer, index=False, sheet_name="Policy Details")
        
        st.download_button(
            label="📥 Download Extracted Policy Data as Excel",
            data=output.getvalue(),
            file_name="policy_details_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No data was extracted.")

st.markdown("---")
st.caption("Built with PyPDF2 for text extraction and Streamlit for the user interface.")
