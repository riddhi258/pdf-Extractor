import streamlit as st
import pandas as pd
import re
from io import BytesIO
import pdfplumber
from datetime import datetime


# ============================================================
# STREAMLIT CONFIG
# ============================================================
st.set_page_config(
    page_title="PDF to Excel - Bajaj Policy Extractor",
    layout="centered"
)

st.title("📄 Bajaj PDF Policy Extractor → Excel")
st.write(
    "Upload Bajaj General Insurance policy PDFs to extract key details into Excel."
)


# ============================================================
# SIDEBAR - ACCESSORY VALUE
# ============================================================

st.sidebar.header("🔧 Direct Accessory Adjustment")

accessory_value = st.sidebar.number_input(
    "Total Value of Non-Electronic Accessories (₹)",
    min_value=0.0,
    step=100.0,
    value=0.0,
    help=(
        "Enter the total charges/values for non-electronic accessories "
        "(e.g., roof racks, mats). This will be directly added to "
        "the Sum Insured / IDV for all extracted policies."
    )
)


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_files = st.file_uploader(
    "Upload Policy PDFs",
    type=["pdf"],
    accept_multiple_files=True
)


# ============================================================
# OUTPUT COLUMNS
# ============================================================

columns = [
    "Customer Id",
    "Customer Name",
    "Policy No",
    "Effective Date",
    "Expiry Date",
    "Product Name",
    "Sum Insured / IDV",
    "Premium Paid (Incl. GST)",
    "Intermediary Name",
    "Customer Number",
    "cust_email",
    "Fuel Type",
    "Vehicle No / Registration Number",
    "CHASSIS NUM",
    "ENGINE NUM",
    "VEHICLE INFO",
    "Payment Mode",
    "File Name"
]


# ============================================================
# HELPER FUNCTION
# ============================================================

def find(pattern, text, flags=re.IGNORECASE | re.DOTALL):
    try:
        match = re.search(pattern, text, flags)

        if not match:
            return "N/A"

        if match.lastindex:
            return match.group(1).strip()

        return match.group(0).strip()

    except Exception:
        return "N/A"


# ============================================================
# DATE FORMATTER
# ============================================================

def format_date(date_str):

    if not date_str or date_str == "N/A":
        return "N/A"

    try:
        clean_date = date_str.strip()

        # Remove extra words
        clean_date = clean_date.split()[0]

        formats = (
            "%d-%m-%Y",
            "%d/%m/%Y",
            "%d-%m-%y",
            "%d/%m/%y",
            "%d %b %Y",
            "%d %b '%y",
            "%d %B %Y",
            "%d %B '%y"
        )

        for fmt in formats:
            try:
                d = datetime.strptime(clean_date, fmt)
                return d.strftime("%d %b '%y")
            except ValueError:
                continue

        return date_str.strip()

    except Exception:
        return date_str


# ============================================================
# IDV CLEANER
# ============================================================

def clean_amount(value):

    if not value or value == "N/A":
        return "N/A"

    try:
        value = value.replace(",", "")
        value = value.replace("₹", "")
        value = value.replace("Rs.", "")
        value = value.replace("Rs", "")
        value = value.strip()

        return f"{float(value):,.2f}"

    except Exception:
        return value


# ============================================================
# MAIN EXTRACTION FUNCTION
# ============================================================

def extract_policy_details(text, file_name):

    # --------------------------------------------------------
    # CLEAN TEXT
    # --------------------------------------------------------

    t = text.replace("\n", " ")
    t = re.sub(r"\s+", " ", t)
    t = t.strip()

    t_lower = t.lower()


    # ========================================================
    # POLICY NUMBER
    # ========================================================

    policy_no = find(
        r"(?:Policy\s*(?:Number|No\.?)|Policy\s*Number)"
        r"\s*[:\-'\s]*"
        r"("
        r"[0-9]{2}\-[0-9]{4}\-[0-9]{10}\-[0-9]{2}"
        r"|OG\-\d{2}\-\d{4}\-\d{4}\-\d{8}"
        r"|\d{6,16}"
        r")",
        t
    )


    # ========================================================
    # EFFECTIVE / EXPIRY DATE
    # ========================================================

    eff_date = "N/A"
    exp_date = "N/A"

    date_match = re.search(
        r"From\s*[:\s]*"
        r"("
        r"\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}"
        r"|"
        r"\d{1,2}\s+[A-Za-z]{3}\s+'?\d{2,4}"
        r")"
        r".{0,300}?"
        r"(?:To\s*[:\s]*"
        r"("
        r"\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}"
        r"|"
        r"\d{1,2}\s+[A-Za-z]{3}\s+'?\d{2,4}"
        r"|Midnight"
        r"))?",
        t,
        re.IGNORECASE
    )

    if date_match:

        eff_date = format_date(date_match.group(1))

        if date_match.group(2):
            if date_match.group(2).lower() != "midnight":
                exp_date = format_date(date_match.group(2))


    # --------------------------------------------------------
    # FALLBACK DATE EXTRACTION
    # --------------------------------------------------------

    if eff_date == "N/A":

        all_dates = re.findall(
            r"\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b",
            t
        )

        if len(all_dates) >= 2:

            eff_date = format_date(all_dates[0])
            exp_date = format_date(all_dates[1])

        elif len(all_dates) == 1:

            eff_date = format_date(all_dates[0])


    # ========================================================
    # CUSTOMER ID
    # ========================================================

    cust_id = find(
        r"Customer\s*ID\s*[:\-]?\s*([0-9A-Z]+)",
        t
    )


    # ========================================================
    # CUSTOMER NAME
    # ========================================================

    cust_name = find(
        r"(?:Insured\s*Name|Name\s*\(Proposer\)|"
        r"Received\s*with\s*thanks\s*from)"
        r"\s*[:\-]?\s*"
        r"([A-Za-z][A-Za-z\s\.]*?)"
        r"(?=\s*(?:Name|Address|Customer|Policy|GSTIN|PAN|"
        r"a\s*total|Contact|Mobile|Phone|$))",
        t
    )

    if cust_name == "N/A":

        cust_name = find(
            r"Dear\s+([A-Za-z][A-Za-z\s\.]+?),",
            t
        )

    if cust_name != "N/A":

        cust_name = re.sub(r"\s+", " ", cust_name)
        cust_name = cust_name.strip().title()

    else:

        cust_name = "N/A"


    # ========================================================
    # PRODUCT NAME
    # ========================================================

    product = "N/A"


    # --------------------------------------------------------
    # 1. Product Name
    # --------------------------------------------------------

    product = find(
        r"(?:Product\s*Name|Product)"
        r"\s*[:\-]?\s*"
        r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
        r"(?=\s*(?:UIN|Policy|Customer|Insured|Vehicle|$))",
        t
    )


    # --------------------------------------------------------
    # 2. Policy Type
    # --------------------------------------------------------

    if product == "N/A":

        product = find(
            r"Policy\s*Type"
            r"\s*[:\-]?\s*"
            r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
            r"(?=\s*(?:UIN|Policy|Customer|Insured|Vehicle|$))",
            t
        )


    # --------------------------------------------------------
    # 3. Plan Name
    # --------------------------------------------------------

    if product == "N/A":

        product = find(
            r"Plan\s*Name"
            r"\s*[:\-]?\s*"
            r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
            r"(?=\s*(?:UIN|Policy|Customer|Insured|Vehicle|$))",
            t
        )


    # --------------------------------------------------------
    # 4. Product / Plan
    # --------------------------------------------------------

    if product == "N/A":

        product = find(
            r"Product\s*\/\s*Plan"
            r"\s*[:\-]?\s*"
            r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
            r"(?=\s*(?:UIN|Policy|Customer|Insured|Vehicle|$))",
            t
        )


    # --------------------------------------------------------
    # 5. Type of Policy
    # --------------------------------------------------------

    if product == "N/A":

        product = find(
            r"Type\s*of\s*Policy"
            r"\s*[:\-]?\s*"
            r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
            r"(?=\s*(?:UIN|Policy|Customer|Insured|Vehicle|$))",
            t
        )


    # --------------------------------------------------------
    # 6. TRANSCRIPT OF PROPOSAL FOR
    # --------------------------------------------------------

    if product == "N/A":

        product = find(
            r"TRANSCRIPT\s*OF\s*PROPOSAL\s*FOR\s*"
            r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
            r"(?=\s*(?:UIN|Customer|Insured|Vehicle|$))",
            t
        )


    # --------------------------------------------------------
    # 7. POLICY SCHEDULE FALLBACK
    # --------------------------------------------------------

    if product == "N/A":

        product = find(
            r"([A-Za-z0-9\s\-\&\/\(\)]+?)"
            r"\s*-\s*POLICY\s*SCHEDULE",
            t
        )


    # --------------------------------------------------------
    # 8. KNOWN POLICY WORDING
    # --------------------------------------------------------

    if product == "N/A":

        if "liability only" in t_lower:

            product = "Liability Only Policy For Private Car"

        elif "package policy" in t_lower:

            product = "Private Car Package Policy"

        elif "two wheeler policy" in t_lower:

            product = "two wheeler Policy"

        elif "comprehensive" in t_lower:

            product = "Comprehensive Motor Insurance Policy"

        elif "third party" in t_lower:

            product = "Third Party Motor Insurance Policy"


    # --------------------------------------------------------
    # 9. FINAL PRODUCT FALLBACK
    # --------------------------------------------------------

    if product == "N/A":

        product = "Motor Insurance Policy"


    product = re.sub(r"\s+", " ", product).strip()


    # ========================================================
    # SUM INSURED / IDV
    # ========================================================

    idv = find(
        r"(?:Total\s*IDV|Vehicle\s*IDV|IDV|"
        r"Insured\s*Declared\s*Value)"
        r"\s*[:\|]?\s*"
        r"(?:Rs\.?|₹)?\s*"
        r"([\d,]+(?:\.\d+)?)",
        t
    )


    # --------------------------------------------------------
    # IDV FALLBACKS
    # --------------------------------------------------------

    if idv == "N/A":

        idv = find(
            r"Total\s*IDV"
            r"\s*[:\|]?\s*"
            r"(?:Rs\.?|₹)?\s*"
            r"([\d,]+(?:\.\d+)?)",
            t
        )


    if idv == "N/A":

        idv = find(
            r"Vehicle\s*IDV"
            r"\s*[:\|]?\s*"
            r"(?:Rs\.?|₹)?\s*"
            r"([\d,]+(?:\.\d+)?)",
            t
        )


    idv = clean_amount(idv)


    # ========================================================
    # PREMIUM
    # ========================================================

    premium = find(
        r"(?:Final\s*Premium|Total\s*Amount|Gross\s*Premium)"
        r"\s*[:\-]?\s*"
        r"(?:Rs\.?|₹)?\s*"
        r"([\d,\.]+)",
        t
    )


    if premium == "N/A":

        premium = find(
            r"Final\s*Premium"
            r"\s*[:\-]?\s*"
            r"(?:Rs\.?|₹)?\s*"
            r"([\d,]+(?:\.\d+)?)",
            t
        )


    if premium == "N/A":

        premium = find(
            r"Total\s*Amount"
            r"\s*[:\-]?\s*"
            r"(?:Rs\.?|₹)?\s*"
            r"([\d,]+(?:\.\d+)?)",
            t
        )


    premium = clean_amount(premium)


    # ========================================================
    # INTERMEDIARY NAME
    # ========================================================

    intermediary = find(
        r"(?:Intermediary\s*Name|Intermediary|Agency\s*Name|"
        r"Agent\s*Name|Broker\s*Name)"
        r"\s*[:\|]?\s*"
        r"([A-Za-z0-9\s\.\&]+?)"
        r"(?=\s*(?:Email|Contact|Mobile|Phone|Sub|SP|$))",
        t
    )


    if intermediary != "N/A":

        intermediary = re.sub(
            r"\s+",
            " ",
            intermediary
        ).strip().title()


    # ========================================================
    # CUSTOMER MOBILE NUMBER
    # ========================================================

    mobile = find(
        r"Mobile\s*(?:Number|No\.?)"
        r"\s*[:\|]?\s*"
        r"([6-9]\d{9})",
        t
    )


    if mobile == "N/A":

        mobile = find(
            r"\b[6-9]\d{9}\b",
            t
        )


    # ========================================================
    # CUSTOMER EMAIL
    # ========================================================

    email = find(
        r"Email\s*(?:ID)?"
        r"\s*[:\|]?\s*"
        r"([A-Za-z0-9._%+\-]+"
        r"@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
        t
    )


    if email == "N/A":

        all_emails = re.findall(
            r"[A-Za-z0-9._%+\-]+"
            r"@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}",
            t,
            re.IGNORECASE
        )

        email = next(
            (
                e for e in all_emails
                if not re.search(
                    r"bajaj|care|support|info|admin",
                    e,
                    re.IGNORECASE
                )
            ),
            "N/A"
        )


    # ========================================================
    # FUEL TYPE
    # ========================================================

    fuel = "N/A"


    # --------------------------------------------------------
    # 1. Direct Fuel Type
    # --------------------------------------------------------

    fuel_match = re.search(
        r"Fuel\s*Type"
        r"\s*[:\|]?\s*"
        r"(PETROL(?:\s*\([A-Z]+\))?|"
        r"DIESEL(?:\s*\([A-Z]+\))?|"
        r"CNG(?:\s*\([A-Z]+\))?|"
        r"LPG(?:\s*\([A-Z]+\))?|"
        r"ELECTRIC|HYBRID)",
        t,
        re.IGNORECASE
    )


    if fuel_match:

        fuel = fuel_match.group(1).strip().upper()


    # --------------------------------------------------------
    # 2. Bajaj Vehicle Table Fallback
    # Example:
    # GJ-33-F-0282 MARUTI DZIRE VXI 1197 61 PETROL(P)
    # --------------------------------------------------------

    if fuel == "N/A":

        fuel_match = re.search(
            r"\b("
            r"PETROL\s*\([A-Z]+\)|"
            r"DIESEL\s*\([A-Z]+\)|"
            r"CNG\s*\([A-Z]+\)|"
            r"LPG\s*\([A-Z]+\)|"
            r"ELECTRIC|HYBRID"
            r")\b",
            t,
            re.IGNORECASE
        )

        if fuel_match:

            fuel = fuel_match.group(1).strip().upper()


    # --------------------------------------------------------
    # 3. Final Fuel Fallback
    # --------------------------------------------------------

    if fuel == "N/A":

        fuel_match = re.search(
            r"\b(PETROL|DIESEL|CNG|LPG|ELECTRIC|HYBRID)\b",
            t,
            re.IGNORECASE
        )

        if fuel_match:

            fuel = fuel_match.group(1).strip().upper()


    # ========================================================
    # REGISTRATION NUMBER
    # ========================================================

    reg_no = find(
        r"(?:Registration\s*Number|Registration\s*No\.?|"
        r"Regn\s*Number|Regn\s*No\.?|"
        r"Vehicle\s*Number|Vehicle\s*No\.?)"
        r"\s*[:\|]?\s*"
        r"([A-Z]{2}[\-\s]?\d{2}[\-\s]?"
        r"[A-Z]{1,3}[\-\s]?\d{4})",
        t
    )


    # --------------------------------------------------------
    # Registration fallback
    # --------------------------------------------------------

    if reg_no == "N/A":

        reg_no = find(
            r"\b([A-Z]{2}[\-\s]?\d{2}[\-\s]?"
            r"[A-Z]{1,2}[\-\s]?\d{4})\b",
            t
        )


    if reg_no != "N/A":

        reg_no = re.sub(r"\s+", "-", reg_no.upper())


    # ========================================================
    # ENGINE NUMBER
    # ========================================================

    engine = "N/A"
    match_eng = None


    # --------------------------------------------------------
    # 1. Direct Engine Number
    # --------------------------------------------------------

    engine = find(
        r"(?:Engine\s*Number|Engine\s*No\.?|Engine\s*#)"
        r"\s*[:\|]?\s*"
        r"([A-Z0-9]{8,25})",
        t
    )


    # --------------------------------------------------------
    # 2. Bajaj Vehicle Schedule Fallback
    # --------------------------------------------------------

    if engine == "N/A":

        match_eng = re.search(
            r"\b([A-HJ-NPR-Z0-9]{17})\b"
            r"\s+"
            r"([A-Z0-9]{8,25})\s+"
            r"(?:0|[\d,]+(?:\.\d+)?)\b",
            t,
            re.IGNORECASE
        )


        if match_eng:

            engine = match_eng.group(2).strip().upper()


    # --------------------------------------------------------
    # 3. Engine After Chassis
    # --------------------------------------------------------

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


    # ========================================================
    # CHASSIS NUMBER
    # ========================================================

    chassis = "N/A"
    match_chas = None


    # --------------------------------------------------------
    # 1. Direct Chassis Number
    # --------------------------------------------------------

    chassis = find(
        r"(?:Chassis\s*Number|Chassis\s*No\.?|Chassis\s*#)"
        r"\s*[:\|]?\s*"
        r"([A-Z0-9]{11,25})",
        t
    )


    # --------------------------------------------------------
    # 2. Standard 17 Character VIN
    # --------------------------------------------------------

    if chassis == "N/A":

        match_chas = re.search(
            r"\b([A-HJ-NPR-Z0-9]{17})\b",
            t,
            re.IGNORECASE
        )


        if match_chas:

            chassis = match_chas.group(1).strip().upper()


    # --------------------------------------------------------
    # 3. Bajaj Chassis + Engine Sequence
    # --------------------------------------------------------

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


    # ========================================================
    # VEHICLE INFO
    # ========================================================

    vehicle_info = "N/A"


    # Bajaj example:
    # GJ-33-F-0282 MARUTI DZIRE VXI 1197 61 PETROL(P)

    vehicle_match = re.search(
        r"\b[A-Z]{2}[\-\s]?\d{2}[\-\s]?"
        r"[A-Z]{1,2}[\-\s]?\d{4}\s+"
        r"(.+?)\s+"
        r"\d{2,5}\s+\d+\s+"
        r"(?:PETROL|DIESEL|CNG|ELECTRIC|HYBRID)",
        t,
        re.IGNORECASE
    )


    if vehicle_match:

        vehicle_info = vehicle_match.group(1).strip()


    # --------------------------------------------------------
    # Vehicle info cleanup
    # --------------------------------------------------------

    if vehicle_info != "N/A":

        vehicle_info = re.sub(
            r"\s+",
            " ",
            vehicle_info
        ).strip()


    # ========================================================
    # PAYMENT MODE
    # ========================================================

    if "online payment" in t_lower:

        pay_mode = "Online Payment"

    elif "cheque" in t_lower:

        pay_mode = "Cheque"

    elif "cash" in t_lower:

        pay_mode = "Cash"

    else:

        pay_mode = find(
            r"Instrument\s*Type"
            r"\s*[:\|]?\s*"
            r"([A-Za-z\s]+)",
            t
        )


    # ========================================================
    # RETURN DATA
    # ========================================================

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


# ============================================================
# MAIN PROCESSING
# ============================================================

if uploaded_files:

    all_data = []

    for file in uploaded_files:

        try:

            # ------------------------------------------------
            # READ PDF
            # ------------------------------------------------

            with pdfplumber.open(file) as pdf:

                text = " ".join(
                    page.extract_text() or ""
                    for page in pdf.pages
                )


            # ------------------------------------------------
            # EXTRACT POLICY
            # ------------------------------------------------

            data = extract_policy_details(
                text,
                file.name
            )


            data["File Name"] = file.name

            all_data.append(data)


        except Exception as e:

            st.error(
                f"❌ Error while processing "
                f"{file.name}: {str(e)}"
            )


    # ========================================================
    # CREATE DATAFRAME
    # ========================================================

    if all_data:

        df = pd.DataFrame(
            all_data,
            columns=columns
        ).fillna("N/A")


        # ====================================================
        # ADD ACCESSORY VALUE TO IDV
        # ====================================================

        if accessory_value > 0:

            def update_idv(idv_str):

                if idv_str == "N/A":

                    return f"{accessory_value:,.2f}"

                try:

                    base_value = float(
                        str(idv_str)
                        .replace(",", "")
                        .replace("₹", "")
                        .strip()
                    )

                    updated_value = (
                        base_value + accessory_value
                    )

                    return f"{updated_value:,.2f}"

                except Exception:

                    return idv_str


            df["Sum Insured / IDV"] = (
                df["Sum Insured / IDV"]
                .apply(update_idv)
            )


            st.sidebar.success(
                f"✅ Sum Insured / IDV updated "
                f"by ₹{accessory_value:,.2f}"
            )


        # ====================================================
        # DISPLAY DATA
        # ====================================================

        st.success(
            "✅ Extraction complete! Review below:"
        )

        st.dataframe(
            df,
            use_container_width=True
        )


        # ====================================================
        # DOWNLOAD EXCEL
        # ====================================================

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="xlsxwriter"
        ) as writer:

            df.to_excel(
                writer,
                index=False,
                sheet_name="Policy Details"
            )


        st.download_button(
            label="📥 Download Extracted Policy Data (Excel)",
            data=output.getvalue(),
            file_name="bajaj_policy_extracted_data.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.caption(
    "Built with 💙 Streamlit + pdfplumber + Regex | "
    "Tailored specifically for Bajaj General Insurance Policy PDFs"
)