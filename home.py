import streamlit as st
import os

# --- Streamlit Page Config (only once) ---
st.set_page_config(page_title="📄 Multi-Company Policy Extractor", layout="centered")

st.title("📄 Multi-Company Policy Extractor → Excel")
st.write("Extract policy details from PDFs for **Tata AIG**, **Royal Sundaram**, **Reliance**, **Zurich Kotak**, and **National** and export them to Excel format.")

# --- Sidebar: Company Selector ---
st.sidebar.header("🏢 Choose Insurance Company")
company = st.sidebar.selectbox(
    "Select the Company",
    ["Tata AIG", "Royal Sundaram", "Reliance", "Zurich Kotak","National"]
)

# --- Map Company to Script File ---
company_scripts = {
    "Tata AIG": "tata.py",
    "Royal Sundaram": "royal.py",
    "Reliance": "reliance.py",
    "Zurich Kotak": "kotak.py",
    "National": "national.py"
}

selected_script = company_scripts[company]

st.sidebar.markdown("---")
st.sidebar.info(f"👆 Selected: {company}")

# --- Load and Display ---
st.markdown("---")

if os.path.exists(selected_script):
    st.success(f"Running extractor for **{company}**")
    # Read the file content
    with open(selected_script, "r", encoding="utf-8") as f:
        code = f.read()

    # Remove duplicate Streamlit config lines to avoid set_page_config() error
    cleaned_lines = []
    for line in code.splitlines():
        if "st.set_page_config" in line:
            continue
        cleaned_lines.append(line)
    cleaned_code = "\n".join(cleaned_lines)

    # Execute safely inside this Streamlit session
    try:
        exec(cleaned_code, globals())
    except Exception as e:
        st.error(f"❌ Error while running {company} extractor: {e}")
else:
    st.error(f"Extractor file not found: {selected_script}")


