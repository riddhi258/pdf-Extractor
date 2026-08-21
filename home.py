Here is the simplified code. The hero banner, redundant badges, and explicit file uploader call have been removed from the main page script.

The layout now directly invokes `run_extractor_engine()`, leaving execution entirely to your designated module script (e.g., `tata.py`):

```python
import streamlit as st
import os
import importlib.util
import time
import pandas as pd

# --- Page Setup ---
st.set_page_config(
    page_title="InsurData AI | Policy Extraction Platform",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Executive Custom CSS ---
st.markdown("""
<style>
    /* Google Fonts Import */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Main Container Padding */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1350px;
    }

    /* Hide Default Streamlit Chrome */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Sidebar Refinement */
    section[data-testid="stSidebar"] {
        background-color: #0B132B;
        border-right: 1px solid #1E293B;
    }
    
    section[data-testid="stSidebar"] * {
        color: #F8FAFC !important;
    }

    section[data-testid="stSidebar"] .stSelectbox label {
        color: #94A3B8 !important;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# --- Configuration Directory ---
COMPANY_CONFIG = {
    "Tata AIG": {
        "script": "tata.py",
        "primary": "#2563EB",
        "light_bg": "#EFF6FF",
        "text": "#1E40AF",
        "icon": "🛡️",
        "fields": ["Policy No", "Insured Name", "Insured Amount", "Expiry Date", "Premium Total"]
    },
    "Royal Sundaram": {
        "script": "royal.py",
        "primary": "#0284C7",
        "light_bg": "#F0F9FF",
        "text": "#0369A1",
        "icon": "🦁",
        "fields": ["Policy No", "Vehicle Registration", "IDV Value", "NCB %", "GST Tax"]
    },
    "Zurich Kotak": {
        "script": "kotak.py",
        "primary": "#DC2626",
        "light_bg": "#FEF2F2",
        "text": "#B91C1C",
        "icon": "🟥",
        "fields": ["Policy No", "Customer ID", "Gross Premium", "Agent Code", "Term Duration"]
    },
    "National": {
        "script": "national.py",
        "primary": "#059669",
        "light_bg": "#ECFDF5",
        "text": "#047857",
        "icon": "🏛️",
        "fields": ["Policy No", "Branch Code", "Sum Insured", "Nominee Details", "Issuance Date"]
    }
}

# --- Engine Mediator Component ---
def run_extractor_engine(script_path: str):
    """
    Mediator component responsible for safely importing and executing 
    dynamic company extraction modules.
    """
    if not os.path.exists(script_path):
        st.warning(f"⚠️ Missing Extractor Module: Script file `{script_path}` was not found in the current root folder.")
        return

    try:
        # Dynamic import resolution
        module_name = os.path.splitext(os.path.basename(script_path))[0]
        spec = importlib.util.spec_from_file_location(module_name, script_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)

        # Execution router checking standard entry points
        if hasattr(module, "main"):
            module.main()
        elif hasattr(module, "extract"):
            module.extract()
        elif hasattr(module, "run"):
            module.run()
        else:
            st.info(
                f"ℹ️ Engine `{script_path}` loaded successfully. "
                f"Implement a `main()`, `extract()`, or `run()` function in your script."
            )

    except Exception as e:
        st.error(f"❌ Execution Exception inside `{script_path}`:")
        st.exception(e)

# --- Sidebar Component ---
with st.sidebar:
    st.markdown("### ✨ InsurData AI")
    st.caption("Enterprise Policy Parsing Studio")
    st.markdown("---")
    
    company = st.selectbox(
        "Select Extraction Engine",
        list(COMPANY_CONFIG.keys()),
        index=0
    )
    
    selected_cfg = COMPANY_CONFIG[company]
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Active Engine Card in Sidebar
    st.markdown(
        f"""
        <div style="background: rgba(255, 255, 255, 0.05); border: 1px solid rgba(255, 255, 255, 0.1); padding: 16px; border-radius: 12px;">
            <div style="font-size: 0.7rem; color: #94A3B8; font-weight: 700; text-transform: uppercase;">Active Driver</div>
            <div style="font-size: 1.1rem; font-weight: 700; margin-top: 4px; color: #FFFFFF;">
                {selected_cfg['icon']} {company}
            </div>
            <div style="font-size: 0.75rem; color: #64748B; margin-top: 8px;">
                Mapped File: <code style="color: #38BDF8; background: transparent;">{selected_cfg['script']}</code>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("##### ⚡ Core Engines")
    st.caption("🟢 OCR Parser v4.2: **Active**")
    st.caption("🟢 Dynamic Excel Generator: **Ready**")

# --- Main Extractor Call ---
run_extractor_engine(selected_cfg["script"])

```
