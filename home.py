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
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1350px;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .hero-banner {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 60%, #0284C7 100%);
        border-radius: 20px;
        padding: 32px 36px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 20px 25px -5px rgba(15, 23, 42, 0.15);
        position: relative;
        overflow: hidden;
    }

    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        letter-spacing: -0.03em;
        margin: 0;
        color: #FFFFFF !important;
        line-height: 1.2;
    }

    .hero-subtitle {
        color: #94A3B8;
        font-size: 1rem;
        margin-top: 8px;
        margin-bottom: 0;
        font-weight: 400;
    }

    .glass-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02);
    }

    .card-label {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #64748B;
    }

    .card-value {
        font-size: 1.5rem;
        font-weight: 800;
        color: #0F172A;
        margin-top: 4px;
    }

    .badge {
        display: inline-flex;
        align-items: center;
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 700;
        margin-top: 14px;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 1px solid #E2E8F0;
        padding-bottom: 4px;
    }

    .stTabs [data-baseweb="tab"] {
        height: 44px;
        border-radius: 10px;
        font-weight: 600;
        font-size: 0.9rem;
        color: #64748B;
        padding: 0 16px;
    }

    .stTabs [aria-selected="true"] {
        background-color: #F1F5F9 !important;
        color: #0F172A !important;
    }

    section[data-testid="stSidebar"] {
        background-color: #0B132B;
        border-right: 1px solid #1E293B;
    }
    
    section[data-testid="stSidebar"] * {
        color: #F8FAFC !important;
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

# --- Sidebar ---
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

# --- Header Banner ---
st.markdown(
    f"""
    <div class="hero-banner">
        <h1 class="hero-title">Insurance Policy Intelligence Platform</h1>
        <p class="hero-subtitle">Automated high-accuracy PDF extraction & key-value schema standardizer.</p>
        <div class="badge" style="background-color: rgba(255,255,255,0.15); color: #FFFFFF; border: 1px solid rgba(255,255,255,0.2);">
            Active Engine: {company} Profile
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# --- TOP LEVEL SINGLE FILE UPLOADER ---
st.subheader("📂 Document Upload Zone")
uploaded_files = st.file_uploader(
    f"Upload PDF policies for {company}",
    type=["pdf"],
    accept_multiple_files=True,
    help="Upload native digital PDF policy documents here once. All modules below will use these files."
)

file_count = len(uploaded_files) if uploaded_files else 0

st.markdown("<br>", unsafe_allow_html=True)

# --- Top Metrics Row ---
col_m1, col_m2, col_m3 = st.columns(3)

with col_m1:
    st.markdown(
        f"""
        <div class="glass-card">
            <div class="card-label">Active Provider</div>
            <div class="card-value" style="color: {selected_cfg['primary']};">{company}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with col_m2:
    st.markdown(
        f"""
        <div class="glass-card">
            <div class="card-label">Queued Documents</div>
            <div class="card-value">{file_count} File(s)</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with col_m3:
    st.markdown(
        f"""
        <div class="glass-card">
            <div class="card-label">Schema Parameters</div>
            <div class="card-value">{len(selected_cfg['fields'])} Fields</div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("<br>", unsafe_allow_html=True)

# --- Navigation Tabs ---
tab_console, tab_schema, tab_audit = st.tabs([
    "🚀 Processing Console", 
    "📐 Schema Inspector", 
    "📊 Batch Analytics & Audit"
])

# --- TAB 1: Main Extractor Studio ---
with tab_console:
    st.subheader(f"Engine Console: {company}")
    
    if not uploaded_files:
        st.info("👆 Please upload policy PDFs in the upload zone above to unlock processing.")
    else:
        st.success(f"✅ {file_count} file(s) loaded and ready for execution.")
        
    st.markdown("---")

    # Safe Module Execution
    script_path = selected_cfg["script"]

    if os.path.exists(script_path):
        try:
            spec = importlib.util.spec_from_file_location("company_module", script_path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)

            # Execution call
            if hasattr(module, "main"):
                module.main(uploaded_files)
            elif hasattr(module, "extract"):
                module.extract(uploaded_files)
            else:
                st.info(f"ℹ️ Engine `{script_path}` ready. Ensure it contains a `main(uploaded_files)` function.")

        except Exception as e:
            st.error(f"❌ Execution error inside `{script_path}`:")
            st.exception(e)
    else:
        st.warning(f"⚠️ Missing script file: `{script_path}` not found in root directory.")

# --- TAB 2: Schema Inspector ---
with tab_schema:
    st.subheader(f"Configured Schema: {company}")
    st.caption("The engine automatically parses and standardizes these target parameters:")
    
    cols = st.columns(len(selected_cfg["fields"]))
    for idx, field in enumerate(selected_cfg["fields"]):
        with cols[idx]:
            st.markdown(
                f"""
                <div style="background-color: {selected_cfg['light_bg']}; border: 1px solid {selected_cfg['primary']}20; padding: 16px; border-radius: 12px; text-align: center;">
                    <div style="font-size: 0.7rem; color: {selected_cfg['text']}; font-weight: 700; text-transform: uppercase;">Field {idx+1}</div>
                    <div style="font-size: 0.95rem; font-weight: 700; color: #0F172A; margin-top: 4px;">{field}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

# --- TAB 3: Batch Audit Logs ---
with tab_audit:
    st.subheader("Session Activity Log")
    
    sample_data = pd.DataFrame([
        {
            "Timestamp": "10:42:15 AM", 
            "Company": "Tata AIG", 
            "Files Processed": 4, 
            "Accuracy": 0.98,
            "Trend": [10, 12, 14, 18],
            "Status": "Success", 
            "Action": "https://example.com/download/tata"
        },
        {
            "Timestamp": "11:15:02 AM", 
            "Company": "Zurich Kotak", 
            "Files Processed": 2, 
            "Accuracy": 0.85,
            "Trend": [5, 6, 4, 8],
            "Status": "Warning", 
            "Action": "https://example.com/download/kotak"
        },
        {
            "Timestamp": "02:04:30 PM", 
            "Company": "Royal Sundaram", 
            "Files Processed": 8, 
            "Accuracy": 1.00,
            "Trend": [8, 15, 22, 30],
            "Status": "Success", 
            "Action": "https://example.com/download/royal"
        },
    ])

    st.dataframe(
        sample_data,
        column_config={
            "Company": st.column_config.TextColumn("Provider"),
            "Accuracy": st.column_config.ProgressColumn("Confidence Score", format="%.0f%%", min_value=0, max_value=1),
            "Trend": st.column_config.LineChartColumn("Volume Trend"),
            "Action": st.column_config.LinkColumn("Report Link", display_text="Download Excel"),
            "Status": st.column_config.SelectboxColumn("Status", options=["Success", "Warning", "Failed"])
        },
        use_container_width=True,
        hide_index=True
    )
