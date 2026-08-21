import streamlit as st
import os
import importlib.util
import time

# --- Page Config ---
st.set_page_config(
    page_title="Policy AI - Multi-Company Extractor",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Modern Custom UI Styling ---
st.markdown("""
    <style>
    /* Main Layout Styling */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }
    
    /* Header Gradient Banner */
    .header-card {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        border-radius: 16px;
        padding: 24px 32px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
    }
    
    .header-card h2 {
        margin: 0;
        font-weight: 700;
        letter-spacing: -0.5px;
        color: #FFFFFF !important;
    }
    
    .header-card p {
        color: #94A3B8;
        margin-top: 6px;
        margin-bottom: 0;
        font-size: 0.95rem;
    }

    /* Active Provider Badge */
    .provider-badge {
        display: inline-block;
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.85rem;
        margin-top: 10px;
    }

    /* Sidebar Refinements */
    section[data-testid="stSidebar"] {
        background-color: #F8FAFC;
        border-right: 1px solid #E2E8F0;
    }
    
    /* Utility Container Styling */
    .info-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    </style>
""", unsafe_allow_html=True)

# --- Configuration & Brand Mapping ---
COMPANY_CONFIG = {
    "Tata AIG": {
        "script": "tata.py", 
        "color": "#1E40AF", 
        "bg": "#EFF6FF", 
        "fields": ["Policy No", "Insured Name", "Insured Amount", "Expiry Date"]
    },
    "Royal Sundaram": {
        "script": "royal.py", 
        "color": "#0369A1", 
        "bg": "#F0F9FF", 
        "fields": ["Policy No", "Vehicle No", "Premium Amount", "NCB %"]
    },
    "Zurich Kotak": {
        "script": "kotak.py", 
        "color": "#B91C1C", 
        "bg": "#FEF2F2", 
        "fields": ["Policy No", "Customer ID", "Gross Premium", "Agent Code"]
    },
    "National": {
        "script": "national.py", 
        "color": "#047857", 
        "bg": "#ECFDF5", 
        "fields": ["Policy No", "Office Code", "Sum Insured", "GST Details"]
    }
}

# --- Sidebar Controls ---
with st.sidebar:
    st.title("⚡ Engine Hub")
    st.caption("Select your extraction profile below")
    
    company = st.selectbox(
        "Insurance Provider",
        list(COMPANY_CONFIG.keys()),
        index=0
    )
    
    selected_cfg = COMPANY_CONFIG[company]
    
    st.markdown("---")
    
    # Active Provider Card
    st.markdown(
        f"""
        <div style="background-color: {selected_cfg['bg']}; border: 1px solid {selected_cfg['color']}30; padding: 16px; border-radius: 12px;">
            <span style="font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.5px; color: #64748B; font-weight: 600;">Selected Engine</span>
            <h4 style="color: {selected_cfg['color']}; margin: 4px 0 0 0; font-weight: 700;">{company}</h4>
            <p style="font-size: 0.8rem; color: #475569; margin-top: 8px;">Target Script: <code>{selected_cfg['script']}</code></p>
        </div>
        """, 
        unsafe_allow_html=True
    )
    
    st.markdown("---")
    st.markdown("##### ⚙️ System Status")
    st.caption("🟢 Dynamic Parser Engine: **Ready**")
    st.caption("🟢 Excel Export Engine: **Active**")

# --- Main Interface ---
# Top Header Banner
st.markdown(
    f"""
    <div class="header-card">
        <h2>📄 Insurance Policy Data Extractor</h2>
        <p>Convert structured and semi-structured PDF policy documents into standardized Excel reports.</p>
        <div class="provider-badge" style="background-color: {selected_cfg['color']}20; color: {selected_cfg['color']}; border: 1px solid {selected_cfg['color']}40;">
            Active Engine: {company}
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# Workspace Navigation Tabs
tab_extract, tab_fields, tab_history = st.tabs([
    "🚀 Processing Console", 
    "📋 Field Schemas", 
    "🕒 Recent Activity"
])

# --- TAB 1: Main Extractor Console ---
with tab_extract:
    col_upload, col_summary = st.columns([2, 1])
    
    with col_upload:
        st.subheader("Upload PDF Documents")
        uploaded_files = st.file_uploader(
            f"Drop files for {company}",
            type=["pdf"],
            accept_multiple_files=True,
            help="Supported document formats: Native digital PDFs"
        )
    
    with col_summary:
        st.subheader("Batch Summary")
        file_count = len(uploaded_files) if uploaded_files else 0
        
        m_col1, m_col2 = st.columns(2)
        m_col1.metric("Selected Engine", company)
        m_col2.metric("Queued Files", file_count)
        
        if uploaded_files:
            st.success("✅ Files queued and ready for processing.")
        else:
            st.info("ℹ️ Upload one or more PDFs to unlock extraction.")

    st.markdown("---")

    # --- Script Loader Logic ---
    script_path = selected_cfg["script"]

    if os.path.exists(script_path):
        try:
            # Safely import the module without polluting global scope
            spec = importlib.util.spec_from_file_location("company_module", script_path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)

            # Execution hook
            if hasattr(module, "main"):
                module.main(uploaded_files)
            elif hasattr(module, "extract"):
                module.extract(uploaded_files)
            else:
                st.warning(f"⚠️ `{script_path}` loaded, but no standard `main()` or `extract()` entry point was detected.")

        except Exception as e:
            st.error(f"❌ Execution error inside `{script_path}`:")
            st.exception(e)
    else:
        st.error(f"❌ Extractor script missing: Could not find `{script_path}` in working directory.")

# --- TAB 2: Schema / Field Mapping ---
with tab_fields:
    st.subheader(f"Data Schema for {company}")
    st.write("This provider engine is configured to automatically parse and format the following metadata attributes:")
    
    cols = st.columns(len(selected_cfg["fields"]))
    for idx, field in enumerate(selected_cfg["fields"]):
        with cols[idx]:
            st.markdown(
                f"""
                <div style="background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 14px; border-radius: 8px; text-align: center;">
                    <span style="font-size: 0.8rem; color: #64748B;">Target Field</span><br>
                    <strong style="color: #0F172A;">{field}</strong>
                </div>
                """, 
                unsafe_allow_html=True
            )

# --- TAB 3: Recent Activity (Placeholder UI) ---
with tab_history:
    st.subheader("Session History")
    st.caption("Track previously generated reports from this session.")
    
    st.dataframe(
        [
            {"Timestamp": "10:42 AM", "Engine": "Tata AIG", "Files Processed": 4, "Status": "Completed"},
            {"Timestamp": "11:15 AM", "Engine": "Zurich Kotak", "Files Processed": 2, "Status": "Completed"},
        ],
        use_container_width=True
    )
