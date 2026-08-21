import streamlit as st
import pandas as pd
import numpy as np

# Sample Data
sample_data = pd.DataFrame([
    {
        "Timestamp": "10:42:15 AM", 
        "Company": "Tata AIG", 
        "Files Processed": 4, 
        "Accuracy": 0.98,
        "Trend": [10, 12, 14, 18],
        "Status": "Success", 
        "Export Size": "24 KB",
        "Action": "https://example.com/download/tata"
    },
    {
        "Timestamp": "11:15:02 AM", 
        "Company": "Zurich Kotak", 
        "Files Processed": 2, 
        "Accuracy": 0.85,
        "Trend": [5, 6, 4, 8],
        "Status": "Warning", 
        "Export Size": "12 KB",
        "Action": "https://example.com/download/kotak"
    },
    {
        "Timestamp": "02:04:30 PM", 
        "Company": "Royal Sundaram", 
        "Files Processed": 8, 
        "Accuracy": 1.00,
        "Trend": [8, 15, 22, 30],
        "Status": "Success", 
        "Export Size": "48 KB",
        "Action": "https://example.com/download/royal"
    },
])

# Advanced Table Configuration
st.dataframe(
    sample_data,
    column_config={
        "Company": st.column_config.TextColumn(
            "Insurance Company",
            help="Configured processing engine profile"
        ),
        "Accuracy": st.column_config.ProgressColumn(
            "Confidence Score",
            help="Extraction confidence level",
            format="%.0f%%",
            min_value=0,
            max_value=1,
        ),
        "Trend": st.column_config.LineChartColumn(
            "Volume Trend",
            help="Recent file processing velocity"
        ),
        "Action": st.column_config.LinkColumn(
            "Report Link",
            display_text="Download Excel"
        ),
        "Status": st.column_config.SelectboxColumn(
            "Execution Status",
            options=["Success", "Warning", "Failed"],
            required=True
        )
    },
    use_container_width=True,
    hide_index=True
)
