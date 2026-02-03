import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
from PIL import Image
import io
import datetime
import xlsxwriter
import plotly.express as px

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="AirSial SMS Digitizer",
    page_icon="✈️",
    layout="wide"
)

# --- SIDEBAR & API SETUP ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/en/thumb/8/8e/AirSial_Logo.svg/1200px-AirSial_Logo.svg.png", width=200)
    st.title("⚙️ Configuration")
    api_key = st.text_input("Enter Google Gemini API Key", type="password", help="Get your free key from aistudio.google.com")
    
    st.info("ℹ️ **Privacy Note:** Data is processed in memory and deleted immediately after you close this tab. No data is stored.")

# --- GOOGLE GEMINI SETUP ---
def configure_gemini(api_key):
    genai.configure(api_key=api_key)
    return genai.GenerativeModel('gemini-1.5-flash')

# --- PROCESSING FUNCTION ---
def process_image(model, image_file):
    img = Image.open(image_file)
    
    prompt = """
    Analyze this AirSial Hazard Identification & Risk Assessment Form. 
    Extract data into a valid JSON object. 
    If a field is empty or unclear, use "N/A".

    Return strictly this JSON structure:
    {
        "report_no": "String",
        "date_of_report": "DD-MM-YYYY",
        "reporter_name": "String",
        "department": "String",
        "location": "String",
        "hazard_description": "String",
        "severity_initial": "String",
        "probability_initial": "String",
        "risk_level_initial": "String",
        "cap_required": "Yes/No",
        "cap_action_plan": "String",
        "responsible_person": "String",
        "target_date": "DD-MM-YYYY",
        "cap_attachment": "Yes/No",
        "residual_severity": "String",
        "residual_probability": "String",
        "residual_risk_level": "String",
        "wet_lease_involved": "Yes/No",
        "operator_name": "String",
        "report_attached": "Yes/No"
    }
    """
    
    # default_error structure ensures columns exist even if AI fails
    default_data = {
        "report_no": f"Error-{image_file.name}",
        "date_of_report": "N/A", "reporter_name": "N/A", "department": "N/A",
        "location": "N/A", "hazard_description": "Extraction Failed", 
        "severity_initial": "N/A", "probability_initial": "N/A", 
        "risk_level_initial": "N/A", "cap_required": "No", 
        "cap_action_plan": "N/A", "responsible_person": "N/A", 
        "target_date": "N/A", "cap_attachment": "No", 
        "residual_severity": "N/A", "residual_probability": "N/A", 
        "residual_risk_level": "N/A", "wet_lease_involved": "No", 
        "operator_name": "N/A", "report_attached": "No"
    }

    try:
        response = model.generate_content([prompt, img])
        # Clean up code blocks
        text = response.text
        if "```json" in text:
            text = text.split("```json")[1].split("```")[0]
        elif "```" in text:
            text = text.split("```")[1]
            
        data = json.loads(text)
        return data
    except Exception as e:
        default_data["hazard_description"] = f"Error: {str(e)}"
        return default_data

# --- EXCEL GENERATION ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1f4e78', 'font_color': 'white', 'border': 1})
        
        # Ensure all expected columns exist before writing
        expected_cols = [
            "report_no", "date_of_report", "reporter_name", "department", "location",
            "hazard_description", "severity_initial", "probability_initial", "risk_level_initial",
            "cap_required", "cap_action_plan", "responsible_person", "target_date",
            "cap_attachment", "residual_severity", "residual_probability", "residual_risk_level",
            "wet_lease_involved", "operator_name", "report_attached"
        ]
        
        # Fill missing columns with "N/A" to prevent KeyError
        for col in expected_cols:
            if col not in df.columns:
                df[col] = "N/A"
                
        # Write Sheet 1
        df.to_excel(writer, sheet_name='Raw SMS Data', index=False)
        worksheet = writer.sheets['Raw SMS Data']
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # Write Sheet 2 (CAP)
        cap_cols = ['report_no', 'cap_action_plan', 'responsible_person', 'target_date', 'cap_required']
        # Filter only existing columns
        existing_cap_cols = [c for c in cap_cols if c in df.columns]
        cap_df = df[existing_cap_cols].copy()
        cap_df.to_excel(writer, sheet_name='CAP Tracker', index=False)
        
    return output.getvalue()

# --- DASHBOARD GENERATION ---
def generate_dashboard(df):
    st.markdown("---")
    st.header("📊 Monthly SMS Dashboard (Preview)")
    
    # SAFEGUARD: Ensure columns exist
    required_cols = ['risk_level_initial', 'wet_lease_involved', 'cap_required', 'location', 'severity_initial']
    for col in required_cols:
        if col not in df.columns:
            df[col] = "N/A"

    # 1. KPI TILES
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Hazards", len(df))
    with col2:
        # Safe filtering
        high_risk = len(df[df['risk_level_initial'].astype(str).str.contains('High', case=False, na=False)])
        st.metric("High Risk Hazards", high_risk)
    with col3:
        wet_lease = len(df[df['wet_lease_involved'].astype(str).str.contains('Yes', case=False, na=False)])
        st.metric("Wet Lease Incidents", wet_lease)
    with col4:
        pending_cap = len(df[df['cap_required'].astype(str).str.contains('Yes', case=False, na=False)])
        st.metric("CAPs Identified", pending_cap)

    # 2. CHARTS
    c1, c2 = st.columns(2)
    
    with c1:
        st.subheader("Hazards by Location")
        if not df.empty:
            # Clean data for chart
            loc_counts = df['location'].value_counts().reset_index()
            loc_counts.columns = ['location', 'count']
            fig_loc = px.bar(loc_counts, x='location', y='count', title="Location Frequency")
            st.plotly_chart(fig_loc, use_container_width=True)
            
    with c2:
        st.subheader("Risk Severity")
        if not df.empty:
            sev_counts = df['severity_initial'].value_counts().reset_index()
            sev_counts.columns = ['severity', 'count']
            fig_sev = px.pie(sev_counts, values='count', names='severity', title="Initial Severity Levels", hole=0.3)
            st.plotly_chart(fig_sev, use_container_width=True)

# --- MAIN APP LOGIC ---
st.title("🛫 AirSial SMS Digitizer & Dashboard")
st.markdown("Upload scans of **AS-SMS-003** forms.")

uploaded_files = st.file_uploader("Upload Reports", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])

if st.button("🚀 Extract Data"):
    if not api_key:
        st.error("Please enter your API Key in the sidebar.")
    elif not uploaded_files:
        st.warning("Please upload at least one image.")
    else:
        model = configure_gemini(api_key)
        results = []
        
        my_bar = st.progress(0, text="Starting AI Scan...")
        
        for i, file in enumerate(uploaded_files):
            data = process_image(model, file)
            results.append(data)
            my_bar.progress(int(((i + 1) / len(uploaded_files)) * 100), text=f"Scanning {file.name}...")
        
        my_bar.empty()
        st.success("✅ Extraction Complete!")
        
        # Create DataFrame
        df = pd.DataFrame(results)
        
        # --- CRITICAL FIX: REORDER & FILL MISSING COLUMNS ---
        expected_cols = [
            "report_no", "date_of_report", "reporter_name", "department", "location",
            "hazard_description", "severity_initial", "probability_initial", "risk_level_initial",
            "cap_required", "cap_action_plan", "responsible_person", "target_date",
            "cap_attachment", "residual_severity", "residual_probability", "residual_risk_level",
            "wet_lease_involved", "operator_name", "report_attached"
        ]
        
        # Ensure all columns exist, fill missing with "N/A"
        for col in expected_cols:
            if col not in df.columns:
                df[col] = "N/A"
        
        # Reorder columns to look clean
        df = df[expected_cols]

        # Show Data
        st.dataframe(df)
        
        # Dashboard
        generate_dashboard(df)
        
        # Excel
        excel_data = to_excel(df)
        st.download_button(
            label="📥 Download Excel",
            data=excel_data,
            file_name=f"AirSial_SMS_Log_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
