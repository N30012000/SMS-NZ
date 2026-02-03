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
    st.info("ℹ️ **Privacy:** Data is processed in-memory and deleted after use.")

# --- GEMINI AI SETUP ---
def configure_gemini(api_key):
    genai.configure(api_key=api_key)
    return genai.GenerativeModel('gemini-1.5-flash')

# --- INTELLIGENT OCR & PARSING ---
def process_image(model, image_file):
    img = Image.open(image_file)
    
    prompt = """
    Analyze this AirSial Hazard Identification & Risk Assessment Form (AS-SMS-003).
    Extract all data into a strictly valid JSON format.
    
    CRITICAL INSTRUCTIONS:
    1. If a field is handwritten, transcribe it exactly.
    2. If a field is empty or illegible, use "N/A".
    3. For 'Risk Level', look for the ticked box in the matrix.
    4. Normalize dates to DD-MM-YYYY.

    Return JSON with these exact keys:
    {
        "report_no": "String",
        "date_of_report": "DD-MM-YYYY",
        "location": "String",
        "department": "String",
        "hazard_description": "String",
        "severity_initial": "String",
        "probability_initial": "String",
        "risk_level_initial": "String",
        "cap_required": "Yes/No",
        "cap_action_plan": "String",
        "responsible_person": "String",
        "target_date": "DD-MM-YYYY",
        "wet_lease_involved": "Yes/No",
        "operator_name": "String",
        "report_attached": "Yes/No"
    }
    """
    
    # Default fallback data to prevent KeyErrors if AI fails
    default_data = {
        "report_no": f"Error-{image_file.name}", "date_of_report": "N/A", 
        "location": "N/A", "department": "N/A", "hazard_description": "Extraction Failed",
        "severity_initial": "N/A", "probability_initial": "N/A", "risk_level_initial": "N/A",
        "cap_required": "No", "cap_action_plan": "N/A", "responsible_person": "N/A",
        "target_date": "N/A", "wet_lease_involved": "No", "operator_name": "N/A", 
        "report_attached": "No"
    }

    try:
        response = model.generate_content([prompt, img])
        text = response.text
        # Clean markdown formatting
        if "```json" in text:
            text = text.split("```json")[1].split("```")[0]
        elif "```" in text:
            text = text.split("```")[1]
        
        data = json.loads(text)
        # Merge with default to ensure all keys exist
        return {**default_data, **data}
    except Exception as e:
        default_data["hazard_description"] = f"AI Error: {str(e)}"
        return default_data

# --- EXCEL GENERATOR (AUDIT READY) ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1f4e78', 'font_color': 'white', 'border': 1})
        
        # 1. Raw Data Sheet
        df.to_excel(writer, sheet_name='Raw SMS Data', index=False)
        worksheet = writer.sheets['Raw SMS Data']
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
            
        # 2. CAP Tracker Sheet
        if not df.empty:
            cap_df = df[['report_no', 'cap_action_plan', 'responsible_person', 'target_date', 'cap_required']].copy()
            cap_df.to_excel(writer, sheet_name='CAP Tracker', index=False)
            ws_cap = writer.sheets['CAP Tracker']
            for col_num, value in enumerate(cap_df.columns.values):
                ws_cap.write(0, col_num, value, header_fmt)

    return output.getvalue()

# --- DASHBOARD GENERATOR ---
def generate_dashboard(df):
    st.markdown("---")
    st.header("📊 Monthly SMS Dashboard")
    
    # Helper to safe-get counts
    def safe_count(column, value_substring):
        if column not in df.columns: return 0
        return len(df[df[column].astype(str).str.contains(value_substring, case=False, na=False)])

    # 1. KPI Cards
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Reports", len(df))
    k2.metric("High Risk", safe_count('risk_level_initial', 'High'))
    k3.metric("Wet Lease Incidents", safe_count('wet_lease_involved', 'Yes'))
    k4.metric("CAPs Pending", safe_count('cap_required', 'Yes'))

    # 2. Charts
    c1, c2 = st.columns(2)
    
    with c1:
        st.subheader("📍 Hazards by Location")
        if 'location' in df.columns:
            loc_counts = df['location'].value_counts().reset_index()
            loc_counts.columns = ['Location', 'Count']
            fig = px.bar(loc_counts, x='Location', y='Count', color='Location')
            st.plotly_chart(fig, use_container_width=True)
            
    with c2:
        st.subheader("⚠️ Risk Severity")
        if 'severity_initial' in df.columns:
            sev_counts = df['severity_initial'].value_counts().reset_index()
            sev_counts.columns = ['Severity', 'Count']
            fig2 = px.pie(sev_counts, values='Count', names='Severity', hole=0.4)
            st.plotly_chart(fig2, use_container_width=True)

# --- MAIN APP LOGIC ---
st.title("🛫 AirSial SMS Digitizer & Dashboard")
st.write("Upload **AS-SMS-003** forms (Images). The AI will digitize handwriting, check boxes, and generate your dashboard.")

uploaded_files = st.file_uploader("Upload Report Images", accept_multiple_files=True, type=['jpg', 'png', 'jpeg'])

if st.button("🚀 Process Reports"):
    if not api_key:
        st.error("❌ Please enter your Google Gemini API Key in the sidebar.")
    elif not uploaded_files:
        st.warning("⚠️ Please upload at least one file.")
    else:
        model = configure_gemini(api_key)
        results = []
        
        # Progress Bar
        bar = st.progress(0, text="Initializing AI...")
        
        for i, file in enumerate(uploaded_files):
            # Process
            data = process_image(model, file)
            results.append(data)
            # Update bar
            bar.progress(int(((i + 1) / len(uploaded_files)) * 100), text=f"Scanning {file.name}...")
        
        bar.empty()
        st.success("✅ Extraction Complete!")
        
        # Create DataFrame
        df = pd.DataFrame(results)
        
        # Display Dashboard
        generate_dashboard(df)
        
        # Display Data Table
        with st.expander("📄 View Raw Data"):
            st.dataframe(df)
            
        # Download Button
        excel_data = to_excel(df)
        st.download_button(
            label="📥 Download Audit-Ready Excel",
            data=excel_data,
            file_name=f"AirSial_SMS_Log_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
