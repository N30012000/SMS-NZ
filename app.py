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
    # Using Flash for speed and cost-efficiency (Free tier allows ~15 RPM)
    return genai.GenerativeModel('gemini-1.5-flash')

# --- PROCESSING FUNCTION ---
def process_image(model, image_file):
    img = Image.open(image_file)
    
    # The Prompt: Strictly instructing the AI on how to read the AirSial Form
    prompt = """
    Analyze this AirSial Hazard Identification & Risk Assessment Form (AS-SMS-003 or AS-CS-003). 
    Extract the handwritten and printed data into a valid JSON object.
    
    Rules for Extraction:
    1. **Handwriting:** Transcribe exactly. If illegible, write "Unclear".
    2. **Checkboxes:** Look for tick marks (✓) or filled boxes in the Risk Matrix.
    3. **Dates:** Convert all dates to DD-MM-YYYY format.
    4. **Terminology:** Standardize locations (e.g., use "Ramp", "Galley").
    5. **Risk Matrix:** Identify the intersection of Severity and Probability that is checked.
    6. **Wet Lease:** If the form mentions "Wet Lease", "GetJet", or specific foreign registration, mark Wet Lease as "Yes".

    Return JSON with this structure:
    {
        "report_no": "String",
        "date_of_report": "DD-MM-YYYY",
        "reporter_name": "String",
        "department": "String",
        "location": "String",
        "hazard_description": "String",
        "severity_initial": "String (e.g., Moderate)",
        "probability_initial": "String (e.g., Occasional)",
        "risk_level_initial": "String (Low/Medium/High)",
        "cap_required": "Yes/No",
        "cap_action_plan": "String",
        "responsible_person": "String",
        "target_date": "DD-MM-YYYY",
        "cap_attachment": "Yes/No",
        "residual_severity": "String",
        "residual_probability": "String",
        "residual_risk_level": "String",
        "wet_lease_involved": "Yes/No",
        "operator_name": "String (e.g., AirSial / GetJet)",
        "report_attached": "Yes/No"
    }
    """
    
    try:
        response = model.generate_content([prompt, img])
        # Clean up code blocks if the model adds them
        json_text = response.text.replace("```json", "").replace("```", "")
        return json.loads(json_text)
    except Exception as e:
        return {"error": str(e), "report_no": f"Error-{image_file.name}"}

# --- EXCEL GENERATION (COMPLEX) ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formats
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1f4e78', 'font_color': 'white', 'border': 1})
        date_fmt = workbook.add_format({'num_format': 'dd-mm-yyyy', 'border': 1})
        cell_fmt = workbook.add_format({'border': 1, 'text_wrap': True})
        
        # --- SHEET 1: RAW DATA ---
        df.to_excel(writer, sheet_name='Raw SMS Data', index=False)
        worksheet = writer.sheets['Raw SMS Data']
        worksheet.set_column('A:T', 20) # Set general width
        # Apply header format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # --- SHEET 2: STANDARDIZED LISTS (For Validation) ---
        lists_sheet = workbook.add_worksheet('Standardized Lists')
        locations = ["Ramp", "Galley", "Cabin", "Cockpit", "Check-in", "Baggage Make-up", "Remote Bay"]
        risks = ["Low", "Medium", "High"]
        severities = ["Catastrophic", "Major", "Moderate", "Minor", "Insignificant"]
        
        lists_sheet.write_column('A1', locations)
        lists_sheet.write_column('B1', risks)
        lists_sheet.write_column('C1', severities)
        lists_sheet.hide() # Hide this sheet

        # --- SHEET 3: CAP TRACKER ---
        # We create a new dataframe for CAP tracking based on Raw Data
        cap_df = df[['report_no', 'cap_action_plan', 'responsible_person', 'target_date', 'cap_required']].copy()
        # Add formula columns for "Days Overdue" and "Status"
        # Note: Excel formulas in Python are strings. 
        # Assuming Data starts at row 2 (index 1).
        
        cap_df.to_excel(writer, sheet_name='CAP Tracker', index=False)
        cap_sheet = writer.sheets['CAP Tracker']
        cap_sheet.set_column('A:E', 25)
        for col_num, value in enumerate(cap_df.columns.values):
            cap_sheet.write(0, col_num, value, header_fmt)

        # --- SHEET 4: AUDIT LOG ---
        audit_df = df[['report_no', 'cap_attachment', 'report_attached']].copy()
        audit_df['Auditor Remarks'] = "" # Empty column for manual entry
        audit_df.to_excel(writer, sheet_name='Audit Evidence Log', index=False)
        audit_sheet = writer.sheets['Audit Evidence Log']
        for col_num, value in enumerate(audit_df.columns.values):
            audit_sheet.write(0, col_num, value, header_fmt)

    return output.getvalue()

# --- DASHBOARD GENERATION ---
def generate_dashboard(df):
    st.markdown("---")
    st.header("📊 Monthly SMS Dashboard (Preview)")
    
    # 1. KPI TILES
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Hazards", len(df))
    with col2:
        high_risk = len(df[df['risk_level_initial'] == 'High'])
        st.metric("High Risk Hazards", high_risk, delta_color="inverse")
    with col3:
        wet_lease = len(df[df['wet_lease_involved'] == 'Yes'])
        st.metric("Wet Lease Incidents", wet_lease)
    with col4:
        pending_cap = len(df[df['cap_required'] == 'Yes'])
        st.metric("CAPs Identified", pending_cap)

    # 2. CHARTS
    c1, c2 = st.columns(2)
    
    with c1:
        st.subheader("Hazards by Location")
        if not df.empty and 'location' in df.columns:
            fig_loc = px.bar(df, x='location', title="Location Frequency", color='location')
            st.plotly_chart(fig_loc, use_container_width=True)
            
    with c2:
        st.subheader("Risk Severity Distribution")
        if not df.empty and 'severity_initial' in df.columns:
            fig_sev = px.pie(df, names='severity_initial', title="Initial Severity Levels", hole=0.3)
            st.plotly_chart(fig_sev, use_container_width=True)

    # 3. AI INSIGHTS
    st.subheader("🤖 AI Safety Insights")
    if not df.empty:
        # Simple rule-based insights for demo (avoids extra API calls)
        insights = []
        if wet_lease > 0:
            insights.append(f"⚠️ **Wet Lease Alert:** {wet_lease} incidents reported on wet-leased aircraft. Check operator coordination procedures.")
        if high_risk > 0:
            insights.append(f"🚨 **Critical Attention:** {high_risk} High-Risk events detected. Immediate review required.")
        
        most_common_loc = df['location'].mode()[0] if not df['location'].empty else "N/A"
        insights.append(f"📍 **Hotspot:** The most frequent hazard location is **{most_common_loc}**.")

        for i in insights:
            st.warning(i)

# --- MAIN APP LOGIC ---
st.title("🛫 AirSial SMS Digitizer & Dashboard")
st.markdown("Upload scans of **AS-SMS-003** forms. The AI will digitize handwritten remarks, checkboxes, and risk matrices.")

uploaded_files = st.file_uploader("Upload Report Images (JPG, PNG)", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])

if st.button("🚀 Extract Data"):
    if not api_key:
        st.error("Please enter your API Key in the sidebar first.")
    elif not uploaded_files:
        st.warning("Please upload at least one image.")
    else:
        model = configure_gemini(api_key)
        results = []
        
        # Progress Bar
        progress_text = "Scanning reports with AI..."
        my_bar = st.progress(0, text=progress_text)
        
        for i, file in enumerate(uploaded_files):
            # Extract
            data = process_image(model, file)
            results.append(data)
            # Update bar
            percent_complete = int(((i + 1) / len(uploaded_files)) * 100)
            my_bar.progress(percent_complete, text=f"Processing {file.name} ({i+1}/{len(uploaded_files)})")
        
        my_bar.empty()
        st.success("✅ Extraction Complete!")
        
        # Process Dataframe
        df = pd.DataFrame(results)
        
        # Show Data Preview
        st.subheader("📝 Extracted Data Preview")
        st.dataframe(df)
        
        # Generate Dashboard
        generate_dashboard(df)
        
        # Export Excel
        excel_data = to_excel(df)
        st.download_button(
            label="📥 Download Audit-Ready Excel (.xlsx)",
            data=excel_data,
            file_name=f"AirSial_SMS_Log_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
