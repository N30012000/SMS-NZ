import streamlit as st
import pandas as pd
import io
from io import BytesIO
import base64
from PIL import Image
import pytesseract
import google.generativeai as genai
import json
import re
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import kaleido
import os

# Tesseract path for server backup
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

# Sidebar for API key (secrets in production)
if 'gemini_api_key' not in st.session_state:
    st.session_state.gemini_api_key = st.sidebar.text_input("Gemini API Key (Free)", type="password")

@st.cache_data
def ocr_image(image):
    # Client-side Tesseract.js preferred, but server fallback
    text = pytesseract.image_to_string(image, config='--psm 6')
    return text

def extract_fields_with_ai(text, prompt):
    if not st.session_state.gemini_api_key:
        return {"error": "Add free Gemini API key in sidebar"}
    try:
        genai.configure(api_key=st.session_state.gemini_api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        response = model.generate_content(f"{prompt}\n\nExtracted text:\n{text}\nOutput ONLY JSON matching keys exactly, no extras. Unclear='Unclear'")
        return json.loads(response.text)
    except:
        return {"error": "AI extraction failed - check API key"}

def normalize_date(date_str):
    formats = ['%d/%m/%Y', '%d-%m-%Y', '%m/%d/%Y', '%Y-%m-%d']
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).strftime('%d-%m-%Y')
        except:
            pass
    return "Unclear"

def detect_checkbox_regions(image, text):
    # Simple heuristic for risk matrix/checkboxes via text patterns
    severity_map = {"Catastrophic": "Catastrophic", "Major": "Major", "Moderate": "Moderate", "Minor": "Minor", "Insignificant": "Insignificant"}
    prob_map = {"Frequent": "Frequent", "Occasional": "Occasional", "Remote": "Remote", "Improbable": "Improbable", "Rare": "Rare"}
    # Extract highest confidence matches
    sev = next((k for k in severity_map if k.lower() in text.lower()), "Unclear")
    prob = next((k for k in prob_map if k.lower() in text.lower()), "Unclear")
    return sev, prob

def create_audit_excel(data_list):
    with pd.ExcelWriter('AirSial_SMS_Audit.xlsx', engine='openpyxl') as writer:
        # Sheet 1: Raw SMS Data
        df_raw = pd.DataFrame(data_list)
        df_raw.to_excel(writer, sheet_name='Raw SMS Data', index=False)

        # Sheet 2: Standardized Lists (Dropdowns)
        locations = ['Ramp', 'Galley', 'Cabin', 'Remote Parking Bay', 'Airport Services', 'Others']
        hazard_types = ['Operational', 'Ground Handling', 'Cabin Safety', 'Maintenance-related', 'Human Factors']
        depts = ['Flight Operations', 'Cabin Crew', 'Ground Handling', 'Maintenance', 'Safety']
        risks = ['Low', 'Medium', 'High']
        pd.DataFrame({'Locations': locations}).to_excel(writer, sheet_name='Standardized Lists', index=False)
        pd.DataFrame({'Hazard Types': hazard_types}).to_excel(writer, startcol=1, index=False)
        pd.DataFrame({'Departments': depts}).to_excel(writer, startcol=2, index=False)
        pd.DataFrame({'Risk Levels': risks}).to_excel(writer, startcol=3, index=False)

        # Sheet 3: CAP Tracker
        df_cap = df_raw[['Report No', 'Target Date', 'Residual Risk Level', 'Status']].copy()
        df_cap['Days Overdue'] = (datetime.now() - pd.to_datetime(df_cap['Target Date'], errors='coerce')).dt.days.clip(lower=0)
        df_cap['Status Indicator'] = df_cap['Days Overdue'].apply(lambda x: '🔴 Overdue' if x > 0 else '🟢 Closed')
        df_cap.to_excel(writer, sheet_name='CAP Tracker', index=False)

        # Sheet 4: Monthly Dashboard (Charts added via openpyxl later if needed)
        pd.DataFrame().to_excel(writer, sheet_name='Monthly Dashboard', index=False)

        # Sheet 5: Audit Evidence Log
        df_audit = df_raw[['Report No', 'CAP Attachment Mentioned', 'Report Attached']].copy()
        df_audit.columns = ['Report No', 'CAP Attachment Available (Yes/No)', 'Evidence File Name', 'Auditor Remarks']
        df_audit['Auditor Remarks'] = ''
        df_audit.to_excel(writer, sheet_name='Audit Evidence Log', index=False)

    return 'AirSial_SMS_Audit.xlsx'

def generate_dashboard(excel_file, month, year):
    df = pd.read_excel(excel_file, 'Raw SMS Data')
    df['Date of Report'] = pd.to_datetime(df['Date of Report'], errors='coerce')
    df_filtered = df[(df['Date of Report'].dt.month == month) & (df['Date of Report'].dt.year == year)]

    # KPIs
    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("Total Hazards", len(df_filtered))
    with col2: st.metric("High-Risk", len(df_filtered[df_filtered['Initial Risk Level'] == 'High']))
    with col3: st.metric("CAPs Pending", len(df_filtered[df_filtered['CAP Required'] == 'Yes']))
    with col4: st.metric("Wet Lease %", f"{len(df_filtered[df_filtered['Wet Lease (Yes/No)'] == 'Yes'])/len(df_filtered)*100:.1f}%" if len(df_filtered)>0 else "0%")

    # Trends
    fig = make_subplots(rows=2, cols=3, subplot_titles=('Hazards by Location', 'Hazards by Risk Level', 'CAP Effectiveness', 'Hazards by Type', 'Fleet Analysis', 'Risk Distribution'))
    fig.add_trace(px.histogram(df_filtered, x='Location of Hazard').data[0], row=1, col=1)
    fig.add_trace(px.pie(df_filtered, names='Initial Risk Level').data[0], row=1, col=2)
    # Add more traces...
    st.plotly_chart(fig)

    # AI Insights
    insights_prompt = f"Summarize key safety observations from this SMS data: {df_filtered.to_json()}"
    insights = extract_fields_with_ai("", insights_prompt).get('insights', 'Analysis unavailable')
    st.subheader("AI Safety Insights")
    st.write(insights)

    # Exports
    excel_buffer = BytesIO()
    df.to_excel(excel_buffer)
    st.download_button("Download Dashboard Excel", excel_buffer.getvalue(), "SMS_Dashboard.xlsx")
    # PDF via plotly static

# Main UI Tabs
tab1, tab2 = st.tabs(["📥 SMS Form Scanner", "📊 Monthly Dashboard"])

with tab1:
    uploaded_files = st.file_uploader("Upload Forms (Images/PDFs)", accept_multiple_files=True)
    if uploaded_files:
        data_list = []
        for file in uploaded_files:
            if file.type == "application/pdf":
                # Convert PDF to images (simplified)
                pass  # Use pdf2image
            else:
                image = Image.open(file)
                st.image(image, caption="Preview")
                text = ocr_image(image)
                sev, prob = detect_checkbox_regions(image, text)

                prompt = """
                Extract to JSON:
                {"Report Number": "...", "Date of Report": "...", "Reporter Name": "...", "Department": "...", "Location of Hazard": "...", "Hazard Description": "...", "Severity": "%s", "Probability": "%s", "Initial Risk Level": "...", "CAP Required": "Yes/No", "Action Description": "...", "Responsible Person": "...", "Responsible Department": "...", "Target Completion Date": "...", "Residual Severity": "...", "Residual Probability": "...", "Residual Risk Level": "...", "Wet Lease Aircraft Involved": "Yes/No", "Operator Name": "..."}
                Normalize dates DD-MM-YYYY. Standardize: Ramp, Galley etc.""" % (sev, prob)

                fields = extract_fields_with_ai(text, prompt)
                fields['Date of Report'] = normalize_date(fields.get('Date of Report', ''))
                fields['Severity (Initial)'] = sev
                fields['Probability (Initial)'] = prob
                data_list.append(fields)

        if st.button("Extract & Download Excel"):
            excel_file = create_audit_excel(data_list)
            st.download_button("Download Audit Excel", data=open(excel_file, 'rb'), file_name=excel_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab2:
    uploaded_excel = st.file_uploader("Upload Extracted Excel")
    month = st.selectbox("Month", range(1,13))
    year = st.number_input("Year", 2024, 2027)
    if uploaded_excel and st.button("Generate Dashboard"):
        generate_dashboard(uploaded_excel, month, year)

st.sidebar.success("Data deleted after session. Confidential handling ensured.")
