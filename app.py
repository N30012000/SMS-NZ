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

# --- SIDEBAR & SETUP ---
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/en/thumb/8/8e/AirSial_Logo.svg/1200px-AirSial_Logo.svg.png", width=200)
    st.title("⚙️ Configuration")
    
    api_key = st.text_input("Enter Google Gemini API Key", type="password", help="Get your free key from aistudio.google.com")
    
    # --- DEBUGGER: Check Available Models ---
    if api_key:
        try:
            genai.configure(api_key=api_key)
            st.success("API Key Accepted ✅")
            
            # List available models to help debug the 404 error
            with st.expander("🛠️ View Available Models"):
                try:
                    models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
                    st.write(models)
                except Exception as e:
                    st.error(f"Could not list models: {e}")
        except:
            st.error("Invalid API Key")

    st.info("ℹ️ **Privacy:** Data is processed in-memory and deleted after use.")

# --- GEMINI AI SETUP ---
def get_model(api_key):
    genai.configure(api_key=api_key)
    
    # Try the most stable specific version first, then fallbacks
    # 'gemini-1.5-flash-latest' often resolves alias 404s
    return genai.GenerativeModel('gemini-1.5-flash-latest')

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
    
    # Default fallback data
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
        return {**default_data, **data}
    except Exception as e:
        # If the model fails, we capture the error in the description so you can see it in Excel
        default_data["hazard_description"] = f"AI Error: {str(e)}"
        return default_data

# --- EXCEL GENERATOR ---
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
            # Check which columns exist before selecting
            cols = ['report_no', 'cap_action_plan', 'responsible_person', 'target_date', 'cap_required']
            existing = [c for c in cols if c in df.columns]
            cap_df = df[existing].copy()
            
            cap_df.to_excel(writer, sheet_name='CAP Tracker', index=False)
            ws_cap = writer.sheets['CAP Tracker']
            for col_num, value in enumerate(cap_df.columns.values):
                ws_cap.write(0, col_num, value, header_
