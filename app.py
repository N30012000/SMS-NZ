# app.py - Simplified version for GitHub deployment
import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import plotly.express as px
import plotly.graph_objects as go
import tempfile
import os
import base64
import io
import warnings
warnings.filterwarnings('ignore')

# For GitHub deployment, we'll use pytesseract instead of PaddleOCR
# as it's easier to install
try:
    import pytesseract
    OCR_AVAILABLE = True
except:
    OCR_AVAILABLE = False

# Initialize session state
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = []
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = pd.DataFrame()

# Standardized values
STANDARDIZED_VALUES = {
    'locations': ['Ramp', 'Galley', 'Cabin', 'Flight Deck', 'Remote Parking Bay', 'Terminal'],
    'hazard_types': ['Operational', 'Ground Handling', 'Cabin Safety', 'Maintenance', 'Human Factors'],
    'departments': ['Airport Services', 'Flight Operations', 'Maintenance', 'Cabin Crew', 'Ground Handling'],
    'severity_levels': ['Catastrophic', 'Major', 'Moderate', 'Minor', 'Insignificant'],
    'probability_levels': ['Frequent', 'Occasional', 'Remote', 'Improbable', 'Rare'],
    'risk_levels': ['High', 'Medium', 'Low']
}

def extract_text_from_image(image):
    """Extract text using pytesseract if available, otherwise use fallback"""
    if OCR_AVAILABLE:
        try:
            # Convert to RGB if needed
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # Extract text
            text = pytesseract.image_to_string(image)
            return text
        except Exception as e:
            st.warning(f"OCR Error: {str(e)}")
            return ""
    else:
        # Fallback for demo - return sample text
        st.info("OCR not available in this environment. Using sample data for demonstration.")
        return """AIRISIAL Hazard Identification & Risk Assessment Form
        Tracking #: HR-Con-123/2025
        Date of Report: 18-12-2025
        Reporter Name & Contact no.: Confidential
        Department: Airport Services
        Location of Hazard: Ramp Area A
        Date/Time Hazard Identified: 15-12-2025
        Initial Risk Assessment
        Severity: Major
        Probability: Occasional
        Initial Risk Rating: Medium
        Corrective Action Plan
        Action Plan: Crew reported incident. CAP attached.
        Target Date: 31-01-2026
        Residual Risk: Low
        Remarks: Bus availability issue at remote bays."""

def parse_date(date_str):
    """Parse and normalize dates"""
    if not date_str or date_str.lower() == 'unclear':
        return ""
    
    patterns = [
        r'(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})',
        r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, date_str)
        if match:
            try:
                if len(match.group(3)) == 4:  # DD-MM-YYYY
                    return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
                elif len(match.group(1)) == 4:  # YYYY-MM-DD
                    return f"{match.group(3)}-{match.group(2)}-{match.group(1)}"
            except:
                continue
    
    return date_str

def extract_form_data(text):
    """Extract form data from text"""
    data = {
        'report_number': extract_field(text, r'Tracking\s*#\s*[:]?\s*([A-Za-z0-9\-/]+)') or 
                       extract_field(text, r'Report\s*no\.?\s*[:]?\s*([A-Za-z0-9\-/]+)') or 
                       'N/A',
        'date_of_report': parse_date(extract_field(text, r'Date\s*of\s*Report[:\s]*(\d{1,2}[-/.]\d{1,2}[-/.]\d{4})')),
        'reporter_name': extract_field(text, r'Reporter\s*Name[:\s]*([A-Za-z\s]+)') or 'Confidential',
        'department': extract_field(text, r'Department[:\s]*([A-Za-z\s&]+)') or 'Airport Services',
        'location': extract_field(text, r'Location\s*of\s*Hazard[:\s]*([A-Za-z\s]+)') or 'Ramp',
        'hazard_date': parse_date(extract_field(text, r'Date/Time\s*Hazard\s*Identified[:\s]*(\d{1,2}[-/.]\d{1,2}[-/.]\d{4})')),
        'severity': extract_risk_field(text, ['Catastrophic', 'Major', 'Moderate', 'Minor', 'Insignificant']),
        'probability': extract_risk_field(text, ['Frequent', 'Occasional', 'Remote', 'Improbable', 'Rare']),
        'initial_risk': extract_risk_field(text, ['High', 'Medium', 'Low'], context='Initial'),
        'cap_required': 'Yes' if 'Action Plan' in text or 'Corrective Action' in text else 'No',
        'action_description': extract_field(text, r'Action\s*Plan[:\s]+(.+?)(?:Target Date|$)') or 
                            extract_field(text, r'Corrective\s*Action[:\s]+(.+?)(?:Target Date|$)') or '',
        'target_date': parse_date(extract_field(text, r'Target\s*Date[:\s]*(\d{1,2}[-/.]\d{1,2}[-/.]\d{4})')),
        'residual_risk': extract_risk_field(text, ['High', 'Medium', 'Low'], context='Residual'),
        'remarks': extract_field(text, r'Remarks[:\s]+(.+?)(?:Sign-off|$)') or '',
        'wet_lease': 'Yes' if 'wet lease' in text.lower() else 'No',
        'status': 'Open'
    }
    
    return data

def extract_field(text, pattern):
    """Extract field using regex pattern"""
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    return match.group(1).strip() if match else None

def extract_risk_field(text, options, context=None):
    """Extract risk field from options"""
    text_lower = text.lower()
    if context:
        # Find context section
        if context.lower() == 'initial':
            sections = text.split('Initial Risk')
            if len(sections) > 1:
                text_lower = sections[1].lower()[:500]
        elif context.lower() == 'residual':
            sections = text.split('Residual Risk')
            if len(sections) > 1:
                text_lower = sections[1].lower()[:500]
    
    for option in options:
        if option.lower() in text_lower:
            return option
    
    return ''

def create_excel_file(data_list):
    """Create Excel file with extracted data"""
    wb = openpyxl.Workbook()
    
    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        default_sheet = wb['Sheet']
        wb.remove(default_sheet)
    
    # Sheet 1: Raw Data
    ws1 = wb.create_sheet(title="Raw SMS Data")
    
    headers = [
        'Report No', 'Date of Report', 'Reporter Name', 'Department',
        'Location', 'Hazard Date', 'Severity', 'Probability', 
        'Initial Risk', 'CAP Required', 'Action Description', 'Target Date',
        'Residual Risk', 'Remarks', 'Wet Lease', 'Status'
    ]
    
    # Write headers
    for col, header in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    
    # Write data
    for row_idx, data in enumerate(data_list, 2):
        row_data = [
            data.get('report_number', ''),
            data.get('date_of_report', ''),
            data.get('reporter_name', ''),
            data.get('department', ''),
            data.get('location', ''),
            data.get('hazard_date', ''),
            data.get('severity', ''),
            data.get('probability', ''),
            data.get('initial_risk', ''),
            data.get('cap_required', ''),
            data.get('action_description', ''),
            data.get('target_date', ''),
            data.get('residual_risk', ''),
            data.get('remarks', ''),
            data.get('wet_lease', ''),
            data.get('status', '')
        ]
        
        for col, value in enumerate(row_data, 1):
            ws1.cell(row=row_idx, column=col, value=value)
    
    # Adjust column widths
    for column in ws1.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws1.column_dimensions[column_letter].width = adjusted_width
    
    return wb

def create_dashboard(data_list, month=None, year=None):
    """Create dashboard from data"""
    if not data_list:
        return None
    
    df = pd.DataFrame(data_list)
    
    # Filter by month/year if specified
    if month and year:
        df['date_obj'] = pd.to_datetime(df['date_of_report'], errors='coerce', dayfirst=True)
        df = df[(df['date_obj'].dt.month == month) & (df['date_obj'].dt.year == year)]
    
    if df.empty:
        return None
    
    # Create dashboard
    st.subheader("📊 Safety Dashboard")
    
    # KPIs
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Hazards", len(df))
    
    with col2:
        high_risk = len(df[df['initial_risk'] == 'High'])
        st.metric("High Risk", high_risk)
    
    with col3:
        caps = len(df[df['cap_required'] == 'Yes'])
        st.metric("CAPs Required", caps)
    
    with col4:
        wet_lease = len(df[df['wet_lease'] == 'Yes'])
        wet_lease_pct = (wet_lease / len(df) * 100) if len(df) > 0 else 0
        st.metric("Wet Lease", f"{wet_lease_pct:.1f}%")
    
    # Charts
    col1, col2 = st.columns(2)
    
    with col1:
        if 'location' in df.columns:
            location_counts = df['location'].value_counts()
            if not location_counts.empty:
                fig1 = px.pie(values=location_counts.values, 
                            names=location_counts.index,
                            title="Hazards by Location")
                st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        if 'initial_risk' in df.columns:
            risk_counts = df['initial_risk'].value_counts()
            if not risk_counts.empty:
                fig2 = px.bar(x=risk_counts.index, y=risk_counts.values,
                            title="Risk Level Distribution")
                st.plotly_chart(fig2, use_container_width=True)
    
    return df

def main():
    st.set_page_config(
        page_title="AirSial SMS Scanner",
        page_icon="✈️",
        layout="wide"
    )
    
    st.title("✈️ AirSial SMS Form Scanner")
    st.markdown("Automated Hazard Report Processing")
    
    # Sidebar
    with st.sidebar:
        st.header("Navigation")
        app_mode = st.radio(
            "Select Mode:",
            ["📤 Upload Forms", "📊 View Dashboard", "📁 Export Data"]
        )
        
        st.divider()
        
        if OCR_AVAILABLE:
            st.success("✅ OCR Available")
        else:
            st.warning("⚠️ OCR not available - using demo mode")
        
        st.info("Upload images/PDFs of hazard reports for automatic data extraction.")
    
    if app_mode == "📤 Upload Forms":
        st.header("Upload Hazard Reports")
        
        uploaded_files = st.file_uploader(
            "Choose files",
            type=['png', 'jpg', 'jpeg', 'pdf'],
            accept_multiple_files=True,
            help="Upload scanned forms or images"
        )
        
        if uploaded_files:
            st.success(f"📁 {len(uploaded_files)} files uploaded")
            
            if st.button("🚀 Extract Data", type="primary"):
                progress_bar = st.progress(0)
                
                for i, file in enumerate(uploaded_files):
                    try:
                        # Read file
                        if file.type == "application/pdf":
                            st.warning("PDF support requires additional setup. Using demo data.")
                            text = extract_text_from_image(Image.new('RGB', (100, 100), color='white'))
                        else:
                            image = Image.open(file)
                            text = extract_text_from_image(image)
                        
                        # Extract data
                        form_data = extract_form_data(text)
                        form_data['filename'] = file.name
                        
                        # Store in session
                        st.session_state.extracted_data.append(form_data)
                        
                    except Exception as e:
                        st.error(f"Error processing {file.name}: {str(e)}")
                    
                    progress_bar.progress((i + 1) / len(uploaded_files))
                
                st.success(f"✅ Extracted data from {len(st.session_state.extracted_data)} files")
        
        # Show extracted data
        if st.session_state.extracted_data:
            st.subheader("Extracted Data")
            
            # Convert to DataFrame for display
            df_display = pd.DataFrame(st.session_state.extracted_data)
            df_display = df_display.drop('filename', axis=1, errors='ignore')
            st.dataframe(df_display, use_container_width=True)
            
            # Download button
            if st.button("📥 Download Excel", type="primary"):
                wb = create_excel_file(st.session_state.extracted_data)
                
                # Save to bytes
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)
                
                # Download
                st.download_button(
                    label="⬇️ Download Excel File",
                    data=excel_buffer,
                    file_name="AirSial_SMS_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    elif app_mode == "📊 View Dashboard":
        st.header("Safety Dashboard")
        
        if not st.session_state.extracted_data:
            st.warning("Please upload and extract data first")
        else:
            # Month/Year selector
            col1, col2 = st.columns(2)
            with col1:
                month = st.selectbox("Month", range(1, 13), format_func=lambda x: f"{x:02d}")
            with col2:
                year = st.selectbox("Year", range(2020, 2026), index=5)
            
            if st.button("Generate Dashboard"):
                df = create_dashboard(st.session_state.extracted_data, month, year)
                
                if df is not None:
                    # Show insights
                    st.subheader("📋 Safety Insights")
                    
                    insights = []
                    if len(df) > 0:
                        insights.append(f"• Total reports: {len(df)}")
                        
                        high_risk = len(df[df['initial_risk'] == 'High'])
                        if high_risk > 0:
                            insights.append(f"• High-risk incidents: {high_risk}")
                        
                        caps = len(df[df['cap_required'] == 'Yes'])
                        if caps > 0:
                            insights.append(f"• CAPs required: {caps}")
                    
                    if insights:
                        st.info("\n".join(insights))
    
    elif app_mode == "📁 Export Data":
        st.header("Export Data")
        
        if not st.session_state.extracted_data:
            st.info("No data to export. Please upload files first.")
        else:
            # Export options
            export_format = st.selectbox("Export Format", ["Excel", "CSV", "JSON"])
            
            if st.button(f"Export as {export_format}"):
                df = pd.DataFrame(st.session_state.extracted_data)
                
                if export_format == "Excel":
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='SMS_Data')
                    buffer.seek(0)
                    
                    st.download_button(
                        label="Download Excel",
                        data=buffer,
                        file_name="sms_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                elif export_format == "CSV":
                    csv = df.to_csv(index=False)
                    st.download_button(
                        label="Download CSV",
                        data=csv,
                        file_name="sms_data.csv",
                        mime="text/csv"
                    )
                
                elif export_format == "JSON":
                    json_str = df.to_json(orient='records', indent=2)
                    st.download_button(
                        label="Download JSON",
                        data=json_str,
                        file_name="sms_data.json",
                        mime="application/json"
                    )

if __name__ == "__main__":
    main()
