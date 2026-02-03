import os, re, tempfile, shutil
from pdf2image import convert_from_path
import easyocr
import pytesseract
import cv2
import numpy as np
import pandas as pd
from datetime import datetime

# Initialize EasyOCR reader once
READER = easyocr.Reader(['en'], gpu=False)

# Normalization maps
LOCATION_MAP = {
    'ramp':'Ramp', 'galley':'Galley', 'airport services':'Airport Services',
    'cabin':'Cabin', 'remote parking bay':'Remote Parking Bay'
}
SEVERITY_MAP = {
    'catastrophic':'Catastrophic','major':'Major','moderate':'Moderate','minor':'Minor','insignificant':'Insignificant'
}
PROB_MAP = {
    'frequent':'Frequent','occasional':'Occasional','remote':'Remote','improbable':'Improbable','rare':'Rare'
}
RISK_LEVEL_MAP = {'low':'Low','medium':'Medium','high':'High'}

def normalize_term(text, mapping):
    if not text: return text
    t = text.strip().lower()
    for k,v in mapping.items():
        if k in t:
            return v
    return text

def pdf_to_images(path):
    images = convert_from_path(path, dpi=300)
    tmp = []
    for i,img in enumerate(images):
        p = os.path.join(tempfile.gettempdir(), f'page_{os.path.basename(path)}_{i}.png')
        img.save(p, 'PNG')
        tmp.append(p)
    return tmp

def run_ocr_on_image(path):
    # combine EasyOCR and pytesseract for robustness
    img = cv2.imread(path, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # EasyOCR
    try:
        res = READER.readtext(gray, detail=0)
    except Exception:
        res = []
    # Tesseract fallback
    try:
        t = pytesseract.image_to_string(gray)
        t_lines = [l.strip() for l in t.splitlines() if l.strip()]
    except Exception:
        t_lines = []
    combined = res + t_lines
    return '\n'.join(combined), img

def detect_checkboxes(img):
    # returns dict of bounding boxes and whether filled
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _,th = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)
    # morphological to connect
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT,(3,3))
    closed = cv2.morphologyEx(th, cv2.MORPH_CLOSE, kernel, iterations=2)
    contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    boxes = []
    for c in contours:
        x,y,w,h = cv2.boundingRect(c)
        if 10 < w < 200 and 10 < h < 200 and 0.5 < w/h < 2.0:
            roi = gray[y:y+h, x:x+w]
            nonzero = cv2.countNonZero(cv2.threshold(roi, 200, 255, cv2.THRESH_BINARY_INV)[1])
            area = w*h
            fill_ratio = nonzero/area
            filled = fill_ratio > 0.12  # tuned threshold
            boxes.append({'x':x,'y':y,'w':w,'h':h,'filled':filled})
    return boxes

def parse_fields_from_text(text):
    # Heuristic parsing using regex and keywords. Mark unclear if not found.
    def find_after(key):
        m = re.search(rf'{key}[:\s\-]*([^\n\r]+)', text, re.IGNORECASE)
        return m.group(1).strip() if m else None

    report_no = find_after('report no|report number') or 'Unclear'
    date_report = find_after('date of report|date') or 'Unclear'
    # normalize date
    date_report = normalize_date(date_report)
    reporter = find_after('reporter name|reported by') or 'Unclear'
    dept = find_after('department') or 'Unclear'
    location = find_after('location of hazard|location') or 'Unclear'
    location = normalize_term(location, LOCATION_MAP)
    hazard_desc = find_after('hazard description|description') or 'Unclear'
    root_cause = find_after('root cause') or ''
    remarks = find_after('remarks') or ''
    # initial risk
    severity = find_after('severity') or 'Unclear'
    severity = normalize_term(severity, SEVERITY_MAP)
    probability = find_after('probability') or 'Unclear'
    probability = normalize_term(probability, PROB_MAP)
    initial_risk = find_after('initial risk level') or 'Unclear'
    initial_risk = normalize_term(initial_risk, RISK_LEVEL_MAP)
    cap_required = find_after('cap required|corrective action required') or 'Unclear'
    cap_required = 'Yes' if 'yes' in (cap_required or '').lower() else ('No' if 'no' in (cap_required or '').lower() else 'Unclear')
    cap_desc = find_after('action description|corrective action') or ''
    resp_person = find_after('responsible person') or ''
    resp_dept = find_after('responsible department') or ''
    target_date = normalize_date(find_after('target completion date|target date'))
    cap_attachment = find_after('attachment') or ''
    cap_attachment = 'Yes' if 'yes' in (cap_attachment or '').lower() else ('No' if 'no' in (cap_attachment or '').lower() else 'Unclear')
    # residual
    res_sev = normalize_term(find_after('residual severity') or '', SEVERITY_MAP) or 'Unclear'
    res_prob = normalize_term(find_after('residual probability') or '', PROB_MAP) or 'Unclear'
    res_risk = normalize_term(find_after('residual risk level') or '', RISK_LEVEL_MAP) or 'Unclear'
    risk_status = find_after('risk status') or 'Unclear'
    wet_lease = find_after('wet lease') or ''
    wet_lease = 'Yes' if 'yes' in (wet_lease or '').lower() else ('No' if 'no' in (wet_lease or '').lower() else 'Unclear')
    operator = find_after('operator name|operator') or ''
    report_attached = find_after('report attached') or ''
    report_attached = 'Yes' if 'yes' in (report_attached or '').lower() else ('No' if 'no' in (report_attached or '').lower() else 'Unclear')

    return {
        'Report Number': report_no,
        'Date of Report': date_report,
        'Reporter Name': reporter,
        'Department': dept,
        'Location of Hazard': location,
        'Hazard Description': hazard_desc,
        'Root Cause': root_cause or 'Unclear',
        'Remarks': remarks or 'Unclear',
        'Severity (Initial)': severity or 'Unclear',
        'Probability (Initial)': probability or 'Unclear',
        'Initial Risk Level': initial_risk or 'Unclear',
        'CAP Required': cap_required,
        'CAP Description': cap_desc or 'Unclear',
        'Responsible Person': resp_person or 'Unclear',
        'Responsible Dept': resp_dept or 'Unclear',
        'Target Date': target_date or 'Unclear',
        'CAP Attachment Mentioned': cap_attachment,
        'Residual Severity': res_sev,
        'Residual Probability': res_prob,
        'Residual Risk Level': res_risk,
        'Risk Status': risk_status or 'Unclear',
        'Wet Lease Aircraft Involved': wet_lease,
        'Operator Name': operator or 'Unclear',
        'Report Attached': report_attached
    }

def normalize_date(text):
    if not text:
        return 'Unclear'
    text = text.strip()
    # try common formats
    for fmt in ('%d-%m-%Y','%d/%m/%Y','%Y-%m-%d','%d %b %Y','%d %B %Y','%d.%m.%Y'):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime('%d-%m-%Y')
        except:
            pass
    # try to extract numbers
    m = re.search(r'(\d{1,2})[^\d](\d{1,2})[^\d](\d{2,4})', text)
    if m:
        d,mn,y = m.groups()
        y = y if len(y)==4 else ('20'+y if len(y)==2 else y)
        try:
            dt = datetime(int(y), int(mn), int(d))
            return dt.strftime('%d-%m-%Y')
        except:
            pass
    return 'Unclear'

def process_files_to_dataframe(folder):
    rows = []
    for fname in os.listdir(folder):
        path = os.path.join(folder, fname)
        ext = fname.rsplit('.',1)[-1].lower()
        images = []
        if ext == 'pdf':
            images = pdf_to_images(path)
        else:
            images = [path]
        full_text = ''
        checkbox_info = []
        for img_path in images:
            text, img = run_ocr_on_image(img_path)
            full_text += '\n' + text
            boxes = detect_checkboxes(img)
            checkbox_info.extend(boxes)
        parsed = parse_fields_from_text(full_text)
        # Use checkbox_info heuristics to override CAP Required or attachments if detected near keywords (advanced: spatial mapping)
        # Simple heuristic: if many filled boxes exist, mark CAP Required Yes
        if len([b for b in checkbox_info if b['filled']]) >= 1 and parsed.get('CAP Required','Unclear')=='Unclear':
            parsed['CAP Required'] = 'Yes'
        rows.append(parsed)
        # cleanup page images created by pdf2image
        for p in images:
            if p.startswith(tempfile.gettempdir()):
                try:
                    os.remove(p)
                except:
                    pass
    df = pd.DataFrame(rows)
    # Ensure all required columns exist
    cols = [
        'Report Number','Date of Report','Reporter Name','Department','Location of Hazard',
        'Hazard Description','Root Cause','Remarks','Severity (Initial)','Probability (Initial)',
        'Initial Risk Level','CAP Required','CAP Description','Responsible Person','Responsible Dept',
        'Target Date','CAP Attachment Mentioned','Residual Severity','Residual Probability',
        'Residual Risk Level','Risk Status','Wet Lease Aircraft Involved','Operator Name','Report Attached'
    ]
    for c in cols:
        if c not in df.columns:
            df[c] = 'Unclear'
    df = df[cols]
    return df
