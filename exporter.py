import pandas as pd
import xlsxwriter
from datetime import datetime

def create_audit_excel(df, out_path):
    # df is pandas DataFrame with one row per form
    writer = pd.ExcelWriter(out_path, engine='xlsxwriter', datetime_format='dd-mm-yyyy')
    workbook = writer.book

    # Sheet 1 Raw SMS Data (locked)
    df.to_excel(writer, sheet_name='Raw SMS Data', index=False)
    ws1 = writer.sheets['Raw SMS Data']
    # Protect sheet (no password)
    ws1.protect()

    # Sheet 2 Standardized Lists
    lists = {
        'Locations':['Ramp','Galley','Cabin','Remote Parking Bay','Others'],
        'Hazard Types':['Operational','Ground Handling','Cabin Safety','Maintenance-related','Human Factors'],
        'Departments':['Ramp Ops','Cabin Crew','Maintenance','Airport Services','Flight Ops'],
        'Risk Levels':['Low','Medium','High']
    }
    df_lists = pd.DataFrame(dict([(k, pd.Series(v)) for k,v in lists.items()]))
    df_lists.to_excel(writer, sheet_name='Standardized Lists', index=False)
    ws2 = writer.sheets['Standardized Lists']

    # Sheet 3 CAP Tracker
    cap_cols = ['Report No','Target Date','Status','Days Overdue']
    cap_df = pd.DataFrame(columns=cap_cols)
    # populate from df
    for _,row in df.iterrows():
        status = 'Open' if row.get('CAP Required','No')=='Yes' and row.get('Report Attached','No')!='Yes' else 'Closed'
        target = row.get('Target Date','Unclear')
        days_overdue = ''
        try:
            if target!='Unclear':
                dt = datetime.strptime(target, '%d-%m-%Y')
                days_overdue = (datetime.now() - dt).days
        except:
            days_overdue = ''
        cap_df = cap_df.append({
            'Report No': row.get('Report Number',''),
            'Target Date': target,
            'Status': status,
            'Days Overdue': days_overdue
        }, ignore_index=True)
    cap_df.to_excel(writer, sheet_name='CAP Tracker', index=False)
    ws3 = writer.sheets['CAP Tracker']
    # Add conditional formatting for traffic lights
    ws3.conditional_format('C2:C1000', {'type':'text','criteria':'containing','value':'Overdue','format':workbook.add_format({'bg_color':'#FF0000'})})
    # Sheet 4 Monthly Dashboard placeholder
    pd.DataFrame().to_excel(writer, sheet_name='Monthly Dashboard', index=False)
    # Sheet 5 Audit Evidence Log
    audit_df = pd.DataFrame(columns=['Report No','CAP Attachment Available','Evidence File Name','Auditor Remarks'])
    audit_df.to_excel(writer, sheet_name='Audit Evidence Log', index=False)

    writer.save()
