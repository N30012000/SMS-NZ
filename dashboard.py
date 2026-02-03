import pandas as pd
import matplotlib.pyplot as plt
import os
from io import BytesIO
from matplotlib.backends.backend_pdf import PdfPages
import tempfile
from datetime import datetime

def generate_dashboard_from_excel(excel_path, month, year):
    df = pd.read_excel(excel_path, sheet_name='Raw SMS Data')
    # filter by month/year on Date of Report
    def parse_date(s):
        try:
            return datetime.strptime(s, '%d-%m-%Y')
        except:
            return None
    df['parsed_date'] = df['Date of Report'].apply(parse_date)
    df_month = df[df['parsed_date'].notnull() & (df['parsed_date'].dt.month==month) & (df['parsed_date'].dt.year==year)]

    # KPIs
    total = len(df_month)
    high_risk = len(df_month[df_month['Initial Risk Level'].str.lower()=='high'])
    caps_pending = len(df_month[(df_month['CAP Required']=='Yes') & (df_month['Report Attached']!='Yes')])
    wet_lease_pct = 0
    if total>0:
        wet_lease_pct = round(100 * len(df_month[df_month['Wet Lease Aircraft Involved']=='Yes']) / total, 1)

    # Charts
    charts = {}
    # Hazards by Location
    loc_counts = df_month['Location of Hazard'].fillna('Others').value_counts()
    fig1, ax1 = plt.subplots(figsize=(6,4))
    loc_counts.plot(kind='bar', ax=ax1, color='tab:blue')
    ax1.set_title('Hazards by Location')
    charts['hazards_by_location'] = fig1

    # Hazards by Type (requires mapping; use Hazard Type column if exists)
    if 'Hazard Type' in df_month.columns:
        type_counts = df_month['Hazard Type'].fillna('Others').value_counts()
    else:
        type_counts = pd.Series(dtype=int)
    fig2, ax2 = plt.subplots(figsize=(6,4))
    if not type_counts.empty:
        type_counts.plot(kind='pie', ax=ax2, autopct='%1.1f%%')
        ax2.set_ylabel('')
        ax2.set_title('Hazards by Type')
    charts['hazards_by_type'] = fig2

    # Risk Level Distribution
    risk_counts = df_month['Initial Risk Level'].fillna('Unclear').value_counts()
    fig3, ax3 = plt.subplots(figsize=(6,4))
    risk_counts.plot(kind='bar', ax=ax3, color='tab:orange')
    ax3.set_title('Risk Level Distribution')
    charts['risk_level'] = fig3

    # CAP Effectiveness
    # For demo: closed on time vs overdue vs pending using CAP Tracker
    cap_df = df_month[['Report Number','Target Date','CAP Required','Report Attached']].copy()
    def cap_status(row):
        if row['CAP Required']!='Yes':
            return 'N/A'
        if row['Report Attached']=='Yes':
            return 'Closed on time'  # heuristic
        return 'Pending'
    cap_df['status'] = cap_df.apply(cap_status, axis=1)
    cap_counts = cap_df['status'].value_counts()
    fig4, ax4 = plt.subplots(figsize=(6,4))
    cap_counts.plot(kind='bar', ax=ax4, color='tab:green')
    ax4.set_title('CAP Effectiveness')
    charts['cap_effectiveness'] = fig4

    # Save charts to temporary HTML preview and PDF and Excel
    tmpdir = tempfile.mkdtemp(prefix='dashboard_')
    pdf_path = os.path.join(tmpdir, f'dashboard_{month}_{year}.pdf')
    excel_path = os.path.join(tmpdir, f'dashboard_{month}_{year}.xlsx')
    # Save PDF
    with PdfPages(pdf_path) as pdf:
        for k, fig in charts.items():
            pdf.savefig(fig)
            plt.close(fig)
    # Save Excel with charts as images
    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    df_month.to_excel(writer, sheet_name='Raw SMS Data', index=False)
    workbook = writer.book
    ws = workbook.add_worksheet('Charts')
    writer.sheets['Raw SMS Data'] = writer.book.add_worksheet('Raw SMS Data')  # ensure exists
    # Insert chart images
    row = 0
    for k, fig in charts.items():
        img_path = os.path.join(tmpdir, f'{k}.png')
        fig.savefig(img_path)
        ws.insert_image(row, 0, img_path, {'x_scale':0.8,'y_scale':0.8})
        row += 20
    writer.close()

    # Create a simple HTML preview (list KPIs and embed images)
    html_path = os.path.join(tmpdir, 'preview.html')
    with open(html_path, 'w') as f:
        f.write(f"<h1>SMS Dashboard {month}/{year}</h1>")
        f.write(f"<p><b>Total Hazards Reported:</b> {total}</p>")
        f.write(f"<p><b>High-Risk Hazards:</b> {high_risk}</p>")
        f.write(f"<p><b>CAPs Pending:</b> {caps_pending}</p>")
        f.write(f"<p><b>Wet Lease Incidents (%):</b> {wet_lease_pct}</p>")
        for k in charts.keys():
            img = os.path.join(tmpdir, f'{k}.png')
            f.write(f"<h3>{k.replace('_',' ').title()}</h3>")
            f.write(f"<img src='{os.path.basename(img)}' style='max-width:800px'><br>")
    # copy images into same dir for HTML
    for k in charts.keys():
        img_path = os.path.join(tmpdir, f'{k}.png')
        # already saved above
    return {'html_preview': html_path, 'excel': excel_path, 'pdf': pdf_path}
