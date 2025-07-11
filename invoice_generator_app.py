import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from decimal import Decimal

def generate_invoice(timesheet_file, projects_file, monthly_salary):
    df_time_raw = pd.read_csv(timesheet_file)
    df_projects = pd.read_csv(projects_file)

    person_name = df_time_raw['Person'].iloc[0]
    dept_raw = df_time_raw['Department'].iloc[0]
    dept_num = re.search(r'\d+', str(dept_raw)).group()
    project_code_map = dict(zip(df_projects['Project'], df_projects['Project code']))

    df_time_raw['Total hrs per line'] = (
        df_time_raw['Logged hrs'].fillna(0) +
        df_time_raw['Time off hrs'].fillna(0) +
        df_time_raw['Holiday hrs'].fillna(0)
    )
    df_time = df_time_raw[df_time_raw['Total hrs per line'] > 0].copy()

    overtime_hrs = df_time_raw[
        df_time_raw['Time off'].fillna("").str.contains("ausgleich für zusätzliche arbeitszeit", case=False)
    ]['Time off hrs'].sum()

    all_logged_hrs = df_time_raw['Logged hrs'].sum()
    all_paid_timeoff_hrs = (
        df_time_raw['Time off hrs'].sum() +
        df_time_raw['Holiday hrs'].sum() -
        overtime_hrs
    )

    filename = timesheet_file.name
    match = re.search(r'Table-(\d{8})', filename)
    year, month = match.group(1)[:4], match.group(1)[4:6]
    time_period = datetime.strptime(f"{year}-{month}-01", "%Y-%m-%d").strftime("%B %Y")

    admin_keywords = ['Admin', 'BF coordination', 'IT', 'Finances']
    pc_keywords = ['HR', 'P&C']
    sales_projects = ['Sales I proposal support', 'Tender Screening']

    def resolve_project_code_and_company(row):
        project = row['Project']
        code = project_code_map.get(project)
        if pd.notna(code) and str(code).strip():
            return pd.Series({'Project Code': code, 'Company': 'PCG' if str(code).startswith('1') else 'PCR'})
        tags = df_projects[df_projects['Project'] == project]['Tags'].dropna().astype(str).str.upper()
        all_tags = ','.join(tags.tolist())
        if 'PCG' in all_tags:
            return pd.Series({'Project Code': 'no project code', 'Company': 'PCG'})
        elif 'PCR' in all_tags:
            return pd.Series({'Project Code': 'no project code', 'Company': 'PCR'})
        else:
            return pd.Series({'Project Code': 'no project code', 'Company': 'no company in project tags'})

    # Billable projects input for quota and final export
    df_billable_input = df_time[df_time['Project'].fillna('').str.match(r'^\d{2}_')].copy()
    df_billable_input[['Project Code', 'Company']] = df_billable_input.apply(resolve_project_code_and_company, axis=1)

    # Quota based on full data
    quota = df_billable_input.groupby('Company')['Logged hrs'].sum()
    total_real = quota.sum()
    pcg_ratio = float(quota.get('PCG', 0)) / total_real if total_real > 0 else 0.5
    pcr_ratio = float(quota.get('PCR', 0)) / total_real if total_real > 0 else 0.5

    admin_time = df_time[df_time['Project'].str.startswith(tuple(admin_keywords), na=False)]['Logged hrs'].sum()
    pc_time = df_time[df_time['Project'].str.startswith(tuple(pc_keywords), na=False)]['Logged hrs'].sum()
    sales_time = df_time[df_time['Project'].isin(sales_projects)]['Logged hrs'].sum()
    bf_general_hrs = admin_time + all_paid_timeoff_hrs

    def split_row(label, base_hrs, pcg_ratio, pcr_ratio, pcg_code, pcr_code):
        if base_hrs == 0:
            return pd.DataFrame()
        b = Decimal(str(base_hrs)).quantize(Decimal('0.0001'))
        pcg = Decimal(str(pcg_ratio)).quantize(Decimal('0.0001'))
        pcr = Decimal(str(pcr_ratio)).quantize(Decimal('0.0001'))
        return pd.DataFrame({
            'Project': [f'{label} (PCG)', f'{label} (PCR)'],
            'Total hrs': [None, None],
            'Project Code': [pcg_code, pcr_code],
            'Company': ['PCG', 'PCR'],
            'Formula': [f'={b}*{pcg}', f'={b}*{pcr}']
        })

    df_bf_split = pd.concat([
        split_row(f'BF{dept_num} General', bf_general_hrs, pcg_ratio, pcr_ratio, f'1{dept_num}000', f'2{dept_num}000'),
        split_row('People & Culture', pc_time, pcg_ratio, pcr_ratio, '199500', '299500'),
        split_row('Sales', sales_time, pcg_ratio, pcr_ratio, '199300', '299300')
    ], ignore_index=True)

    df_summary = df_time.groupby('Project', as_index=False)['Logged hrs'].sum()
    df_bf_dist = df_summary[df_summary['Project'].str.match(r'^(Sales_BF|Marketing_BF)\d{2}', na=False)].copy()

    dist_rows = []
    for _, row in df_bf_dist.iterrows():
        bf_match = re.search(r'(?:Sales_BF|Marketing_BF)(\d{2})', row['Project'])
        if bf_match:
            bf = bf_match.group(1)
            h = Decimal(str(row['Logged hrs'])).quantize(Decimal('0.0001'))
            pcg = Decimal(str(pcg_ratio)).quantize(Decimal('0.0001'))
            pcr = Decimal(str(pcr_ratio)).quantize(Decimal('0.0001'))
            dist_rows.append({'Project': f'BF{bf} General (PCG)', 'Project Code': f'1{bf}000', 'Company': 'PCG', 'Formula': f'={h}*{pcg}'})
            dist_rows.append({'Project': f'BF{bf} General (PCR)', 'Project Code': f'2{bf}000', 'Company': 'PCR', 'Formula': f'={h}*{pcr}'})
    df_sales_split = pd.DataFrame(dist_rows)

    # Final billable DataFrame for export
    df_billable = df_billable_input.groupby(['Project', 'Project Code', 'Company'], as_index=False)['Logged hrs'].sum()
    df_billable = df_billable.rename(columns={'Logged hrs': 'Total hrs'})
    df_billable["Formula"] = None

    frames = [df_billable, df_sales_split, df_bf_split]
    frames = [df for df in frames if not df.empty and df.dropna(how="all").shape[0] > 0]
    df_final = pd.concat(frames, ignore_index=True)

    # Deduplicate BFXX General by grouping across keys (not Formula)
    df_final = df_final.groupby(['Project', 'Project Code', 'Company'], as_index=False).agg({
        'Total hrs': 'sum',
        'Formula': 'first'
    })
    df_final['Days'] = df_final['Total hrs'] / 8

    # --- Excel Output ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    ws['A1'] = "Name:"; ws['B1'] = person_name
    ws['A2'] = "Business Field:"; ws['B2'] = dept_num
    ws['A3'] = "Time Period:"; ws['B3'] = time_period
    ws['A4'] = "Monthly Salary:"; ws['B4'] = monthly_salary; ws['B4'].number_format = u'€#,##0.00'
    ws['A5'] = "Number of Days Worked:"; ws['B5'] = f"={all_logged_hrs}/8"
    ws['A6'] = "Paid Time-off Days:"; ws['B6'] = f"=({round(all_paid_timeoff_hrs, 1)})/8"
    ws['A7'] = "Total days:"; ws['B7'] = "=B5+B6"
    ws['A8'] = "Day Rate:"; ws['B8'] = "=B4/B7"; ws['B8'].number_format = u'€#,##0.00'

    def write_table(df, start_row, title):
        ws[f"A{start_row}"] = title
        ws[f"A{start_row}"].font = Font(bold=True)
        ws.append(['Project Code', 'Project', 'Logged hrs', 'Days', 'Day Rate', 'Costs'])
        for cell in ws[start_row+1]:
            cell.font = Font(bold=True)
        start_row += 2
        for _, row in df.iterrows():
            ws.append([row['Project Code'], row['Project'], None])
        data_start = start_row
        data_end = start_row + len(df) - 1
        for i, (_, row) in enumerate(df.iterrows(), start=data_start):
            formula = row['Formula'] if pd.notna(row['Formula']) else None
            ws[f"C{i}"] = formula if formula else row['Total hrs']
            ws[f"D{i}"] = f"=C{i}/8"
            ws[f"E{i}"] = "=$B$8"; ws[f"E{i}"].number_format = u'€#,##0.00'
            ws[f"F{i}"] = f"=D{i}*E{i}"; ws[f"F{i}"].number_format = u'€#,##0.00'
        ws[f"E{data_end+1}"] = "Subtotal:"; ws[f"E{data_end+1}"].font = Font(bold=True)
        ws[f"F{data_end+1}"] = f"=SUM(F{data_start}:F{data_end})"
        ws[f"F{data_end+1}"].number_format = u'€#,##0.00'
        ws[f"F{data_end+1}"].font = Font(bold=True)
        return data_end + 3, data_end + 1

    r, pcg_sub = write_table(df_final[df_final['Company'] == 'PCG'], 12, "PCG Projects")
    r, pcr_sub = write_table(df_final[df_final['Company'] == 'PCR'], r, "PCR Projects")
    ws[f"E{r+2}"] = "Grand Total:"; ws[f"E{r+2}"].font = Font(bold=True)
    ws[f"F{r+2}"] = f"=F{pcg_sub}+F{pcr_sub}"; ws[f"F{r+2}"].number_format = u'€#,##0.00'; ws[f"F{r+2}"].font = Font(bold=True)

    for col in ws.columns:
        width = max(len(str(cell.value)) if cell.value else 0 for cell in col) + 2
        ws.column_dimensions[get_column_letter(col[0].column)].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, person_name, time_period

# --- Streamlit UI ---
st.title("Invoice Generator")

st.write("""
1. Export your Float data: first Person view, then Project view.
2. Upload both CSVs below, enter your monthly salary.
3. Click generate to download your invoice.
""")

with st.form("input_form"):
    timesheet_file = st.file_uploader("Upload 'Your Name-Table-...'.csv", type="csv")
    projects_file = st.file_uploader("Upload 'Projects-Table...csv'", type="csv")
    monthly_salary = st.number_input("Monthly Salary (€)", step=100.0)
    generate_button = st.form_submit_button("Generate Invoice")

if generate_button and timesheet_file and projects_file and monthly_salary > 0:
    result, person_name, time_period = generate_invoice(timesheet_file, projects_file, monthly_salary)
    st.download_button("Download Invoice", result, file_name=f"Invoice_{person_name}_{time_period}.xlsx")
