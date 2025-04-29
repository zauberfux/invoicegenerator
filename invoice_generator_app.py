import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import re
from datetime import datetime

def generate_invoice(timesheet_file, projects_file, monthly_salary):
    df_time_raw = pd.read_csv(timesheet_file)
    df_projects = pd.read_csv(projects_file)

    person_name = df_time_raw['Person'].iloc[0]
    dept_raw = df_time_raw['Department'].iloc[0]
    dept_num = re.search(r'\d+', str(dept_raw)).group()
    all_logged_hrs_value = df_time_raw[df_time_raw['Project'] == 'All']['Logged hrs'].sum()

    file_name = timesheet_file.name
    time_period_raw = file_name.split('-Table-')[1].replace('.csv','')
    year = time_period_raw[:4]
    month = time_period_raw[4:6]
    time_period = datetime.strptime(f"{year}-{month}-01", "%Y-%m-%d").strftime("%B %Y")

    df_time = df_time_raw[df_time_raw['Project'] != 'All'].copy()

    def merge_project_row(row):
        if pd.notna(row['Project']) and str(row['Project']).strip():
            return str(row['Project']).strip()
        elif pd.notna(row['Time off']) and str(row['Time off']).strip():
            return 'Sick leave' if str(row['Time off']).strip() == 'Krankheit' else str(row['Time off']).strip()
        elif pd.notna(row['Holiday']) and str(row['Holiday']).strip():
            return str(row['Holiday']).strip()
        return None

    df_time['Merged Project'] = df_time.apply(merge_project_row, axis=1)
    df_time['Total hrs'] = df_time['Logged hrs'].fillna(0) + df_time['Time off hrs'].fillna(0) + df_time['Holiday hrs'].fillna(0)
    df_time = df_time[df_time['Total hrs'] != 0]

    def clean_project_name(proj):
        if isinstance(proj, str):
            proj = proj.strip()
            if proj.startswith('Sales_BF'):
                return proj
            if proj.startswith(('Admin', 'BF')):
                return 'Admin'
            return proj
        return proj

    df_time['Project_clean'] = df_time['Merged Project'].apply(clean_project_name)
    df_summary = df_time.groupby('Project_clean', as_index=False)['Total hrs'].sum()

    project_code_map = dict(zip(df_projects['Project'], df_projects['Project code']))
    manual_codes = {}

    def generate_bf_general_code(num: str, prefix: str):
        return f"{prefix}{num.zfill(2)}000"

    # Reassign Sales_BFxx projects to BFxx General (PCG/PCR)
    sales_to_general = []
    for _, row in df_summary.iterrows():
        proj = row['Project_clean']
        hrs = row['Total hrs']
        if proj.startswith("Sales_BF"):
            bf_code = proj.replace("Sales_BF", "").strip()
            pcg_name = f"BF{bf_code} General (PCG)"
            pcr_name = f"BF{bf_code} General (PCR)"
            manual_codes[pcg_name] = generate_bf_general_code(bf_code, '1')
            manual_codes[pcr_name] = generate_bf_general_code(bf_code, '2')
            sales_to_general.append({'Project_clean': pcg_name, 'Total hrs': hrs / 2})
            sales_to_general.append({'Project_clean': pcr_name, 'Total hrs': hrs / 2})

    df_summary = df_summary[~df_summary['Project_clean'].str.startswith('Sales_BF')]
    df_sales = pd.DataFrame(sales_to_general)

    # Split Sick leave, Holiday, Admin to BFxx General (PCG/PCR)
    sick_hrs = df_summary[df_summary['Project_clean'] == 'Sick leave']['Total hrs'].sum()
    holi_hrs = df_summary[df_summary['Project_clean'] == 'Holiday']['Total hrs'].sum()
    admin_hrs = df_summary[df_summary['Project_clean'] == 'Admin']['Total hrs'].sum()
    bf_total = sick_hrs + holi_hrs + admin_hrs

    pcg_name = f"BF{dept_num} General (PCG)"
    pcr_name = f"BF{dept_num} General (PCR)"
    manual_codes[pcg_name] = generate_bf_general_code(dept_num, '1')
    manual_codes[pcr_name] = generate_bf_general_code(dept_num, '2')

    df_bf_split = pd.DataFrame({
        'Project_clean': [pcg_name, pcr_name],
        'Total hrs': [bf_total / 2, bf_total / 2]
    })

    # Combine and aggregate
    df_final = pd.concat([df_summary, df_sales, df_bf_split], ignore_index=True)
    df_final = df_final.groupby('Project_clean', as_index=False)['Total hrs'].sum()
    df_final['Project Code'] = df_final['Project_clean'].map(lambda x: project_code_map.get(x, manual_codes.get(x, '')))
    df_final['Days'] = df_final['Total hrs'] / 8

    df_pcg = df_final[df_final['Project Code'].astype(str).str.startswith('1')]
    df_pcr = df_final[df_final['Project Code'].astype(str).str.startswith('2')]

    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    ws['A1'] = "Name:"; ws['B1'] = person_name
    ws['A2'] = "Business Field:"; ws['B2'] = dept_num
    ws['A3'] = "Time Period:"; ws['B3'] = time_period
    ws['A4'] = "Monthly Salary:"; ws['B4'] = monthly_salary
    ws['A5'] = "Number of Days Worked:"; ws['B5'] = f"={all_logged_hrs_value}/8"
    ws['A6'] = "Sick leave days:"; ws['B6'] = round(sick_hrs/8,3)
    ws['A7'] = "Holiday days:"; ws['B7'] = round(holi_hrs/8,3)
    ws['A8'] = "Total days:"; ws['B8'] = "=B5+B6+B7"
    ws['A9'] = "Day Rate:"; ws['B9'] = "=B4/B8"
    ws['B4'].number_format = u'€#,##0.00'
    ws['B9'].number_format = u'€#,##0.00'

    def write_table(df, start_row, title):
        ws[f"A{start_row}"] = title
        ws[f"A{start_row}"].font = Font(bold=True)
        ws.append(['Project Code', 'Project', 'Logged hrs', 'Days', 'Day Rate', 'Costs'])
        for cell in ws[start_row + 1]:
            cell.font = Font(bold=True)
        start_row += 2
        for _, row in df.iterrows():
            ws.append([row['Project Code'], row['Project_clean'], row['Total hrs']])
        start = start_row
        end = start + len(df) - 1
        for r in range(start, end + 1):
            ws[f"D{r}"] = f"=C{r}/8"
            ws[f"E{r}"] = "=$B$9"; ws[f"E{r}"].number_format = u'€#,##0.00'
            ws[f"F{r}"] = f"=D{r}*E{r}"; ws[f"F{r}"].number_format = u'€#,##0.00'
        ws[f"E{end+1}"] = "Subtotal:"
        ws[f"E{end+1}"].font = Font(bold=True)
        ws[f"F{end+1}"] = f"=SUM(F{start}:F{end})"
        ws[f"F{end+1}"].number_format = u'€#,##0.00'
        ws[f"F{end+1}"].font = Font(bold=True)
        return end + 3, end + 1

    row, pcg_total = write_table(df_pcg, 12, "PCG Projects")
    row, pcr_total = write_table(df_pcr, row, "PCR Projects")
    ws[f"E{row}"] = "Grand Total:"
    ws[f"E{row}"].font = Font(bold=True)
    ws[f"F{row}"] = f"=F{pcg_total}+F{pcr_total}"
    ws[f"F{row}"].number_format = u'€#,##0.00'
    ws[f"F{row}"].font = Font(bold=True)

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit App Interface
st.title("Invoice Generator")
st.write("Export your timesheet from Float twice (Person-Table and Projects-Table), upload them here, and generate your monthly invoice.")

with st.form("input_form"):
    timesheet_file = st.file_uploader("Upload Your Name-Table CSV", type="csv")
    projects_file = st.file_uploader("Upload Projects-Table CSV", type="csv")
    monthly_salary = st.number_input("Monthly Salary (€)", step=100.0)
    generate_button = st.form_submit_button("Generate Invoice")

if generate_button and timesheet_file and projects_file and monthly_salary > 0:
    result = generate_invoice(timesheet_file, projects_file, monthly_salary)
    st.download_button("Download Invoice Excel", result, file_name="Generated_Invoice.xlsx")
