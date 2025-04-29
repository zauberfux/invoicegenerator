import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

# --- Core invoice generation logic ---
def generate_invoice(timesheet_file, projects_file, monthly_salary):
    df_time_raw = pd.read_csv(timesheet_file)
    df_projects = pd.read_csv(projects_file)

    person_name = df_time_raw['Person'].iloc[0]
    dept_raw = df_time_raw['Department'].iloc[0]
    dept_num = re.search(r'\d+', str(dept_raw)).group()
    all_logged_hrs_value = df_time_raw[df_time_raw['Project'] == 'All']['Logged hrs'].sum()

    filename = timesheet_file.name
    match = re.search(r'Table-(\d{8})', filename)
    year, month = match.group(1)[:4], match.group(1)[4:6]
    time_period = datetime.strptime(f"{year}-{month}-01", "%Y-%m-%d").strftime("%B %Y")

    def resolve_project(row):
        if pd.notna(row['Project']) and str(row['Project']).strip():
            return str(row['Project']).strip()
        elif pd.notna(row['Time off']) and str(row['Time off']).strip():
            return 'Sick leave' if row['Time off'] == 'Krankheit' else str(row['Time off']).strip()
        elif pd.notna(row['Holiday']) and str(row['Holiday']).strip():
            return str(row['Holiday']).strip()
        return None

    df_time = df_time_raw[df_time_raw['Project'] != 'All'].copy()
    df_time['Merged Project'] = df_time.apply(resolve_project, axis=1)
    df_time['Total hrs'] = df_time['Logged hrs'].fillna(0) + df_time['Time off hrs'].fillna(0) + df_time['Holiday hrs'].fillna(0)
    df_time = df_time[df_time['Total hrs'] > 0]

    def clean_name(proj):
        if proj == 'BF coordination':
            return 'Admin'
        return proj

    df_time['Project_clean'] = df_time['Merged Project'].apply(clean_name)
    df_time = df_time[df_time['Project_clean'] != 'Sales']

    project_code_map = dict(zip(df_projects['Project'], df_projects['Project code']))
    df_summary = df_time.groupby('Project_clean', as_index=False)['Total hrs'].sum()

    # Sales_BFxx splitting
    df_sales = df_summary[df_summary['Project_clean'].str.startswith('Sales_BF', na=False)].copy()
    sales_rows = []
    for _, row in df_sales.iterrows():
        match = re.search(r'Sales_BF(\d{2})', row['Project_clean'])
        if match:
            bf = match.group(1)
            hrs = row['Total hrs'] / 2
            sales_rows.extend([
                {'Project_clean': f'BF{bf} General (PCG)', 'Total hrs': hrs, 'Project Code': f'1{bf}000'},
                {'Project_clean': f'BF{bf} General (PCR)', 'Total hrs': hrs, 'Project Code': f'2{bf}000'}
            ])
    df_sales_split = pd.DataFrame(sales_rows)

    # Admin/Sick/Holiday/BF coordination to BFXX General
    is_bf_general = (
        df_time['Project_clean'].fillna('').str.startswith('Admin') |
        df_time['Project_clean'].isin(['Sick leave', 'Holiday', 'BF coordination'])
    )
    bf_general_hours = df_time[is_bf_general]['Total hrs'].sum()
    df_bf_split = pd.DataFrame({
        'Project_clean': [f'BF{dept_num} General (PCG)', f'BF{dept_num} General (PCR)'],
        'Total hrs': [bf_general_hours/2, bf_general_hours/2],
        'Project Code': [f'1{dept_num}000', f'2{dept_num}000']
    })

    # Regular projects
    is_regular = ~df_time['Project_clean'].str.startswith('Sales_BF', na=False) & ~is_bf_general
    df_regular = df_time[is_regular].groupby('Project_clean', as_index=False)['Total hrs'].sum()
    df_regular['Project Code'] = df_regular['Project_clean'].map(project_code_map)

    # Combine all
    df_final = pd.concat([df_regular, df_sales_split, df_bf_split], ignore_index=True)
    df_final = df_final.groupby(['Project_clean', 'Project Code'], as_index=False)['Total hrs'].sum()
    df_final['Days'] = df_final['Total hrs'] / 8

    df_pcg = df_final[df_final['Project Code'].astype(str).str.startswith('1')]
    df_pcr = df_final[df_final['Project Code'].astype(str).str.startswith('2')]

    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    ws['A1'] = "Name:"; ws['B1'] = person_name
    ws['A2'] = "Business Field:"; ws['B2'] = dept_num
    ws['A3'] = "Time Period:"; ws['B3'] = time_period
    ws['A4'] = "Monthly Salary:"; ws['B4'] = monthly_salary
    ws['A5'] = "Number of Days Worked:"; ws['B5'] = f"={all_logged_hrs_value}/8"
    ws['A6'] = "Sick leave days:"; ws['B6'] = round(df_time[df_time['Project_clean'] == 'Sick leave']['Total hrs'].sum()/8,3)
    ws['A7'] = "Holiday days:"; ws['B7'] = round(df_time[df_time['Project_clean'] == 'Holiday']['Total hrs'].sum()/8,3)
    ws['A8'] = "Total days:"; ws['B8'] = "=B5+B6+B7"
    ws['A9'] = "Day Rate:"; ws['B9'] = "=B4/B8"

    def write_table(df, start_row, title):
        ws[f"A{start_row}"] = title
        ws[f"A{start_row}"].font = Font(bold=True)
        ws.append(['Project Code', 'Project', 'Logged hrs', 'Days', 'Day Rate', 'Costs'])
        for cell in ws[start_row+1]:
            cell.font = Font(bold=True)
        start_row += 2

        for _, row in df.iterrows():
            ws.append([row['Project Code'], row['Project_clean'], row['Total hrs']])

        data_start = start_row
        data_end = start_row + len(df) - 1

        for r in range(data_start, data_end+1):
            ws[f"D{r}"] = f"=C{r}/8"
            ws[f"E{r}"] = "=$B$9"; ws[f"E{r}"].number_format = u'€#,##0.00'
            ws[f"F{r}"] = f"=D{r}*E{r}"; ws[f"F{r}"].number_format = u'€#,##0.00'

        ws[f"E{data_end+1}"] = "Subtotal:"; ws[f"E{data_end+1}"].font = Font(bold=True)
        ws[f"F{data_end+1}"] = f"=SUM(F{data_start}:F{data_end})"
        ws[f"F{data_end+1}"] .number_format = u'€#,##0.00'
        ws[f"F{data_end+1}"].font = Font(bold=True)

        return data_end + 3, data_end + 1

    r, pcg_subtotal = write_table(df_pcg, 12, "PCG Projects")
    r, pcr_subtotal = write_table(df_pcr, r, "PCR Projects")
    ws[f"E{r}"] = "Grand Total:"; ws[f"E{r}"].font = Font(bold=True)
    ws[f"F{r}"] = f"=F{pcg_subtotal}+F{pcr_subtotal}"; ws[f"F{r}"].number_format = u'€#,##0.00'; ws[f"F{r}"].font = Font(bold=True)

    for col in ws.columns:
        width = max(len(str(cell.value)) if cell.value else 0 for cell in col) + 2
        ws.column_dimensions[get_column_letter(col[0].column)].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Streamlit App ---
st.title("Invoice Generator")
st.write("Upload your timesheet CSV and project CSV. Set your monthly salary and download your generated invoice!")

with st.form("input_form"):
    timesheet_file = st.file_uploader("Upload Timesheet CSV", type="csv")
    projects_file = st.file_uploader("Upload Projects CSV", type="csv")
    monthly_salary = st.number_input("Monthly Salary (€)", step=100.0)
    generate_button = st.form_submit_button("Generate Invoice")

if generate_button and timesheet_file and projects_file and monthly_salary > 0:
    result = generate_invoice(timesheet_file, projects_file, monthly_salary)
    st.download_button("Download Invoice", result, file_name="Generated_Invoice.xlsx")
