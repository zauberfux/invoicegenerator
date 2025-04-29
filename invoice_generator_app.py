import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
import re
from datetime import datetime

# Helper functions
def generate_invoice(timesheet_file, projects_file, monthly_salary):
    df_time_raw = pd.read_csv(timesheet_file)
    df_projects = pd.read_csv(projects_file)

    # Extract basic info
    person_name = df_time_raw['Person'].iloc[0]
    dept_raw = df_time_raw['Department'].iloc[0]
    dept_num = re.search(r'\d+', str(dept_raw)).group()

    # Extract 'All' Logged hrs
    all_logged_hrs_value = df_time_raw[df_time_raw['Project'] == 'All']['Logged hrs'].sum()

    # Extract time period
    file_name = timesheet_file.name
    time_period_raw = file_name.split('-Table-')[1].replace('.csv','')
    year = time_period_raw[:4]
    month = time_period_raw[4:6]
    time_period = datetime.strptime(f"{year}-{month}-01", "%Y-%m-%d").strftime("%B %Y")

    # Prepare clean project table
    df_time = df_time_raw[df_time_raw['Project'] != 'All'].copy()

    def merge_project_row(row):
        if pd.notna(row['Project']) and str(row['Project']).strip():
            return str(row['Project']).strip()
        elif pd.notna(row['Time off']) and str(row['Time off']).strip():
            return 'Sick leave' if str(row['Time off']).strip() == 'Krankheit' else str(row['Time off']).strip()
        elif pd.notna(row['Holiday']) and str(row['Holiday']).strip():
            return str(row['Holiday']).strip()
        else:
            return None

    df_time['Merged Project'] = df_time.apply(merge_project_row, axis=1)
    df_time['Total hrs'] = df_time['Logged hrs'].fillna(0) + df_time['Time off hrs'].fillna(0) + df_time['Holiday hrs'].fillna(0)
    df_time = df_time[df_time['Total hrs'] != 0]

    def clean_project_name(proj):
        if isinstance(proj, str):
            proj = proj.strip()
            if proj.startswith(('Admin', 'BF')):
                return 'Admin'
            if proj.startswith('Sales_BF'):
                return proj  # preserve Sales_BFxx for reassignment
            return proj
        return proj

    df_time['Project_clean'] = df_time['Merged Project'].apply(clean_project_name)
    df_summary = df_time.groupby('Project_clean', as_index=False)['Total hrs'].sum()

    # Map project codes from uploaded file
    project_code_map = dict(zip(df_projects['Project'], df_projects['Project code']))

    def generate_bf_general_code(dept_num: str, company_prefix: str):
        return f"{company_prefix}{dept_num.zfill(2)}000"

    # Manual code overrides
    manual_codes = {
        f'BF{dept_num} General (PCR)': generate_bf_general_code(dept_num, '2'),
        f'BF{dept_num} General (PCG)': generate_bf_general_code(dept_num, '1'),
    }

    # Handle Sales_BFxx reassignment to BFxx General (PCG/PCR)
    sales_to_general_rows = []

    for idx, row in df_summary.iterrows():
        project = row['Project_clean']
        total_hrs = row['Total hrs']

        if project.startswith('Sales_BF'):
            bf_code = project.replace('Sales_BF', '').strip()
            general_pc = f'BF{bf_code} General (PCG)'
            general_pr = f'BF{bf_code} General (PCR)'
            hrs_half = total_hrs / 2
            sales_to_general_rows.append({'Project_clean': general_pc, 'Total hrs': hrs_half})
            sales_to_general_rows.append({'Project_clean': general_pr, 'Total hrs': hrs_half})

    df_summary = df_summary[~df_summary['Project_clean'].str.startswith('Sales_BF')]
    df_sales_to_general = pd.DataFrame(sales_to_general_rows)
    df_summary = pd.concat([df_summary, df_sales_to_general], ignore_index=True)

    sickleave_hours = df_summary[df_summary['Project_clean'] == 'Sick leave']['Total hrs'].sum()
    holiday_hours = df_summary[df_summary['Project_clean'] == 'Holiday']['Total hrs'].sum()
    admin_hours = df_summary[df_summary['Project_clean'] == 'Admin']['Total hrs'].sum()
    bf_general_hours = sickleave_hours + holiday_hours + admin_hours

    df_bf_split = pd.DataFrame({
        'Project_clean': [f'BF{dept_num} General (PCR)', f'BF{dept_num} General (PCG)'],
        'Total hrs': [bf_general_hours/2, bf_general_hours/2]
    })

    df_final = pd.concat([
        df_summary,
        df_bf_split
    ], ignore_index=True)

    df_final['Project Code'] = df_final['Project_clean'].map(lambda x: project_code_map.get(x, manual_codes.get(x, '')))
    df_final['Days'] = df_final['Total hrs'] / 8

    df_pcg = df_final[df_final['Project Code'].astype(str).str.startswith('1')]
    df_pcr = df_final[df_final['Project Code'].astype(str).str.startswith('2')]

    # Prepare workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    # Write header
    ws['A1'] = "Name:"; ws['B1'] = person_name
    ws['A2'] = "Business Field:"; ws['B2'] = dept_num
    ws['A3'] = "Time Period:"; ws['B3'] = time_period
    ws['A4'] = "Monthly Salary:"; ws['B4'] = monthly_salary
    ws['A5'] = "Number of Days Worked:"; ws['B5'] = f"={all_logged_hrs_value}/8"
    ws['A6'] = "Sick leave days:"; ws['B6'] = round(sickleave_hours/8,3)
    ws['A7'] = "Holiday days:"; ws['B7'] = round(holiday_hours/8,3)
    ws['A8'] = "Total days:"; ws['B8'] = "=B5+B6+B7"
    ws['A9'] = "Day Rate:"; ws['B9'] = "=B4/B8"

    ws['B4'].number_format = u'€#,##0.00'
    ws['B9'].number_format = u'€#,##0.00'

    # Write table function
    def write_table(df, start_row, title):
        ws[f"A{start_row}"] = title
        ws[f"A{start_row}"].font = Font(bold=True)
        ws.append(['Project Code', 'Project', 'Logged hrs', 'Days', 'Day Rate', 'Costs'])
        for cell in ws[start_row+1]:
            cell.font = Font(bold=True)
        start_row += 2

        for idx, row in df.iterrows():
            ws.append([row['Project Code'], row['Project_clean'], row['Total hrs']])

        data_start = start_row
        data_end = start_row + len(df) -1

        for row in range(data_start, data_end+1):
            ws[f"D{row}"] = f"=C{row}/8"
            ws[f"E{row}"] = "=$B$9"; ws[f"E{row}"].number_format = u'€#,##0.00'
            ws[f"F{row}"] = f"=D{row}*E{row}"; ws[f"F{row}"].number_format = u'€#,##0.00'

        subtotal_row = data_end+1
        ws[f"E{subtotal_row}"] = "Subtotal:"
        ws[f"E{subtotal_row}"].font = Font(bold=True)
        ws[f"F{subtotal_row}"] = f"=SUM(F{data_start}:F{data_end})"
        ws[f"F{subtotal_row}"].number_format = u'€#,##0.00'
        ws[f"F{subtotal_row}"].font = Font(bold=True)

        return subtotal_row + 2, subtotal_row

    current_row = 12
    current_row, pcg_subtotal = write_table(df_pcg, current_row, "PCG Projects")
    current_row, pcr_subtotal = write_table(df_pcr, current_row, "PCR Projects")

    ws[f"E{current_row}"] = "Grand Total:"
    ws[f"E{current_row}"].font = Font(bold=True)
    ws[f"F{current_row}"] = f"=F{pcg_subtotal}+F{pcr_subtotal}"
    ws[f"F{current_row}"].number_format = u'€#,##0.00'
    ws[f"F{current_row}"].font = Font(bold=True)

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

# Streamlit App
st.title("Invoice Generator")
st.write("Go to Float, select Person (You!) and Time period (month). Move the Person/Projects slider to 'Person' and 'Export Table Data' for your time tracking data. Now move the slider to 'Projects' and 'Export Table Data' again, this time for exporting the Project codes. Upload both CSV files, enter your Monthly Salary, and download your invoice!")

with st.form("input_form"):
    timesheet_file = st.file_uploader("Upload Your Name-Table-... CSV", type="csv")
    projects_file = st.file_uploader("Upload Projects-Table... CSV", type="csv")
    monthly_salary = st.number_input("Monthly Salary (€)", step=100.0)
    generate_button = st.form_submit_button("Generate Invoice")

if generate_button and timesheet_file and projects_file and monthly_salary > 0:
    result = generate_invoice(timesheet_file, projects_file, monthly_salary)
    st.download_button("Download Invoice Excel", result, file_name="Generated_Invoice.xlsx")
