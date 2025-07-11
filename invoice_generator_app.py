import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

def generate_invoice(timesheet_file, projects_file, monthly_salary):
    df_time_raw = pd.read_csv(timesheet_file)
    df_projects = pd.read_csv(projects_file)

    person_name = df_time_raw['Person'].iloc[0]
    dept_raw = df_time_raw['Department'].iloc[0]
    dept_num = re.search(r'\d+', str(dept_raw)).group()
    project_code_map = dict(zip(df_projects['Project'], df_projects['Project code']))

    # Total hrs per line
    df_time_raw['Total hrs per line'] = (
        df_time_raw['Logged hrs'].fillna(0) +
        df_time_raw['Time off hrs'].fillna(0) +
        df_time_raw['Holiday hrs'].fillna(0)
    )
    df_time = df_time_raw[df_time_raw['Total hrs per line'] > 0].copy()

    # Overtime
    overtime_hrs = df_time_raw[
        df_time_raw['Time off'].fillna("").str.contains("ausgleich für zusätzliche arbeitszeit", case=False)
    ]['Time off hrs'].sum()

    all_logged_hrs = df_time_raw['Logged hrs'].sum()
    all_paid_timeoff_hrs = (
    df_time_raw['Time off hrs'].sum() +
    df_time_raw['Holiday hrs'].sum() -
    overtime_hrs
)

    )

    # Time period
    filename = timesheet_file.name
    match = re.search(r'Table-(\d{8})', filename)
    year, month = match.group(1)[:4], match.group(1)[4:6]
    time_period = datetime.strptime(f"{year}-{month}-01", "%Y-%m-%d").strftime("%B %Y")

    # Admin/internal hours
    admin_keywords = ['Admin', 'BF coordination', 'HR', 'IT', 'Finances']
    admin_time_own_bf = df_time[
        df_time['Project'].str.startswith(tuple(admin_keywords), na=False)
    ]['Logged hrs'].fillna(0).sum()

    bf_general_hours = admin_time_own_bf + all_paid_timeoff_hrs
    df_bf_split = pd.DataFrame({
        'Project': [f'BF{dept_num} General (PCG)', f'BF{dept_num} General (PCR)'],
        'Total hrs': [bf_general_hours / 2, bf_general_hours / 2],
        'Project Code': [f'1{dept_num}000', f'2{dept_num}000'],
        'Company': ['PCG', 'PCR']
    })

    # Sales split
    df_summary = df_time.groupby('Project', as_index=False)['Logged hrs'].sum()
    df_sales = df_summary[df_summary['Project'].str.startswith('Sales_BF', na=False)].copy()
    sales_rows = []
    for _, row in df_sales.iterrows():
        match = re.search(r'Sales_BF(\d{2})', row['Project'])
        if match:
            bf = match.group(1)
            hrs = row['Logged hrs'] / 2
            sales_rows.extend([
                {'Project': f'BF{bf} General (PCG)', 'Total hrs': hrs, 'Project Code': f'1{bf}000', 'Company': 'PCG'},
                {'Project': f'BF{bf} General (PCR)', 'Total hrs': hrs, 'Project Code': f'2{bf}000', 'Company': 'PCR'}
            ])
    df_sales_split = pd.DataFrame(sales_rows)

    # Regular projects with two-digit prefix
    df_regular = df_time[df_time['Project'].fillna("").str.match(r"^\d{2}_")].copy()

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

    df_regular[['Project Code', 'Company']] = df_regular.apply(resolve_project_code_and_company, axis=1)
    df_regular = df_regular.groupby(['Project', 'Project Code', 'Company'], as_index=False)['Logged hrs'].sum()
    df_regular = df_regular.rename(columns={'Logged hrs': 'Total hrs'})

    # Combine all
    df_final = pd.concat([df_regular, df_sales_split, df_bf_split], ignore_index=True)
    df_final = df_final.groupby(['Project', 'Project Code', 'Company'], as_index=False)['Total hrs'].sum()
    df_final['Days'] = df_final['Total hrs'] / 8

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
            ws.append([row['Project Code'], row['Project'], row['Total hrs']])
        data_start = start_row
        data_end = start_row + len(df) - 1
        for r in range(data_start, data_end+1):
            ws[f"D{r}"] = f"=C{r}/8"
            ws[f"E{r}"] = "=$B$8"; ws[f"E{r}"].number_format = u'€#,##0.00'
            ws[f"F{r}"] = f"=D{r}*E{r}"; ws[f"F{r}"].number_format = u'€#,##0.00'
        ws[f"E{data_end+1}"] = "Subtotal:"; ws[f"E{data_end+1}"].font = Font(bold=True)
        ws[f"F{data_end+1}"] = f"=SUM(F{data_start}:F{data_end})"
        ws[f"F{data_end+1}"].number_format = u'€#,##0.00'
        ws[f"F{data_end+1}"].font = Font(bold=True)
        return data_end + 3, data_end + 1

    r, pcg_subtotal = write_table(df_final[df_final['Company'] == 'PCG'], 12, "PCG Projects")
    r, pcr_subtotal = write_table(df_final[df_final['Company'] == 'PCR'], r, "PCR Projects")
    ws[f"E{r}"] = "Grand Total:"; ws[f"E{r}"].font = Font(bold=True)
    ws[f"F{r}"] = f"=F{pcg_subtotal}+F{pcr_subtotal}"; ws[f"F{r}"].number_format = u'€#,##0.00'; ws[f"F{r}"].font = Font(bold=True)

    for col in ws.columns:
        width = max(len(str(cell.value)) if cell.value else 0 for cell in col) + 2
        ws.column_dimensions[get_column_letter(col[0].column)].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, person_name, time_period

# --- Streamlit App UI ---
st.title("Invoice Generator")

st.write("""
1. Go to Float and select your name and the relevant month.
2. Set the view to **Person** and export the **Table Data** CSV.
3. Then switch to **Projects** view and export the **Project Table Data** CSV.
4. Upload both files below, input your salary, and generate your invoice.
""")

with st.form("input_form"):
    timesheet_file = st.file_uploader("Upload 'Your Name-Table-...'.csv", type="csv")
    projects_file = st.file_uploader("Upload 'Projects-Table...csv'", type="csv")
    monthly_salary = st.number_input("Monthly Salary (€)", step=100.0)
    generate_button = st.form_submit_button("Generate Invoice")

if generate_button and timesheet_file and projects_file and monthly_salary > 0:
    result, person_name, time_period = generate_invoice(timesheet_file, projects_file, monthly_salary)
    st.download_button("Download Invoice", result, file_name=f"Invoice_{person_name}_{time_period}.xlsx")
