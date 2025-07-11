import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from decimal import Decimal
from collections import defaultdict

def generate_invoice(timesheet_file, projects_file, monthly_salary):
    df_time_raw = pd.read_csv(timesheet_file)
    df_projects = pd.read_csv(projects_file)

    person_name = df_time_raw['Person'].iloc[0]
    dept_raw = df_time_raw['Department'].iloc[0]
    dept_num = re.search(r'\d+', str(dept_raw)).group()
    project_code_map = dict(zip(df_projects['Project'], df_projects['Project code']))

    overtime_hrs = df_time_raw[
        df_time_raw['Time off'].fillna("").str.contains("ausgleich für zusätzliche arbeitszeit", case=False)
    ]['Time off hrs'].sum()

    all_logged_hrs = df_time_raw['Logged hrs'].sum()
    all_paid_timeoff_hrs = (
        df_time_raw['Time off hrs'].sum() +
        df_time_raw['Holiday hrs'].sum() -
        overtime_hrs
    )

    match = re.search(r'Table-(\d{8})', timesheet_file.name)
    year, month = match.group(1)[:4], match.group(1)[4:6]
    time_period = datetime.strptime(f"{year}-{month}-01", "%Y-%m-%d").strftime("%B %Y")

    admin_keywords = ['Admin', 'BF coordination', 'IT', 'Finances']
    pc_keywords = ['HR', 'P&C']
    sales_support_projects = ['Sales I proposal support', 'Tender Screening']

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

    df_billable_input = df_time_raw[df_time_raw['Project'].fillna('').str.match(r'^\d{2}_')].copy()
    df_billable_input[['Project Code', 'Company']] = df_billable_input.apply(resolve_project_code_and_company, axis=1)

    quota = df_billable_input.groupby('Company')['Logged hrs'].sum()
    total_real = quota.sum()
    pcg_ratio = float(quota.get('PCG', 0)) / total_real if total_real > 0 else 0.5
    pcr_ratio = float(quota.get('PCR', 0)) / total_real if total_real > 0 else 0.5

    def split_quota_rows(label, hrs, pcg_code, pcr_code):
        b = Decimal(str(hrs)).quantize(Decimal('0.0001'))
        pcg = Decimal(str(pcg_ratio)).quantize(Decimal('0.0001'))
        pcr = Decimal(str(pcr_ratio)).quantize(Decimal('0.0001'))
        return pd.DataFrame({
            'Project': [f'{label} (PCG)', f'{label} (PCR)'],
            'Project Code': [pcg_code, pcr_code],
            'Company': ['PCG', 'PCR'],
            'Total hrs': [None, None],
            'Formula': [f'={b}*{pcg}', f'={b}*{pcr}']
        })

    # 1. Own BF hours
    admin_hrs = df_time_raw[df_time_raw['Project'].str.startswith(tuple(admin_keywords), na=False)]['Logged hrs'].sum()
    paid_timeoff_hrs = (
        df_time_raw['Time off hrs'].sum() +
        df_time_raw['Holiday hrs'].sum() -
        overtime_hrs
    )
    own_bf_total = admin_hrs + paid_timeoff_hrs

    # 2. P&C + Sales Support
    pc_hrs = df_time_raw[df_time_raw['Project'].str.startswith(tuple(pc_keywords), na=False)]['Logged hrs'].sum()
    sales_support_hrs = df_time_raw[df_time_raw['Project'].isin(sales_support_projects)]['Logged hrs'].sum()

    # 3. Sales/Marketing redistribution
    df_sm = df_time_raw_
