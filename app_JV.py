from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from datetime import datetime
import pandas as pd
import uvicorn
import tempfile
import shutil
import os
import re

app = FastAPI()

# Allow CORS if using frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"]
)

def clean_column_name(col):
    return re.sub(r'[^\w]', '', str(col).lower().strip())

def validate_date(date_str):
    try:
        date_obj = datetime.strptime(date_str, "%d/%m/%y")
        month_year = date_obj.strftime("%B %Y").upper()
        return date_obj, month_year
    except ValueError:
        raise ValueError("Invalid date format. Please use (dd/mm/yy).")

def get_gl_account(description, department_code):
    if str(int(department_code)) in ['3003', '3006']:
        account_map = {
            'TOTAL BASIC SALARY': '640100',
            'Retroactive Appraisal/Arrears': '640100',
            'MONTHLY FOOD': '640140',
            'MONTHLY TRANSP': '640142',
            'MONTHLY HOUSING': '640143',
            'MONTHLY OTHER ALL': '640141',
            'Educatin All': '701200',
            'MONTHLY OVER TIME': '640120'
        }
    else:
        account_map = {
            'TOTAL BASIC SALARY': '701100',
            'Retroactive Appraisal/Arrears': '701100',
            'MONTHLY FOOD': '701210',
            'MONTHLY TRANSP': '701220',
            'MONTHLY HOUSING': '701230',
            'MONTHLY OTHER ALL': '701200',
            'Educatin All': '701200',
            'MONTHLY OVER TIME': '701150'
        }

    base_description = re.sub(r'\s+\w+\s+\d{4}$', '', description).strip()
    return account_map.get(base_description, '')

@app.post("/upload/")
async def upload_excel(
    file: UploadFile = File(...),
    sheet_name: str = Form(...),
    posting_date: str = Form(...),
    journal_code: str = Form(...)
):
    try:
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        with temp_input as f:
            shutil.copyfileobj(file.file, f)

        try:
            excel_file = pd.ExcelFile(temp_input.name)
        except Exception:
            raise HTTPException(status_code=400, detail="Invalid Excel file format.")

        sheet_map = {sheet.strip(): sheet for sheet in excel_file.sheet_names}
        actual_sheet_name = sheet_map.get(sheet_name)
        if actual_sheet_name is None:
            raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found.")

        date_obj, month_year = validate_date(posting_date)

        df = pd.read_excel(temp_input.name, sheet_name=actual_sheet_name)
        df.columns = [clean_column_name(col) for col in df.columns]

        group_col = 'departmentcode'
        if group_col not in df.columns:
            raise HTTPException(status_code=400, detail="Missing 'Department Code' column.")

        agg_columns = {
            'totalbasicsalary': 'TOTAL BASIC SALARY',
            'retroactiveappraisalarrears': 'Retroactive Appraisal/Arrears',
            'monthlyfood': 'MONTHLY FOOD',
            'monthlytransp': 'MONTHLY TRANSP',
            'monthlyhousing': 'MONTHLY HOUSING',
            'monthlyotherall': 'MONTHLY OTHER ALL',
            'educatinall': 'Educatin All',
            'monthlyovertime': 'MONTHLY OVER TIME'
        }

        available_agg = {k: v for k, v in agg_columns.items() if k in df.columns}
        if not available_agg:
            raise HTTPException(status_code=400, detail="No required aggregation columns found.")

        for col in available_agg.keys():
            df[col] = pd.to_numeric(df[col], errors='coerce')

        agg_df = df.groupby(group_col)[list(available_agg.keys())].sum().reset_index()
        melted_df = agg_df.melt(
            id_vars=[group_col],
            value_vars=list(available_agg.keys()),
            var_name='Description',
            value_name='Amount'
        )
        melted_df['Description'] = melted_df['Description'].map(available_agg) + " " + month_year
        melted_df['Account Number'] = melted_df.apply(
            lambda row: get_gl_account(row['Description'], row[group_col]), axis=1
        )
        melted_df['Posting Date'] = posting_date
        melted_df['Journal Code'] = journal_code
        melted_df['G/L Account'] = "G/L Account"
        melted_df['Department Code'] = melted_df[group_col]

        final_cols = [
            'Posting Date', 'Journal Code', 'G/L Account',
            'Department Code', 'Account Number', 'Description', 'Amount'
        ]
        melted_df = melted_df[melted_df['Amount'] != 0]
        melted_df = melted_df[final_cols].sort_values(by=['Department Code', 'Description'])

        temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        melted_df.to_excel(temp_output.name, index=False)

        return FileResponse(path=temp_output.name, filename="processed_output.xlsx")

    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    finally:
        if os.path.exists(temp_input.name):
            os.remove(temp_input.name)
