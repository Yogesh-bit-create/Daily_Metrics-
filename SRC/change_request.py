# -*- coding: utf-8 -*-
"""
Created on Fri May 23 00:29:40 2025

@author: yogesh.sanjay.gavade
"""

# === src/change_request.py ===
import pandas as pd
import time
import warnings
from config import METRICS_FILE

CHANGE_TASK_FILE = r"C:\Excel\change_task.xlsx"
CHANGE_REQUEST_FILE = r"C:\Excel\change_request.xlsx"

def run():
    start_time = time.time()
    warnings.filterwarnings("ignore", message="Workbook contains no default sty")

    # Step 1: Build lookup dictionary from change_task.xlsx
    task_df = pd.read_excel(CHANGE_TASK_FILE, sheet_name='Page 1')
    lookup_dict = {str(row[0]).strip(): row[0] for _, row in task_df.iterrows()}

    # Step 2: Load the change request data
    df = pd.read_excel(CHANGE_REQUEST_FILE, sheet_name='Page 1')

    # Step 3: Annotate open task status
    df['Has Open Task'] = [
        lookup_dict.get(str(row[0]).strip(), "#N/A") for _, row in df.iterrows()
    ]

    df['Has Open Task'] = df['Has Open Task'].apply(
        lambda x: 'NO' if x == '#N/A' else 'YES' if str(x).startswith('CHG') else x
    )

    # Step 4: Add Ageing formula
    df['Ageing'] = [f'=TODAY()-(J{row})' for row in range(2, len(df) + 2)]

    # Step 5: Save to Metrics Trend_Date.xlsx
    with pd.ExcelWriter(METRICS_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name="Change Requests", index=False)

    elapsed_time = time.time() - start_time
    print(f"✅ Change Request module completed in {elapsed_time:.2f} seconds.")