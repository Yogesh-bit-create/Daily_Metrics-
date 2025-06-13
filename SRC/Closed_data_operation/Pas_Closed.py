# -*- coding: utf-8 -*-
"""
Created on Fri May 23 00:15:36 2025

@author: yogesh.sanjay.gavade
"""

# === src/closed/pas.py ===
import pandas as pd
import time
from config import CLOSED_INC_FILE, METRICS_FILE

def run():
    start_time = time.time()

    # Step 1: Load the closed incident file
    df = pd.read_excel(CLOSED_INC_FILE)

    # Step 2: Filter based on PAS-specific SLA definitions
    sla_definitions = ["PAS Initiated RUSH Incidents", "PAS Helpdesk Incident Closure"]
    filtered_df = df[df['SLA definition'].isin(sla_definitions)].copy()

    # Step 3: Add formulas for time and SLA status
    filtered_df['Business elapsed time (Days)'] = [
        f'=J{row}/60/60/24' for row in range(2, len(filtered_df) + 2)
    ]

    filtered_df['SLA Status for PAS Helpdesk'] = [
        f'=IF(U{row}<=10,"Met 10 Days SLA",IF(AND(U{row}>10,U{row}<=15),"Met 15 Days SLA","Missed PAS SLA"))'
        for row in range(2, len(filtered_df) + 2)
    ]

    filtered_df['SLA Status for RUSH'] = [
        f'=IF(U{row}<=5,"Met 5 Days SLA","Missed SLA")' for row in range(2, len(filtered_df) + 2)
    ]

    # Step 4: Write to 'PAS Closed' sheet
    with pd.ExcelWriter(METRICS_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        filtered_df.to_excel(writer, sheet_name='PAS Closed', index=False)

    elapsed_time = time.time() - start_time
    print(f"✅ PAS Closed module completed in {elapsed_time:.2f} seconds.")
