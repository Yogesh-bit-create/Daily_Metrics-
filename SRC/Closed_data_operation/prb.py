# -*- coding: utf-8 -*-
"""
Created on Thu May 22 16:42:49 2025

@author: yogesh.sanjay.gavade
"""

# === src/closed/prb.py ===
import time
import pandas as pd
from openpyxl import load_workbook
from config import CLOSED_INC_FILE, METRICS_FILE
from src.utils.excel_utils import (
    copy_filtered_rows,
    apply_portfolio_lookup,
    apply_quarterly_sla_filter
)

def run():
    start_time = time.time()

    # Step 1: Copy filtered PRB rows from Closed Tickets
    copy_filtered_rows(
        source_file=CLOSED_INC_FILE,
        target_file=METRICS_FILE,
        target_sheet="Closed PRBs",
        starts_with="PRB"
    )

    # Step 2: Apply VLOOKUP-like mapping to Portfolio and Q/N columns
    wb = load_workbook(METRICS_FILE)
    sheet = wb['Closed PRBs']
    portfolio_sheet = wb['AG to Portfolio mapping']
    apply_portfolio_lookup(sheet, portfolio_sheet, target_col=21, match_col_index=1)  # Portfolio
    apply_portfolio_lookup(sheet, portfolio_sheet, target_col=20, match_col_index=1, value_col=4)  # Q/N
    wb.save(METRICS_FILE)

    print("✅ VLOOKUPs applied to Closed PRBs sheet.")

    # Step 3: Load DataFrame and apply duplicate cleanup logic
    df = pd.read_excel(METRICS_FILE, sheet_name="Closed PRBs")
    qn_sla_pairs = [
        ('Quarterly', 'Accenture_P4 Defect Closure (Quarterly)_App Dev SLA'),
        ('Non-Quterly', 'Accenture_P4 Defect Closure (Non-Quarterly)_App Dev SLA')
    ]

    final_df = df.copy()
    for qn_value, sla_def in qn_sla_pairs:
        final_df = apply_quarterly_sla_filter(final_df, qn_value, sla_def)

    # Step 4: Add SLA Status column
    final_df['SLA Status'] = [f'=IF(L{row}<100,"Met","Not Met")' for row in range(2, len(final_df) + 2)]

    with pd.ExcelWriter(METRICS_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        final_df.to_excel(writer, sheet_name='Closed PRBs', index=False)

    elapsed_time = time.time() - start_time
    print(f"✅ Closed PRB module completed in {elapsed_time:.2f} seconds.")
