# -*- coding: utf-8 -*-
"""
Created on Fri May 23 00:19:12 2025

@author: yogesh.sanjay.gavade
"""

# === src/open/premium.py ===
import pandas as pd
import time
from config import OPEN_INC_FILE, METRICS_FILE

def run():
    start_time = time.time()

    # Step 1: Load Open tickets file
    df = pd.read_excel(OPEN_INC_FILE)

    # Step 2: Filter for Premium Processing App-related incidents
    sla_definition = "Premium Processing App- Open incident aging"
    assignment_group = "IT.A.TAP"

    filtered_df = df[
        (df['Assignment group'] == assignment_group) &
        (df['SLA definition'] == sla_definition)
    ].copy()

    if filtered_df.empty:
        print("⚠️ No matching data found for Premium App Ageing.")
    else:
        filtered_df['Business elapsed time (Days)'] = [
            f'=J{row}/60/60/24' for row in range(2, len(filtered_df) + 2)
        ]

        with pd.ExcelWriter(METRICS_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            filtered_df.to_excel(writer, sheet_name='Premium app ageing', index=False)

        print("✅ Data copied successfully into 'Premium app ageing'.")

    elapsed_time = time.time() - start_time
    print(f"✅ Premium App Ageing module completed in {elapsed_time:.2f} seconds.")