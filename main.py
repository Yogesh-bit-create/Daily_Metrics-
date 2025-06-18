# -*- coding: utf-8 -*-
"""
Created on Thu May 22 16:14:00 2025

@author: yogesh.sanjay.gavade
"""

# === main.py ===

from src.closed import incident as closed_incident
from src.closed import prb as closed_prb
from src.closed import pas as pas_closed

from src.open import incident as open_incident
from src.open import prb as open_prb
from src.open import premium as premium_app

from src import change_request
from src import last_updated_incident
from src import prb_categorization

import time


try :


  def main():
    start_time = time.time()
    print("🚀 Starting Daily Metrics Pipeline...\n")

    # === Run scripts in correct order ===

    # Closed flow
    closed_incident.run()
    closed_prb.run()
    pas_closed.run()

    # Open flow
    open_incident.run()
    open_prb.run()
    premium_app.run()

    # Independent modules
    change_request.run()
    last_updated_incident.run()
    prb_categorization.run()

    print("\n✅ Pipeline completed successfully.")
    print(f"⏱️ Total time: {time.time() - start_time:.2f} seconds")

finally:
  print("please check a file ")
  print("hello")

if __name__ == "__main__":
    main()
