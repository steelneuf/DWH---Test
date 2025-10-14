#!/usr/bin/env python3
"""
Direct entry point voor de Data Validation Tool.
Voorkomt de runpy warning en maakt het makkelijker om de app te starten.
"""

import sys
sys.dont_write_bytecode = True

from src.main import main

if __name__ == "__main__":
    main()
