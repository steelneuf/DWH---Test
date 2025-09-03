"""
Data Validation Tool

Een tool voor het valideren van data uit verschillende bronnen (VR, VRR, DWH).
"""

import sys
sys.dont_write_bytecode = True

__version__ = "1.0.0"
__author__ = "Data Validation Team"

from .main import main, run
from .compare import compare_sources, normalize_key_value
from .config import read_kolommen_config, find_single_validation_config
from .InputOutput import load_sheet, create_excel_writer
from .utils import discover_input_sources
from .logging_setup import setup_logging

__all__ = [
    "main",
    "run", 
    "compare_sources",
    "normalize_key_value",
    "read_kolommen_config",
    "find_single_validation_config",
    "load_sheet",
    "create_excel_writer",
    "discover_input_sources",
    "setup_logging",
]
