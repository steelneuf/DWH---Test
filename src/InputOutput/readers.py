from pathlib import Path
from typing import Optional
import pandas as pd
import logging


def _read_csv_file(file_path: Path) -> pd.DataFrame:
    """Lees CSV bestand in als DataFrame."""
    try:
        return pd.read_csv(file_path, dtype=str)
    except Exception as exc:
        logging.getLogger("validation").warning(
            "Kon CSV bestand '%s' niet lezen: %s", file_path.name, exc
        )
        return pd.DataFrame()


def _read_excel_file(file_path: Path, sheet_name: str) -> pd.DataFrame:
    """Lees Excel bestand in als DataFrame."""
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl", dtype=str)
    except Exception as exc:
        logging.getLogger("validation").warning(
            "Kon Excel sheet '%s' uit bestand '%s' niet lezen: %s",
            sheet_name, file_path.name, exc
        )
        return pd.DataFrame()


def load_sheet(file_path: Path, sheet_name: str) -> pd.DataFrame:
    """Lees data in als DataFrame uit XLSX (specifieke sheet) of CSV.

    - Voor .xlsx: probeert sheet `sheet_name` te lezen.
    - Voor .csv: negeert `sheet_name` en leest het bestand als CSV.
    """
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        return _read_csv_file(file_path)
    else:
        return _read_excel_file(file_path, sheet_name)


def _read_csv_data(file: Path) -> Optional[pd.DataFrame]:
    """Lees CSV bestand in."""
    try:
        return pd.read_csv(file, dtype=str)
    except Exception:
        return None


def _read_excel_data(file: Path) -> Optional[pd.DataFrame]:
    """Lees Excel bestand in."""
    try:
        xls = pd.ExcelFile(file, engine="openpyxl")
        # Neem de eerste sheet als bron
        sheet_name = next(iter(xls.sheet_names), None)
        if sheet_name:
            return pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
        return None
    except Exception:
        return None


def _read_file_data(file: Path) -> Optional[pd.DataFrame]:
    """Lees data uit een bestand (CSV of Excel)."""
    if file.suffix.lower() == ".csv":
        return _read_csv_data(file)
    else:
        return _read_excel_data(file)


