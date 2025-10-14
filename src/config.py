from pathlib import Path
from typing import Dict, List, Optional
import pandas as pd
import logging

# Basismappen
script_dir = Path(__file__).parent
data_input_dir = script_dir.parent / "Data" / "Input"
data_output_dir = script_dir.parent / "Data" / "Output"
data_validation_dir = script_dir.parent / "Data" / "Validation"
data_inputcolumns_dir = script_dir.parent / "Data" / "InputColumns"

# Constanten (geen onnodige, ongebruikte suffixen of headers)


def _is_bron(name: object) -> bool:
    """Beoordeel of een kolomnaam de bronmarkering is."""
    try:
        return str(name).strip().lower() == "bron"
    except Exception:
        return False


def _extract_columns(header_row: List[str]) -> List[str]:
    """Extraheer kolommen uit header rij exclusief bron kolommen."""
    return [
        col for col in header_row
        if col
        and str(col).strip()
        and str(col).strip().lower() != "nan"
        and not _is_bron(col)
    ]


def _find_key_column_by_marker(header_row: List[str], key_markers: pd.Series) -> Optional[str]:
    """Zoek key kolom op basis van key marker in tweede rij."""
    for idx, marker in enumerate(key_markers.tolist()):
        if marker and marker.strip().lower() == "key":
            if idx < len(header_row):
                candidate = header_row[idx]
                if candidate and candidate.strip():
                    return candidate
    return None


# NB: fallback logica was eerder aanwezig maar wordt niet gebruikt; verwijderd voor eenvoud


def _process_sheet_config(df: pd.DataFrame, sheet: str, xlsx_path: Path) -> Optional[Dict[str, object]]:
    """Verwerk configuratie van één sheet."""
    if df.empty:
        return None
    
    header_row = df.iloc[0].astype(str).tolist()
    key_markers = df.iloc[1].astype(str).str.lower() if len(df) > 1 else pd.Series([], dtype=str)
    
    # Extraheer kolommen
    columns = _extract_columns(header_row)
    if not columns:
        return None
    
    # Zoek key kolom
    key_column = _find_key_column_by_marker(header_row, key_markers)
    if key_column is None:
        logging.getLogger("validation").error(
            "Geen 'key' markering gevonden in sheet '%s' van %s. Definieer de sleutel uitsluitend in de Validation map (Kolommen.xlsx).",
            sheet, xlsx_path.name,
        )
        return None
    
    return {"columns": columns, "key_column": key_column}


def read_kolommen_config(xlsx_path: Path) -> Dict[str, Dict[str, object]]:
    """Lees Kolommen.xlsx in tot een per-sheet configuratie."""
    try:
        xls = pd.ExcelFile(xlsx_path, engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(f"Kan '{xlsx_path.name}' niet openen: {exc}")

    config: Dict[str, Dict[str, object]] = {}
    
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            sheet_config = _process_sheet_config(df, sheet, xlsx_path)
            if sheet_config:
                config[sheet] = sheet_config
        except Exception as exc:
            logging.getLogger("validation").warning(
                "Fout bij verwerken van sheet '%s' in %s: %s", sheet, xlsx_path.name, exc
            )
            continue

    if not config:
        raise RuntimeError("Geen geldige configuratie gevonden in Kolommen.xlsx.")

    return config


def find_single_validation_config(validation_dir: Path) -> Path:
    """Zoek precies één configuratiebestand in de validation map."""
    if not validation_dir.exists():
        raise RuntimeError(f"Validation map bestaat niet: {validation_dir}")
    
    candidates = list(validation_dir.glob("*.xlsx"))
    if len(candidates) != 1:
        raise RuntimeError(
            f"Gevonden {len(candidates)} configuratiebestanden in {validation_dir}; verwacht precies 1"
        )
    
    return candidates[0]
