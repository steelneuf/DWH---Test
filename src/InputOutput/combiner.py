from pathlib import Path
from typing import List
import pandas as pd

from .readers import _read_file_data
from .writers import create_excel_writer, _create_info_dataframe


def _get_target_sheet_name(file: Path) -> str:
    """Bepaal werkbladnaam op basis van bestandsnaam."""
    return file.stem[:31]  # Excel sheetnaam max 31 chars


def _get_input_files(folder: Path, out_path: Path) -> List[Path]:
    """Verzamel alle input bestanden."""
    excel_files = [p for p in sorted(folder.glob("*.xlsx")) if p.name.lower() != out_path.name.lower()]
    csv_files = [p for p in sorted(folder.glob("*.csv"))]
    return excel_files + csv_files


def _write_file_to_sheet(writer, file: Path, target_sheet: str) -> None:
    """Schrijf één bestand naar een sheet."""
    df = _read_file_data(file)
    
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        info_df = _create_info_dataframe(f"Geen data of leesfout voor {file.name}")
        info_df.to_excel(writer, sheet_name=target_sheet, index=False)
    else:
        df.to_excel(writer, sheet_name=target_sheet, index=False)


def combine_folder_to_workbook(folder: Path, out_path: Path) -> Path:
    """Combineer alle .xlsx-bestanden in een map tot één werkboek.
    
    Elke inputfile wordt een apart werkblad. De werkbladnaam is de bestandsnaam zonder extensie.
    Alleen de eerste sheet uit een bestand wordt gebruikt indien er meerdere aanwezig zijn.
    """
    folder.mkdir(parents=True, exist_ok=True)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Verzamel alle input bestanden
    input_files = _get_input_files(folder, out_path)

    with create_excel_writer(out_path) as writer:
        if not input_files:
            # Leeg werkboek met info-tab
            info_df = _create_info_dataframe("Geen inputbestanden gevonden in map")
            info_df.to_excel(writer, sheet_name="Info", index=False)
        else:
            for file in input_files:
                target_sheet = _get_target_sheet_name(file)
                _write_file_to_sheet(writer, file, target_sheet)

    return out_path


