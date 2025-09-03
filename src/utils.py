from pathlib import Path
from typing import List, Dict, Optional
import logging
from .InputOutput import combine_folder_to_workbook


def _get_logger() -> logging.Logger:
    """Haal de validation logger op."""
    return logging.getLogger("validation")


def _get_excel_files(input_dir: Path) -> List[Path]:
    """Verzamel alle Excel bestanden uit een directory."""
    return list(input_dir.glob("*.xlsx"))


def _get_csv_files(input_dir: Path) -> List[Path]:
    """Verzamel alle CSV bestanden uit een directory."""
    return list(input_dir.glob("*.csv"))


def _discover_single_files(input_dir: Path) -> Dict[str, Path]:
    """Vind losse Excel en CSV bestanden in input directory."""
    sources = {}
    
    # Zoek naar .xlsx en .csv bestanden
    excel_files = _get_excel_files(input_dir)
    csv_files = _get_csv_files(input_dir)
    
    for file_path in sorted(excel_files + csv_files):
        label = file_path.stem
        sources[label] = file_path
    
    return sources


def _combine_folder_to_workbook_safe(subdir: Path, combined_path: Path) -> Optional[Path]:
    """Combineer folder tot werkboek met error handling."""
    try:
        return combine_folder_to_workbook(subdir, combined_path)
    except Exception as exc:
        logger = _get_logger()
        logger.warning(
            "Kon map '%s' niet combineren: %s", subdir.name, exc
        )
        return None


def _discover_folder_bundles(input_dir: Path) -> Dict[str, Path]:
    """Vind submappen en combineer deze tot werkboeken."""
    sources = {}
    
    for subdir in sorted([d for d in input_dir.iterdir() if d.is_dir()]):
        combined_name = f"{subdir.name}_combined.xlsx"
        combined_path = subdir / combined_name
        
        combined_file = _combine_folder_to_workbook_safe(subdir, combined_path)
        if combined_file:
            label = subdir.name
            sources[label] = combined_file
    
    return sources


def discover_input_sources(input_dir: Path) -> Dict[str, Path]:
    """Vind invoerbronnen uit losse bestanden en samengestelde folderbundels.

    Ondersteunt twee vormen:
    1) Losse Excel-bestanden direct in `input_dir`.
    2) Submappen onder `input_dir` (zoals `InputColumns/<folder>`) die meerdere Excel-bestanden bevatten.
       Deze worden eerst samengevoegd tot één werkboek in dezelfde map met naam `<folder>_combined.xlsx`.
    """
    sources = {}
    
    # 1) Losse bestanden (.xlsx en .csv)
    single_files = _discover_single_files(input_dir)
    sources.update(single_files)
    
    # 2) Submappen met meerdere bestanden → combineer naar één xlsx
    folder_bundles = _discover_folder_bundles(input_dir)
    sources.update(folder_bundles)
    
    return sources


