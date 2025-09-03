import argparse
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd

from .config import data_input_dir, data_output_dir, data_validation_dir, data_inputcolumns_dir
from .config import read_kolommen_config, find_single_validation_config
from .logging_setup import setup_logging
from .InputOutput import (
    load_sheet,
    create_excel_writer,
    write_dataframe_or_info,
    apply_sheet_styling,
    write_duplicate_sheet,
    write_summary_sheet,
    write_dashboard_sheet,
    write_logs_sheet,
)
from .compare import compare_sources, normalize_key_value
from .utils import discover_input_sources


def _load_single_source_data(path: Path, sheet_name: str) -> pd.DataFrame:
    """Laad data uit één bron."""
    try:
        if path.suffix.lower() == ".csv":
            return load_sheet(path, "")  # sheet_name wordt genegeerd voor CSV
        else:
            return load_sheet(path, sheet_name)
    except Exception as exc:
        # Log error en retourneer lege DataFrame
        return pd.DataFrame()


def _load_source_data(source_map: Dict[str, Path], sheet_name: str) -> Dict[str, pd.DataFrame]:
    """Laad data uit alle bronnen."""
    return {
        label: _load_single_source_data(path, sheet_name)
        for label, path in source_map.items()
    }


def _normalize_key_series(key_series: pd.Series) -> pd.Series:
    """Normaliseer key series veilig."""
    try:
        return key_series.map(normalize_key_value).astype("string")
    except Exception:
        return key_series


def _find_duplicate_keys(
    df: pd.DataFrame, 
    key_col: str, 
    sheet: str, 
    source: str, 
    logger
) -> List[Dict[str, object]]:
    """Vind dubbele sleutelwaarden in een DataFrame."""
    if df.empty or key_col not in df.columns:
        return []
    
    key_series = df[key_col]
    key_series = _normalize_key_series(key_series)
    
    total_na = int(key_series.isna().sum())
    if total_na > 0:
        logger.info(
            "Sheet '%s' bron %s: %d lege key(s) niet meegeteld bij duplicaten",
            sheet, source, total_na
        )
    
    counts = key_series.value_counts(dropna=True)
    duplicates = counts[counts > 1]
    
    return [
        {
            "Sheet": sheet,
            "Bron": source,
            "Key": key_val,
            "Aantal": int(cnt),
        }
        for key_val, cnt in duplicates.items()
    ]


def _find_all_duplicates(
    source_to_df: Dict[str, pd.DataFrame], 
    key_column: str, 
    sheet_name: str, 
    logger
) -> List[Dict[str, object]]:
    """Vind duplicaten in alle bronnen."""
    duplicates = []
    for source_label, df in source_to_df.items():
        duplicates.extend(_find_duplicate_keys(df, key_column, sheet_name, source_label, logger))
    return duplicates


def _create_summary(sheet_name: str, result_df: pd.DataFrame) -> Dict[str, object]:
    """Maak samenvatting van resultaten."""
    return {
        "Sheet": sheet_name,
        "Totaal": len(result_df),
        "Matches": int((result_df["BronMatch"] == "ja").sum()),
        "Mismatches": int((result_df["BronMatch"] == "nee").sum()),
    }


def _safe_series(s: pd.Series) -> pd.Series:
    """Zorg voor een geldige Series voor tellingen."""
    try:
        return s
    except Exception:
        return pd.Series(dtype="object")


def _compute_key_stats(df: pd.DataFrame, key_col: str) -> Dict[str, int]:
    """Bereken key-statistieken veilig."""
    try:
        if key_col in df.columns:
            s = _safe_series(df[key_col])
            non_null = int(s.notna().sum())
            nulls = int(s.isna().sum())
            # Uniek en duplicaten op niet-NA
            s_nn = s.dropna()
            uniques = int(s_nn.nunique())
            dup_count = int((s_nn.duplicated(keep=False)).sum())
        else:
            non_null = 0
            nulls = int(len(df))
            uniques = 0
            dup_count = 0
        return {
            "Key_NonNull": non_null,
            "Key_Null": nulls,
            "Key_Uniek": uniques,
            "Key_Duplicaten": dup_count,
        }
    except Exception:
        return {
            "Key_NonNull": 0,
            "Key_Null": int(len(df)) if isinstance(df, pd.DataFrame) else 0,
            "Key_Uniek": 0,
            "Key_Duplicaten": 0,
        }


def _create_dashboard_records(
    sheet_name: str,
    key_column: str,
    source_to_df: Dict[str, pd.DataFrame],
) -> List[Dict[str, object]]:
    """Maak dashboard records per bron voor deze sheet."""
    records: List[Dict[str, object]] = []
    for source_label, df in source_to_df.items():
        try:
            rows = int(len(df))
            cols = int(len(df.columns)) if not df.empty else 0
        except Exception:
            rows, cols = 0, 0
        key_stats = _compute_key_stats(df, key_column)
        rec = {
            "Sheet": sheet_name,
            "Bron": source_label,
            "Rijen": rows,
            "Kolommen": cols,
            "KeyKolom": key_column,
        }
        rec.update(key_stats)
        records.append(rec)
    return records


def _process_sheet_data_only(
    writer,
    sheet_name: str,
    cfg: Dict[str, object],
    source_map: Dict[str, Path],
    logger
) -> None:
    """Verwerk één sheet alleen voor data output (zonder testresultaten)."""
    # Configuratie ophalen
    columns_to_compare, key_column = _process_sheet_configuration(cfg)
    
    # Data inlezen
    source_to_df = _load_source_data(source_map, sheet_name)
    
    # Vergelijken
    result_df, matches, mismatches = _compare_sheet_data(
        source_to_df, columns_to_compare, key_column
    )
    
    # Output schrijven naar data bestand
    _write_sheet_output(writer, sheet_name, result_df)


def _get_missing_sources(row: pd.Series, result_df: pd.DataFrame) -> List[str]:
    """Haal ontbrekende bronnen op uit een rij."""
    return [
        col.replace("Aanwezig_", "") 
        for col in result_df.columns 
        if col.startswith("Aanwezig_") and row.get(col, "nee") == "nee"
    ]


def _get_mismatched_columns(row: pd.Series, result_df: pd.DataFrame) -> List[str]:
    """Haal mismatched kolommen op uit een rij."""
    return [
        col.replace("Match_", "") 
        for col in result_df.columns 
        if col.startswith("Match_") and col != "Match_Key" and row.get(col, "ja") == "nee"
    ]


def _log_missing_sources(row: pd.Series, result_df: pd.DataFrame, key_val: str, logger) -> None:
    """Log ontbrekende bronnen."""
    missing_sources = _get_missing_sources(row, result_df)
    if missing_sources:
        logger.info(f"  - Key {key_val}: ontbreekt in {', '.join(missing_sources)}")


def _log_mismatched_columns(row: pd.Series, result_df: pd.DataFrame, key_val: str, logger) -> None:
    """Log mismatched kolommen."""
    mismatched_columns = _get_mismatched_columns(row, result_df)
    if mismatched_columns:
        logger.info(f"  - Key {key_val}: mismatch in kolommen: {', '.join(mismatched_columns)}")


def _log_mismatches(result_df: pd.DataFrame, logger) -> None:
    """Log alle mismatches."""
    for _, row in result_df.iterrows():
        if row.get("BronMatch", "nee") == "ja":
            continue
            
        key_val = str(row.get("Key", "<geen sleutel>"))
        _log_missing_sources(row, result_df, key_val, logger)
        _log_mismatched_columns(row, result_df, key_val, logger)


def _write_sheet_output(writer, sheet_name: str, result_df: pd.DataFrame) -> None:
    """Schrijf sheet output naar Excel."""
    write_dataframe_or_info(writer, sheet_name, result_df)
    apply_sheet_styling(writer, sheet_name, result_df)


def _process_sheet_configuration(cfg: Dict[str, object]) -> Tuple[List[str], str]:
    """Verwerk sheet configuratie en retourneer kolommen en key kolom."""
    columns_to_compare = [str(c).strip() for c in cfg["columns"] if str(c).strip()]
    key_column = str(cfg["key_column"]).strip()
    return columns_to_compare, key_column


def _compare_sheet_data(
    source_to_df: Dict[str, pd.DataFrame],
    columns_to_compare: List[str],
    key_column: str
) -> Tuple[pd.DataFrame, int, int]:
    """Vergelijk data van één sheet."""
    return compare_sources(
        source_to_df=source_to_df,
        columns_to_compare=columns_to_compare,
        key_column=key_column,
    )


def process_single_sheet(
    logger, 
    writer, 
    sheet_name: str, 
    cfg: Dict[str, object],
    source_map: Dict[str, Path]
) -> Tuple[Dict[str, object], List[Dict[str, object]], List[Dict[str, object]]]:
    """Verwerk één sheet: inlezen, vergelijken, loggen en wegschrijven."""
    # Configuratie ophalen
    columns_to_compare, key_column = _process_sheet_configuration(cfg)
    
    # Data inlezen
    source_to_df = _load_source_data(source_map, sheet_name)
    
    # Vergelijken
    result_df, matches, mismatches = _compare_sheet_data(
        source_to_df, columns_to_compare, key_column
    )
    
    # Duplicaten vinden
    duplicates = _find_all_duplicates(source_to_df, key_column, sheet_name, logger)
    # Dashboard records maken
    dashboard_records = _create_dashboard_records(sheet_name, key_column, source_to_df)
    
    # Output schrijven (alleen als writer beschikbaar is)
    if writer is not None:
        _write_sheet_output(writer, sheet_name, result_df)
    
    # Samenvatting maken en loggen
    summary = _create_summary(sheet_name, result_df)
    logger.info(f"Sheet '{sheet_name}': {summary['Matches']} matches, {summary['Mismatches']} mismatches")
    _log_mismatches(result_df, logger)
    
    return summary, duplicates, dashboard_records


def _handle_source_conflict(label: str, source_map: Dict[str, Path]) -> str:
    """Los labelconflict op door suffix toe te voegen."""
    if label in source_map:
        return f"{label}_InputColumns"
    return label


def _discover_alternative_sources(source_map: Dict[str, Path]) -> Dict[str, Path]:
    """Ontdek alternatieve bronnen uit InputColumns directory."""
    if not data_inputcolumns_dir.exists():
        return {}
    
    alt_sources = discover_input_sources(data_inputcolumns_dir)
    return {
        _handle_source_conflict(label, source_map): path
        for label, path in alt_sources.items()
    }


def _discover_all_sources(input_dir: Path) -> Dict[str, Path]:
    """Ontdek alle inputbronnen."""
    source_map = {}
    
    # Primaire bronnen
    primary_sources = discover_input_sources(input_dir)
    source_map.update(primary_sources)
    
    # Alternatieve bronnen
    alt_sources = _discover_alternative_sources(source_map)
    source_map.update(alt_sources)
    
    return source_map


def _load_validation_config(validation_dir: Path, logger) -> Dict[str, object]:
    """Laad validatie configuratie."""
    try:
        kolommen_file = find_single_validation_config(validation_dir)
        return read_kolommen_config(kolommen_file)
    except Exception as exc:
        logger.error("Probleem met configuratiebestand in Validation: %s", exc)
        return {}


def _create_error_summary(sheet_name: str) -> Dict[str, object]:
    """Maak error samenvatting voor sheet."""
    return {
        "Sheet": sheet_name,
        "Totaal": 0,
        "Matches": 0,
        "Mismatches": 0,
    }


def _process_sheet_with_error_handling(
    sheet_name: str,
    cfg: Dict[str, object],
    source_map: Dict[str, Path],
    logger
) -> Tuple[Dict[str, object], List[Dict[str, object]], List[Dict[str, object]]]:
    """Verwerk één sheet met error handling."""
    try:
        return process_single_sheet(logger, None, sheet_name, cfg, source_map)
    except Exception as exc:
        logger.error("Fout bij verwerken van sheet '%s': %s", sheet_name, exc)
        return _create_error_summary(sheet_name), [], []


def _process_all_sheets(
    config: Dict[str, object], 
    source_map: Dict[str, Path], 
    output_dir: Path, 
    logger, 
    in_memory_logs
) -> None:
    """Verwerk alle sheets en schrijf output naar twee bestanden."""
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Bestand 1: Data/Validatie resultaten
    data_output_file = output_dir / "ValidatieOutput.xlsx"
    
    # Bestand 2: Testresultaten
    test_output_file = output_dir / "Testresultaat.xlsx"
    
    # Verwerk alle sheets en verzamel resultaten
    sheet_summaries = []
    duplicate_records = []
    dashboard_records: List[Dict[str, object]] = []
    
    # Schrijf data sheets naar ValidatieOutput.xlsx
    with create_excel_writer(data_output_file) as writer:
        for sheet_name, cfg in config.items():
            summary, dups, dash_recs = _process_sheet_with_error_handling(
                sheet_name, cfg, source_map, logger
            )
            sheet_summaries.append(summary)
            duplicate_records.extend(dups)
            dashboard_records.extend(dash_recs)
            
            # Schrijf data sheet naar ValidatieOutput.xlsx
            _process_sheet_data_only(writer, sheet_name, cfg, source_map, logger)
    
    # Schrijf testresultaten naar Testresultaat.xlsx
    with create_excel_writer(test_output_file) as writer:
        write_duplicate_sheet(writer, duplicate_records)
        write_summary_sheet(writer, sheet_summaries)
        write_dashboard_sheet(writer, dashboard_records)
        write_logs_sheet(writer, in_memory_logs)
    
    logger.info(f"Data/Validatie-output geschreven naar: {data_output_file.name}")
    logger.info(f"Testresultaten geschreven naar: {test_output_file.name}")


def _log_discovered_sources(source_map: Dict[str, Path], logger) -> None:
    """Log ontdekte inputbronnen."""
    discovered = ", ".join([f"{label} -> {path.name}" for label, path in source_map.items()])
    logger.info("Gevonden inputbronnen (%d): %s", len(source_map), discovered)


def _validate_source_discovery(source_map: Dict[str, Path], logger) -> bool:
    """Valideer of er bronnen zijn ontdekt."""
    if not source_map:
        logger.error(
            "Geen bronbestanden (*.xlsx) gevonden in %s of %s",
            str(data_input_dir), str(data_inputcolumns_dir)
        )
        return False
    return True


def run(input_dir: Path, output_dir: Path, validation_dir: Path) -> None:
    """Orkestreer de validatie-run van inlezen tot rapportage."""
    logger, in_memory_logs = setup_logging()
    logger.info("Start validatie")
    
    # Bronnen ontdekken
    source_map = _discover_all_sources(input_dir)
    if not _validate_source_discovery(source_map, logger):
        return
    
    _log_discovered_sources(source_map, logger)
    
    # Configuratie laden
    config = _load_validation_config(validation_dir, logger)
    if not config:
        return
    
    # Verwerking uitvoeren
    _process_all_sheets(config, source_map, output_dir, logger, in_memory_logs)
    
    logger.info("Validatie voltooid")


def _parse_arguments() -> argparse.Namespace:
    """Parse command line argumenten."""
    parser = argparse.ArgumentParser(description="Valideer VR/VRR/DWH excelbestanden")
    parser.add_argument(
        "--input", dest="input_dir", type=Path, default=data_input_dir,
        help="Pad naar input map (default: Data/Input)"
    )
    parser.add_argument(
        "--output", dest="output_dir", type=Path, default=data_output_dir,
        help="Pad naar output map (default: Data/Output)"
    )
    parser.add_argument(
        "--validation", dest="validation_dir", type=Path, default=data_validation_dir,
        help="Pad naar validation map met Kolommen.xlsx (default: Data/Validation)"
    )
    return parser.parse_args()


def main() -> None:
    """CLI-entrypoint voor validatieproces."""
    args = _parse_arguments()
    run(args.input_dir, args.output_dir, args.validation_dir)


if __name__ == "__main__":
    main()
