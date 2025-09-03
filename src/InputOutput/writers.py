from pathlib import Path
from typing import List, Dict
import pandas as pd


def _create_xlsxwriter_writer(out_path: Path) -> pd.ExcelWriter:
    """Maak Excel writer met xlsxwriter engine."""
    return pd.ExcelWriter(
        out_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    )


def _create_openpyxl_writer(out_path: Path) -> pd.ExcelWriter:
    """Maak Excel writer met openpyxl engine als fallback."""
    return pd.ExcelWriter(out_path, engine="openpyxl")


def create_excel_writer(out_path: Path) -> pd.ExcelWriter:
    """Maak een Excel-writer met veilige defaults."""
    try:
        return _create_xlsxwriter_writer(out_path)
    except Exception:
        return _create_openpyxl_writer(out_path)


def _create_info_dataframe(message: str) -> pd.DataFrame:
    """Maak DataFrame met informatiebericht."""
    return pd.DataFrame({"Info": [message]})


def write_dataframe_or_info(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    """Schrijf DataFrame of informatieregel bij ontbrekende data."""
    if df.empty:
        info_df = _create_info_dataframe("Geen data gevonden in een of meer bronnen of configuratie leeg.")
        info_df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # Verberg 'BronMatch' kolom in output indien aanwezig
        try:
            out_df = df.drop(columns=["BronMatch"], errors="ignore")
        except Exception:
            out_df = df
        out_df.to_excel(writer, sheet_name=sheet_name, index=False)


def write_duplicate_sheet(writer: pd.ExcelWriter, duplicate_records: List[Dict[str, object]]) -> None:
    """Schrijf overzicht van dubbele sleutels per bron/sheet."""
    dup_sheet_name = "dubbele records"
    
    if duplicate_records:
        dup_df = pd.DataFrame(duplicate_records)
        dup_df = dup_df.sort_values(["Sheet", "Bron", "Key"], kind="stable")
        dup_df.to_excel(writer, sheet_name=dup_sheet_name, index=False)
    else:
        info_df = _create_info_dataframe("Geen dubbele sleutels gevonden in de aangeleverde bronnen.")
        info_df.to_excel(writer, sheet_name=dup_sheet_name, index=False)


def _create_summary_dataframe(sheet_summaries: List[Dict[str, object]]) -> pd.DataFrame:
    """Maak DataFrame van sheet samenvattingen."""
    if sheet_summaries:
        return pd.DataFrame(sheet_summaries)
    else:
        return pd.DataFrame({
            "Sheet": ["<geen>"],
            "Totaal": [0],
            "Matches": [0],
            "Mismatches": [0],
        })


def write_summary_sheet(writer: pd.ExcelWriter, sheet_summaries: List[Dict[str, object]]) -> None:
    """Schrijf de samenvatting per sheet."""
    summary_df = _create_summary_dataframe(sheet_summaries)
    summary_df.to_excel(writer, sheet_name="Samenvatting", index=False)


def write_logs_sheet(writer: pd.ExcelWriter, in_memory_logs: List[Dict[str, str]]) -> None:
    """Schrijf de gelogde berichten naar een tabblad."""
    try:
        if in_memory_logs:
            df = pd.DataFrame(in_memory_logs)
        else:
            df = pd.DataFrame(columns=["Tijd", "Niveau", "Bericht"])
        df.to_excel(writer, sheet_name="Logs", index=False)
    except ValueError as exc:
        # Vaak: invalid sheetname of writer gesloten
        import logging
        logging.getLogger("validation").warning("Kon logs niet schrijven: %s", exc)
    except Exception as exc:
        import logging
        logging.getLogger("validation").warning("Onbekende fout bij schrijven van logs: %s", exc)



def _create_dashboard_dataframe(dashboard_records: List[Dict[str, object]]) -> pd.DataFrame:
    """Maak DataFrame voor Dashboard statistieken."""
    if dashboard_records:
        df = pd.DataFrame(dashboard_records)
        # Zorg voor consistente kolomvolgorde wanneer aanwezig
        preferred_order = [
            "Sheet",
            "Bron",
            "Rijen",
            "Kolommen",
            "KeyKolom",
            "Key_NonNull",
            "Key_Null",
            "Key_Uniek",
            "Key_Duplicaten",
        ]
        cols = [c for c in preferred_order if c in df.columns]
        other_cols = [c for c in df.columns if c not in cols]
        return df[cols + other_cols]
    else:
        return pd.DataFrame({
            "Sheet": ["<geen>"],
            "Bron": ["<geen>"],
            "Rijen": [0],
            "Kolommen": [0],
        })


def write_dashboard_sheet(writer: pd.ExcelWriter, dashboard_records: List[Dict[str, object]]) -> None:
    """Schrijf Dashboard met per-sheet per-bron statistieken."""
    dash_df = _create_dashboard_dataframe(dashboard_records)
    dash_df.to_excel(writer, sheet_name="Dashboard", index=False)


def write_detailed_analysis_sheet(writer: pd.ExcelWriter, detailed_analysis_records: List[Dict[str, object]]) -> None:
    """Schrijf gedetailleerde kolom analyse naar een aparte sheet."""
    if detailed_analysis_records:
        analysis_df = pd.DataFrame(detailed_analysis_records)
        # Sorteer op Sheet, Type, en Kolom voor betere leesbaarheid
        analysis_df = analysis_df.sort_values(["Sheet", "Type", "Kolom"], kind="stable")
        analysis_df.to_excel(writer, sheet_name="Gedetailleerde Analyse", index=False)
    else:
        info_df = _create_info_dataframe("Geen gedetailleerde analyse beschikbaar.")
        info_df.to_excel(writer, sheet_name="Gedetailleerde Analyse", index=False)


def write_mismatches_sheet(writer: pd.ExcelWriter, all_mismatches: List[Dict[str, object]]) -> None:
    """Schrijf alle missmatches naar een aparte sheet."""
    if all_mismatches:
        mismatches_df = pd.DataFrame(all_mismatches)
        # Sorteer op Sheet, Key voor betere leesbaarheid
        mismatches_df = mismatches_df.sort_values(["Sheet", "Key"], kind="stable")
        mismatches_df.to_excel(writer, sheet_name="Missmatches", index=False)
    else:
        info_df = _create_info_dataframe("Geen missmatches gevonden - alle rijen matchen perfect tussen de bronnen.")
        info_df.to_excel(writer, sheet_name="Missmatches", index=False)
