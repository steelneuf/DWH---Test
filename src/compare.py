from typing import List, Tuple, Dict, Union
import pandas as pd
try:
    # Prefer official NAType when available
    from pandas._libs.missing import NAType  # type: ignore
except Exception:  # pragma: no cover - compatibility fallback
    NAType = type(pd.NA)  # type: ignore
import re

# Compile regex pattern once for performance
_NUM_SEP_PATTERN = re.compile(r"^[\d][\d.,]*$")


def _is_numeric_with_separators(value: str) -> bool:
    """Controleer of tekst een getal met scheidingstekens is."""
    if not value:
        return False
    return bool(_NUM_SEP_PATTERN.match(value))


def normalize_key_value(value: object) -> Union[str, NAType]:
    """Normaliseer keys door spaties te verwijderen en numerieke scheidingstekens te strippen."""
    try:
        if pd.isna(value):
            return pd.NA
    except Exception:
        pass

    s = str(value).strip()
    if not s:
        return s

    s_no_spaces = s.replace(' ', '')
    if _is_numeric_with_separators(s_no_spaces):
        return s_no_spaces.translate(str.maketrans('', '', ',.'))

    return s_no_spaces


def _safe_normalize_key_column(df: pd.DataFrame, key_column: str) -> pd.DataFrame:
    """Normaliseer key kolom veilig met error handling."""
    try:
        df_copy = df.copy()
        df_copy[key_column] = df_copy[key_column].map(normalize_key_value).astype("string")
        return df_copy
    except Exception:
        return df


def _ensure_key_column_exists(df: pd.DataFrame, key_column: str) -> pd.DataFrame:
    """Zorg ervoor dat key kolom bestaat, voeg toe indien nodig."""
    if key_column not in df.columns:
        df_copy = df.copy()
        df_copy[key_column] = pd.NA
        return df_copy
    return df


def _get_columns_to_keep(key_column: str, columns_to_compare: List[str], available_columns: List[str]) -> List[str]:
    """Bepaal welke kolommen behouden moeten blijven uit de beschikbare kolommen.

    Zorgt er expliciet voor dat de key kolom behouden blijft en dat alleen kolommen
    die daadwerkelijk bestaan geselecteerd worden, om KeyErrors te voorkomen.
    """
    desired_columns = [key_column] + [c for c in columns_to_compare if c != key_column]
    return [c for c in desired_columns if c in available_columns]


def _project_single_source_data(
    df: pd.DataFrame, 
    columns_to_compare: List[str], 
    key_column: str
) -> pd.DataFrame:
    """Projecteer data van één bron naar consistente structuur."""
    if df.empty:
        return pd.DataFrame(columns=[key_column] + columns_to_compare)
    
    # Zorg ervoor dat key kolom bestaat
    df_with_key = _ensure_key_column_exists(df, key_column)
    
    # Bepaal te behouden kolommen
    keep_cols = _get_columns_to_keep(key_column, columns_to_compare, list(df_with_key.columns))
    projected = df_with_key[keep_cols].copy()
    
    # Voeg genormaliseerde 'Key' toe, behoud de originele sleutelkolom ongewijzigd
    try:
        projected["Key"] = projected[key_column].map(normalize_key_value).astype("string")
    except Exception:
        projected["Key"] = projected.get(key_column)
    
    return projected


def _project_source_data(
    source_to_df: Dict[str, pd.DataFrame], 
    columns_to_compare: List[str], 
    key_column: str
) -> Dict[str, pd.DataFrame]:
    """Projecteer data van elke bron naar een consistente structuur."""
    return {
        source: _project_single_source_data(df, columns_to_compare, key_column)
        for source, df in source_to_df.items()
    }


def _rename_source_columns(df: pd.DataFrame, source: str) -> pd.DataFrame:
    """Hernoem kolommen met bron-suffix (behalve Key)."""
    rename_map = {c: f"{c}_{source}" for c in df.columns if c != "Key"}
    return df.rename(columns=rename_map)


def _merge_projected_data(projected_by_source: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Merge alle geprojecteerde data op Key kolom."""
    if not projected_by_source:
        return pd.DataFrame(columns=["Key"])
    
    merged = None
    sources = list(projected_by_source.keys())
    
    for idx, (source, pdf) in enumerate(projected_by_source.items()):
        # Hernoem kolommen VAN DEZE BRON voordat we mergen, zodat we niet
        # bestaande (al gesuffixte) kolommen telkens opnieuw hernoemen.
        pdf_renamed = _rename_source_columns(pdf, source)
        if idx == 0:
            merged = pdf_renamed
        else:
            merged = merged.merge(pdf_renamed, on="Key", how="outer", suffixes=(None, None))
    
    return merged


def _create_presence_series(merged: pd.DataFrame, source_keys: set) -> pd.Series:
    """Maak een series die aangeeft of keys aanwezig zijn in een bron."""
    if merged.empty:
        return pd.Series([], dtype=bool)
    return merged["Key"].isin(source_keys)


def _add_presence_columns(
    result_df: pd.DataFrame, 
    merged: pd.DataFrame, 
    projected_by_source: Dict[str, pd.DataFrame]
) -> Tuple[pd.DataFrame, pd.Series]:
    """Voeg kolommen toe die aangeven of keys aanwezig zijn in elke bron."""
    all_present_series = pd.Series(True, index=merged.index) if not merged.empty else pd.Series([], dtype=bool)

    # Verzamel alle nieuwe kolommen eerst om gefragmenteerde toewijzingen te voorkomen
    new_cols: Dict[str, pd.Series] = {}
    for source, projected_df in projected_by_source.items():
        keys = set(projected_df["Key"].dropna())
        present_col = f"Aanwezig_{source}"
        present_series = _create_presence_series(merged, keys)
        new_cols[present_col] = present_series.map({True: "ja", False: "nee"})
        all_present_series = all_present_series & present_series

    # Voeg Match_Key kolom toe
    if not merged.empty:
        new_cols["Match_Key"] = all_present_series.map({True: "ja", False: "nee"})
    else:
        new_cols["Match_Key"] = pd.Series([], dtype="string")

    if new_cols:
        result_df = pd.concat([result_df, pd.DataFrame(new_cols, index=merged.index)], axis=1)
        # Defragmenteer
        result_df = result_df.copy()

    return result_df, all_present_series


def _add_source_columns(
    result_df: pd.DataFrame, 
    merged: pd.DataFrame, 
    columns_to_compare: List[str], 
    source_to_df: Dict[str, pd.DataFrame]
) -> pd.DataFrame:
    """Voeg kolommen toe met waarden uit elke bron."""
    new_cols: Dict[str, pd.Series] = {}

    # Voeg per-bron de originele sleutel (prefixed) toe
    for source in source_to_df.keys():
        # Vind de originele sleutelkolomnaam in merged voor deze bron
        src_key_cols = [c for c in merged.columns if c.endswith(f"_{source}") and c != "Key"]
        src_key_cols.sort()
        if src_key_cols:
            new_cols[f"{source}_Key"] = merged[src_key_cols[0]]
        else:
            new_cols[f"{source}_Key"] = pd.Series(pd.NA, index=merged.index)

    # Voeg overige kolommen toe
    for col in columns_to_compare:
        for source in source_to_df.keys():
            src_col = f"{col}_{source}"
            if src_col in merged.columns:
                new_cols[f"{source}_{col}"] = merged[src_col]
            else:
                new_cols[f"{source}_{col}"] = pd.Series(pd.NA, index=merged.index)

    if new_cols:
        result_df = pd.concat([result_df, pd.DataFrame(new_cols, index=merged.index)], axis=1)
        # Defragmenteer
        result_df = result_df.copy()

    return result_df


def _compare_column_values(col_values: List[pd.Series], merged: pd.DataFrame) -> Tuple[pd.Series, pd.Series]:
    """Vergelijk waarden van één kolom tussen bronnen."""
    if not col_values or len(col_values) < 2:
        return pd.Series([], dtype=bool), pd.Series([], dtype=bool)
    
    ref = col_values[0]
    if merged.empty:
        return pd.Series([], dtype=bool), pd.Series([], dtype=bool)
    
    # Vergelijk tegen eerste kolom
    equal_series = pd.Series(True, index=merged.index)
    for s in col_values[1:]:
        equal_series = equal_series & (ref == s)
    
    # Check of alle waarden NA zijn
    all_na_series = pd.Series(True, index=merged.index)
    for s in col_values:
        all_na_series = all_na_series & s.isna()
    
    return equal_series, all_na_series


def _add_match_columns(
    result_df: pd.DataFrame, 
    merged: pd.DataFrame, 
    columns_to_compare: List[str], 
    source_to_df: Dict[str, pd.DataFrame]
) -> Tuple[pd.DataFrame, pd.Series]:
    """Voeg kolommen toe die aangeven of waarden matchen tussen bronnen."""
    if merged.empty:
        all_cols_match = pd.Series([], dtype=bool)
    else:
        all_cols_match = pd.Series(True, index=merged.index)

    new_cols: Dict[str, pd.Series] = {}

    for col in columns_to_compare:
        col_values: List[pd.Series] = []
        for source in source_to_df.keys():
            src_col = f"{col}_{source}"
            if src_col in merged.columns:
                col_values.append(merged[src_col])
            else:
                col_values.append(pd.Series(pd.NA, index=merged.index))

        equal_series, all_na_series = _compare_column_values(col_values, merged)

        if not merged.empty:
            match_series = (equal_series | all_na_series).map({True: "ja", False: "nee"})
            new_cols[f"Match_{col}"] = match_series
            all_cols_match = all_cols_match & (equal_series | all_na_series)

    if new_cols:
        result_df = pd.concat([result_df, pd.DataFrame(new_cols, index=merged.index)], axis=1)
        # Defragmenteer
        result_df = result_df.copy()

    return result_df, all_cols_match


def _get_ordered_columns(
    result_df: pd.DataFrame, 
    columns_to_compare: List[str], 
    source_to_df: Dict[str, pd.DataFrame]
) -> List[str]:
    """Bepaal logische volgorde van kolommen."""
    ordered_cols = ["Key"]
    
    # Aanwezigheid kolommen
    ordered_cols.extend([c for c in result_df.columns if c.startswith("Aanwezig_")])
    
    # Meta kolommen (BronMatch blijft intern, maar we houden hem in het resultaat)
    ordered_cols.extend(["Match_Key", "BronMatch"]) 
    
    # Bron sleutel kolommen eerst, dan bron kolommen en match kolommen
    for source in source_to_df.keys():
        ordered_cols.append(f"{source}_Key")
    
    # Bron kolommen en match kolommen
    for col in columns_to_compare:
        for source in source_to_df.keys():
            ordered_cols.append(f"{source}_{col}")
        ordered_cols.append(f"Match_{col}")
    
    return ordered_cols


def _order_columns(
    result_df: pd.DataFrame, 
    columns_to_compare: List[str], 
    source_to_df: Dict[str, pd.DataFrame]
) -> pd.DataFrame:
    """Order kolommen in logische volgorde."""
    ordered_cols = _get_ordered_columns(result_df, columns_to_compare, source_to_df)
    return result_df[ordered_cols]


def compare_sources(
    source_to_df: Dict[str, pd.DataFrame],
    columns_to_compare: List[str],
    key_column: str,
) -> Tuple[pd.DataFrame, int, int]:
    """Vergelijk bronnen op key en kolommen en geef resultaat en tellingen terug."""
    # Filter kolommen (exclude key column)
    columns_to_compare = [c for c in columns_to_compare if str(c).strip() != str(key_column).strip()]
    
    # Projecteer data van elke bron
    projected_by_source = _project_source_data(source_to_df, columns_to_compare, key_column)
    
    # Merge alle data
    merged = _merge_projected_data(projected_by_source)
    
    # Initialiseer resultaat DataFrame
    result_df = pd.DataFrame()
    result_df["Key"] = merged.get("Key", pd.Series(dtype="string"))
    
    # Voeg aanwezigheid kolommen toe
    result_df, all_present_series = _add_presence_columns(result_df, merged, projected_by_source)
    
    # Voeg bron kolommen toe
    result_df = _add_source_columns(result_df, merged, columns_to_compare, source_to_df)
    
    # Voeg match kolommen toe
    result_df, all_cols_match = _add_match_columns(result_df, merged, columns_to_compare, source_to_df)
    
    # Bepaal totale match status (intern), maar schrijf 'BronMatch' niet uit
    if not merged.empty:
        row_full_match = all_present_series & all_cols_match
        result_df["BronMatch"] = row_full_match.map({True: "ja", False: "nee"})
        matches = int(row_full_match.sum())
    else:
        result_df["BronMatch"] = pd.Series([], dtype="string")
        matches = 0
    
    # Order kolommen
    result_df = _order_columns(result_df, columns_to_compare, source_to_df)
    
    # Tel mismatches
    mismatches = int(len(result_df) - matches)
    
    return result_df, matches, mismatches
