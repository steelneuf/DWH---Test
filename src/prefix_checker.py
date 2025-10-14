import pandas as pd
from pathlib import Path
import logging
import sys
import os

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def get_excel_sheets(file_path: Path):
    """Haal alle sheetnamen op uit een Excel bestand."""
    try:
        xls = pd.ExcelFile(file_path, engine="openpyxl")
        return xls.sheet_names
    except Exception as exc:
        logger.warning(f"Kon Excel bestand '{file_path.name}' niet lezen: {exc}")
        return []

def check_prefixes_in_sheets(sheet_names):
    """Controleer of er prefixes zijn die eindigen met '_' in de sheetnamen."""
    prefixes_found = []
    
    for sheet_name in sheet_names:
        if '_' in sheet_name:
            # Zoek naar de laatste underscore
            last_underscore_pos = sheet_name.rfind('_')
            if last_underscore_pos > 0:  # Er moet iets voor de underscore staan
                prefix = sheet_name[:last_underscore_pos + 1]  # Inclusief de underscore
                new_name = sheet_name[last_underscore_pos + 1:]  # Na de underscore
                prefixes_found.append({
                    'original': sheet_name,
                    'prefix': prefix,
                    'new_name': new_name
                })
    
    return prefixes_found

def rename_sheet_in_excel(file_path: Path, old_name: str, new_name: str):
    """Hernoem een sheet in een Excel bestand."""
    try:
        # Lees alle sheets
        xls = pd.ExcelFile(file_path, engine="openpyxl")
        all_sheets = {}
        
        # Lees alle sheets
        for sheet in xls.sheet_names:
            if sheet == old_name:
                all_sheets[new_name] = pd.read_excel(xls, sheet_name=sheet)
            else:
                all_sheets[sheet] = pd.read_excel(xls, sheet_name=sheet)
        
        # Schrijf terug naar het bestand
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return True
    except Exception as exc:
        logger.error(f"Kon sheet niet hernoemen in '{file_path.name}': {exc}")
        return False

def check_and_rename_prefixes():
    """Hoofdfunctie om prefixes te controleren en eventueel te hernoemen."""
    # Definieer de mappen om te controleren
    input_dirs = [
        Path("Data/Input"),
        Path("Data/InputColumns")
    ]
    
    all_prefixes = []
    files_with_prefixes = {}
    
    # Controleer alle Excel bestanden
    for input_dir in input_dirs:
        if not input_dir.exists():
            logger.warning(f"Map '{input_dir}' bestaat niet")
            continue
            
        logger.info(f"Controleren van map: {input_dir}")
        
        for file_path in input_dir.rglob("*.xlsx"):
            logger.info(f"Controleren van bestand: {file_path}")
            sheet_names = get_excel_sheets(file_path)
            
            if sheet_names:
                prefixes = check_prefixes_in_sheets(sheet_names)
                if prefixes:
                    files_with_prefixes[str(file_path)] = prefixes
                    all_prefixes.extend(prefixes)
    
    # Toon resultaten
    if not all_prefixes:
        print("\n[OK] Geen prefixes gevonden die eindigen met '_' in de sheetnamen.")
        return True
    
    print(f"\n[INFO] Gevonden {len(all_prefixes)} prefixes in {len(files_with_prefixes)} bestanden:")
    print("=" * 60)
    
    for file_path, prefixes in files_with_prefixes.items():
        print(f"\n[FILE] Bestand: {file_path}")
        for prefix_info in prefixes:
            print(f"  [SHEET] Sheet: '{prefix_info['original']}' -> '{prefix_info['new_name']}'")
            print(f"      Prefix: '{prefix_info['prefix']}'")
    
    # Vraag gebruiker om bevestiging
    print("\n" + "=" * 60)
    response = input("\n[VRAAG] Wil je alle prefixes verwijderen? (j/n): ").lower().strip()
    
    if response in ['j', 'ja', 'y', 'yes']:
        print("\n[INFO] Prefixes worden verwijderd...")
        success_count = 0
        total_count = len(all_prefixes)
        
        for file_path, prefixes in files_with_prefixes.items():
            file_path_obj = Path(file_path)
            print(f"\n[FILE] Bezig met bestand: {file_path}")
            
            for prefix_info in prefixes:
                if rename_sheet_in_excel(file_path_obj, prefix_info['original'], prefix_info['new_name']):
                    print(f"  [OK] Sheet hernoemd: '{prefix_info['original']}' -> '{prefix_info['new_name']}'")
                    success_count += 1
                else:
                    print(f"  [ERROR] Fout bij hernoemen: '{prefix_info['original']}'")
        
        print(f"\n[RESULT] Resultaat: {success_count}/{total_count} sheets succesvol hernoemd")
        
        if success_count == total_count:
            print("[OK] Alle prefixes zijn succesvol verwijderd!")
            return True
        else:
            print("[WARNING] Sommige prefixes konden niet worden verwijderd.")
            return False
    else:
        print("[INFO] Geen wijzigingen doorgevoerd.")
        return False

if __name__ == "__main__":
    print("[INFO] Prefix Checker - Controleert sheetnamen op prefixes die eindigen met '_'")
    print("=" * 70)
    
    try:
        success = check_and_rename_prefixes()
        if success:
            print("\n[OK] Script succesvol voltooid!")
            sys.exit(0)
        else:
            print("\n[WARNING] Script voltooid met waarschuwingen.")
            sys.exit(1)
    except Exception as exc:
        logger.error(f"Onverwachte fout: {exc}")
        print(f"\n[ERROR] Fout opgetreden: {exc}")
        sys.exit(1)
