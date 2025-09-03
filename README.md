# Data Validation Tool

Een Python tool voor het valideren van data uit verschillende bronnen (VR, VRR, DWH) met een focus op code kwaliteit en onderhoudbaarheid.

## Features

- **Modulaire architectuur**: Kleine, herbruikbare functies die makkelijk te testen en onderhouden zijn
- **Verbeterde foutafhandeling**: Robuuste error handling met duidelijke logging
- **Type hints**: Volledige type annotaties voor betere code kwaliteit
- **Consistente styling**: PEP 8 compliant code met duidelijke docstrings
- **Excel output**: Professionele Excel rapporten met styling en samenvattingen

## Project Structuur

```
src/
├── __init__.py           # Package initialisatie
├── main.py              # Hoofdlogica en CLI interface
├── compare.py           # Data vergelijking logica
├── config.py            # Configuratie management
├── logging_setup.py     # Logging configuratie
└── utils.py             # Hulpfuncties
```

## Installatie

1. Clone de repository
2. Installeer dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Gebruik

### Command Line Interface

```bash
# Standaard uitvoering
python -m src.main

# Met custom paden
python -m src.main --input Data/CustomInput --output Data/CustomOutput --validation Data/CustomValidation
```

### Programmatisch gebruik

```python
from src import run
from pathlib import Path

run(
    input_dir=Path("Data/Input"),
    output_dir=Path("Data/Output"), 
    validation_dir=Path("Data/Validation")
)
```

## Configuratie

De tool gebruikt een `Kolommen.xlsx` bestand in de `Data/Validation` map om te bepalen welke kolommen vergeleken moeten worden.

### Kolommen.xlsx structuur

- **Rij 1**: Kolomnamen
- **Rij 2**: Markeringen (bijv. "key" voor sleutelkolom, "bron" voor bronkolom)

## Code Kwaliteit Verbeteringen

### 1. Modulaire Functies
- Grote functies opgesplitst in kleine, herbruikbare functies
- Elke functie heeft één verantwoordelijkheid
- Betere testbaarheid en onderhoudbaarheid

### 2. Type Annotaties
- Volledige type hints toegevoegd
- Betere IDE ondersteuning en code validatie
- Duidelijkere interfaces tussen functies

### 3. Error Handling
- Robuuste exception handling
- Duidelijke logging van fouten
- Graceful degradation bij problemen

### 4. Code Styling
- PEP 8 compliant formatting
- Consistente naamgeving conventies
- Duidelijke docstrings voor alle functies

### 5. Performance Optimalisaties
- Efficiënte pandas operaties
- Geoptimaliseerde Excel schrijfacties
- Betere memory management

## Voorbeelden van Verbeteringen

### Voor: Grote, complexe functie
```python
def process_single_sheet(logger, writer, sheet_name, cfg, source_map):
    # 50+ regels code met meerdere verantwoordelijkheden
    pass
```

### Na: Modulaire, herbruikbare functies
```python
def process_single_sheet(logger, writer, sheet_name, cfg, source_map):
    # Hoofdfunctie die andere functies aanroept
    source_to_df = _load_source_data(source_map, sheet_name)
    result_df = _compare_data(source_to_df, cfg)
    _write_output(writer, sheet_name, result_df)
    return _create_summary(result_df)

def _load_source_data(source_map, sheet_name):
    # Eén verantwoordelijkheid: data laden
    pass

def _compare_data(source_to_df, cfg):
    # Eén verantwoordelijkheid: data vergelijken
    pass
```

## Development

### Code Style
- Volg PEP 8 richtlijnen
- Gebruik type hints
- Schrijf docstrings voor alle functies
- Houd functies kort (< 20 regels)

### Testing
- Test elke functie individueel
- Gebruik unit tests voor kleine functies
- Integration tests voor grotere workflows

### Refactoring
- Splits grote functies op in kleinere
- Identificeer herbruikbare code
- Verbeter error handling
- Optimaliseer performance bottlenecks

## Troubleshooting

### Veelvoorkomende problemen

1. **Import errors**: Zorg dat je in de juiste directory bent en dependencies geïnstalleerd zijn
2. **Excel bestanden**: Controleer of bestanden niet open staan in Excel
3. **Permissions**: Zorg dat je schrijfrechten hebt in de output directory

## Licentie

Intern gebruik - Data Validation Team

