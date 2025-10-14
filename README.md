# Data Validation Tool

Een Python tool voor het valideren van data uit verschillende bronnen (VR, VRR, DWH, OR) met een focus op code kwaliteit, onderhoudbaarheid en uitgebreide rapportage.

## 🚀 Features

- **Modulaire architectuur**: Kleine, herbruikbare functies die makkelijk te testen en onderhouden zijn
- **Verbeterde foutafhandeling**: Robuuste error handling met duidelijke logging
- **Type hints**: Volledige type annotaties voor betere code kwaliteit
- **Consistente styling**: PEP 8 compliant code met duidelijke docstrings
- **Excel output**: Professionele Excel rapporten met styling en samenvattingen
- **Uitgebreide validatie**: Vergelijking van data tussen meerdere bronnen
- **Missmatches detectie**: Automatische identificatie van rijen die niet matchen
- **Dubbele records detectie**: Vindt en rapporteert dubbele sleutelwaarden
- **Dashboard rapportage**: Overzichtelijke statistieken per bron en sheet

## 📁 Project Structuur

```
DWH - Test/
├── Data/
│   ├── Input/                 # Input bestanden (VR.xlsx, VRR.xlsx, DWH-Staging.xlsx)
│   ├── InputColumns/          # OR bestanden (hoofdzaak.xlsx, oordeel.xlsx)
│   ├── Output/                # Output bestanden
│   └── Validation/            # Configuratie (Kolommen.xlsx)
├── src/
│   ├── main.py               # Hoofdlogica en CLI interface
│   ├── compare.py            # Data vergelijking logica
│   ├── config.py             # Configuratie management
│   ├── logging_setup.py      # Logging configuratie
│   ├── utils.py              # Hulpfuncties
│   └── InputOutput/
│       ├── readers.py        # Excel bestand inlezen
│       ├── writers.py        # Excel output schrijven
│       ├── styling.py        # Excel styling toepassen
│       └── combiner.py       # Bestanden combineren
├── run.py                    # Eenvoudige uitvoering
└── requirements.txt          # Python dependencies
```

## 🛠️ Installatie

1. Clone de repository
2. Installeer dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## 📖 Gebruik

### Command Line Interface

```bash
# Standaard uitvoering
python run.py

# Of via module
python -m src.main

# Met custom paden
python -m src.main --input Data/CustomInput --output Data/CustomOutput --validation Data/CustomValidation
```

### Programmatisch gebruik

```python
from src.main import run
from pathlib import Path

run(
    input_dir=Path("Data/Input"),
    output_dir=Path("Data/Output"), 
    validation_dir=Path("Data/Validation")
)
```

## ⚙️ Configuratie

De tool gebruikt een `Kolommen.xlsx` bestand in de `Data/Validation` map om te bepalen welke kolommen vergeleken moeten worden.

### Kolommen.xlsx structuur

- **Sheet naam**: Naam van de sheet in de Excel bestanden
- **key_column**: Kolom die gebruikt wordt als unieke sleutel voor vergelijking
- **columns**: Lijst van kolommen die vergeleken moeten worden

### Voorbeeld configuratie:
```
Sheet       | key_column | columns
hoofdzaak   | ID         | [Naam, Status, Datum]
oordeel     | OordeelID  | [Oordeel, Datum, Status]
```

## 📊 Output Bestanden

### 1. ValidatieOutput.xlsx
Bevat de ruwe vergelijkingsresultaten per sheet met alle kolommen en waarden uit elke bron.

### 2. Testresultaat.xlsx
Bevat samenvattende rapporten en analyses in de volgende volgorde:

1. **Overview** - Per-sheet per-bron statistieken
2. **Details** - Overzicht per sheet met matches en mismatches  
3. **Matches** - Per-kolom matching statistieken
4. **Missmatches** - Alle rijen die niet perfect matchen
5. **Dubbele records** - Overzicht van dubbele sleutelwaarden
6. **Logs** - Gedetailleerde logging van het validatieproces

#### 📋 Details
- Overzicht per sheet met totaal aantal rijen, matches en mismatches
- Percentages van matches per kolom (zonder absolute aantallen)
- Aanwezigheid van keys in elke bron

#### 🎯 Overview
- Per-sheet per-bron statistieken
- Aantal rijen en kolommen per bron
- Key kolom analyse (uniek, duplicaten, null waarden)

#### 🔍 Matches
- Per-kolom matching statistieken
- Aanwezigheid per bron
- Match percentages en mismatch percentages

#### ❌ Missmatches
- **Alle rijen die niet perfect matchen** tussen bronnen
- **Bron informatie**: Welke bestanden bevatten de rij wel/niet
- **Mismatch details**: Welke kolommen hebben verschillende waarden
- **Type categorisering**: Alleen in één bron, kolom mismatch, etc.
- **Kolomwaarden**: De daadwerkelijke waarden uit elke bron

#### 🔄 Dubbele Records
- Overzicht van dubbele sleutelwaarden per bron en sheet
- Aantal duplicaten per key

#### 📝 Logs
- Gedetailleerde logging van het validatieproces
- Tijdsstempels en log niveaus

## 🔧 Code Kwaliteit Verbeteringen

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

## 📈 Nieuwe Functionaliteit

### Missmatches Detectie
De tool identificeert automatisch alle rijen die niet perfect matchen tussen bronnen:

```python
# Voorbeeld mismatch record
{
    "Sheet": "hoofdzaak",
    "Key": "12345",
    "Aanwezig_In": "VR, VRR",
    "Afwezig_In": "DWH-Staging",
    "Mismatched_Kolommen": "Naam, Status",
    "Type_Mismatch": "Deels aanwezig + kolom mismatch",
    "VR_Naam": "Oude Naam",
    "VRR_Naam": "Nieuwe Naam"
}
```

### Verbeterde Samenvatting
- Alleen percentages voor kolom matches (geen absolute aantallen)
- Schonere, overzichtelijkere weergave
- Focus op relevante informatie

## 🧪 Development

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

## 🚨 Troubleshooting

### Veelvoorkomende problemen

1. **Import errors**: Zorg dat je in de juiste directory bent en dependencies geïnstalleerd zijn
2. **Excel bestanden**: Controleer of bestanden niet open staan in Excel
3. **Permissions**: Zorg dat je schrijfrechten hebt in de output directory
4. **Configuratie**: Controleer of `Kolommen.xlsx` correct is geconfigureerd

### Debug Tips

- Controleer de logs in de "Logs" tab van Testresultaat.xlsx
- Gebruik `python -m py_compile src/main.py` om syntax fouten te vinden
- Controleer of alle input bestanden de verwachte sheets bevatten

## 📝 Changelog

### v1.1.0 (Huidige versie)
- ✅ Missmatches tab toegevoegd
- ✅ Verbeterde samenvatting (alleen percentages)
- ✅ Uitgebreide mismatch analyse
- ✅ Betere error handling
- ✅ Schonere code structuur

### v1.0.0
- Basis validatie functionaliteit
- Excel output met styling
- Dubbele records detectie

## 📄 Licentie

Intern gebruik - Data Validation Team

---

**💡 Tip**: Gebruik de Missmatches tab om snel inzicht te krijgen in welke data niet consistent is tussen je verschillende bronnen!

