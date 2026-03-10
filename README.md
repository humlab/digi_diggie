# DigiDiggie

Digital edition platform for historical court records and placename data.

## Overview

DigiDiggie is a database and tooling project for managing and analyzing historical Swedish court records (digitala domböcker) with integrated geographic placename data. The project includes:

- **PostgreSQL Database**: Core schema `digidiggie_tng` with PostGIS support
- **MS Access Application**: User-friendly data entry and editing interface
- **Python Tools**: Data export utilities for PostgreSQL and Access databases
- **VBA Code Generator**: Automated MS Access form creation

## Quick Links

### Documentation

- **[MS Access App Setup Guide](docs/ACCESS-APP-SETUP.md)** - Complete walkthrough to create the Access app from scratch
- **[PostgreSQL Linked Tables Guide](docs/LINKED-DATABASE.md)** - How to link Access to PostgreSQL via ODBC
- **[UCanAccess Setup](docs/UCANACCESS.md)** - Java-based Access database connectivity

### Key Files

- **Database Schema**: [`src/schema/digidiggie_en_tng.sql`](src/schema/digidiggie_en_tng.sql)
- **VBA Form Generator**: [`src/vba_tng_app/generate_tng_app.vba`](src/vba_tng_app/generate_tng_app.vba)
- **PostgreSQL Export Tool**: [`src/pgdb_to_excel.py`](src/pgdb_to_excel.py)
- **Access Export Tool**: [`src/accdb_to_excel.py`](src/accdb_to_excel.py)

## Getting Started

### Prerequisites

- Python 3.14+ with `uv` package manager
- PostgreSQL 12+ with PostGIS extension
- Microsoft Access 2016+ (for Access app)
- PostgreSQL ODBC driver (for Access/PostgreSQL linking)

### Python Environment Setup

```bash
# Clone repository
git clone https://github.com/humlab/digi_diggie.git
cd digi_diggie

# Install dependencies
uv sync

# Activate virtual environment
source .venv/bin/activate  # Linux/Mac
# or
.venv\Scripts\activate  # Windows
```

### Database Setup

1. Create PostgreSQL database:
   ```bash
   createdb -U postgres digidiggie
   ```

2. Load schema:
   ```bash
   psql -U postgres -d digidiggie -f src/schema/digidiggie_en_tng.sql
   ```

3. (Optional) Import placename data - see [Placename Database](#download-the-placename-database) section below

### MS Access App Setup

Follow the comprehensive guide: **[ACCESS-APP-SETUP.md](docs/ACCESS-APP-SETUP.md)**

Quick steps:
1. Install PostgreSQL ODBC driver
2. Create System DSN pointing to PostgreSQL
3. Create blank Access database
4. Link PostgreSQL tables via External Data → ODBC
5. Import VBA module `generate_tng_app.vba`
6. Run `BuildForms_DigiDiggie_TNG()` to generate all forms

## Utilities

### Export PostgreSQL to Excel

```bash
# Export all tables
uv run src/pgdb_to_excel.py --host localhost --port 5433 --database digidiggie --schema digidiggie_tng --output data.xlsx

# Or using Makefile
make dump-to-excel
```

### Export Access Database to Excel

```bash
uv run src/accdb_to_excel.py --accdb digidiggie_dev.accdb --output access_export.xlsx
```

## Project Structure

```
digi_diggie/
├── docs/                          # Documentation
│   ├── ACCESS-APP-SETUP.md       # MS Access app setup guide
│   ├── LINKED-DATABASE.md        # PostgreSQL ODBC linking guide
│   └── UCANACCESS.md            # Java Access connectivity
├── src/
│   ├── schema/                   # PostgreSQL schema definitions
│   │   └── digidiggie_en_tng.sql
│   ├── vba_tng_app/              # VBA form generator
│   │   ├── generate_tng_app.vba  # Main generator script
│   │   └── improvements.md       # Feature tracking
│   ├── pgdb_to_excel.py          # PostgreSQL export utility
│   └── accdb_to_excel.py         # Access export utility
├── data/                         # Data files and CSVs
├── resources/                    # SQL resources
└── scripts/                      # Shell scripts
```

## Download the placename database

### Python scripts

### Prerequistes

 - Account and credentials to https://geotorget.lantmateriet.se/
 - Authority to download [Ortnamn, vector](https://geotorget.lantmateriet.se/geodataprodukter/ortnamn-nedladdning-vektor-api)
  
### Steps
 - Log in to https://geotorget.lantmateriet.se/
 - Open 
 - 


---

## 🇸🇪 **Lantmäteriets koordinatsystem**

Lantmäteriets geografiska data (t.ex. Ortnamn, GSD, Topografisk webbkarta, m.m.) levereras **nästan alltid i SWEREF 99 TM**, som har:

| Egenskap          | Värde                                  |
| ----------------- | -------------------------------------- |
| **System**        | SWEREF 99 TM                           |
| **EPSG / SRID**   | **3006**                               |
| **Enhet**         | meter                                  |
| **Koordinater**   | x ≈ 250000–850000, y ≈ 6100000–7700000 |
| **Arealreferens** | Plan (projektion)                      |
| **Typ**           | Projekterat koordinatsystem            |

Detta är **det svenska nationella referenssystemet** för kartdata.

---

## 🌍 **EPSG 4326 (WGS 84)**

| Egenskap        | Värde                                  |
| --------------- | -------------------------------------- |
| **System**      | WGS 84 (lat/long)                      |
| **SRID**        | 4326                                   |
| **Enhet**       | grader (decimal degrees)               |
| **Koordinater** | lat ≈ 55–69, lon ≈ 11–24               |
| **Typ**         | Geografiskt koordinatsystem (sfäriskt) |

Detta används främst för GPS, webbkartor (t.ex. Leaflet, OpenStreetMap, Google Maps), och datautbyte.

---

## 🔍 **Hur du ser vilket system din data använder**

Titta på dina koordinater (`nkoordinat`, `ekoordinat`):

| Kolumn       | Exempelvärde | Tolkning                                              |
| ------------ | ------------ | ----------------------------------------------------- |
| `nkoordinat` | 6588464      | → tydligt **Y i meter** (≈ 6,588 km norr om ekvatorn) |
| `ekoordinat` | 583500       | → **X i meter** (≈ 583 km östligt)                    |

Eftersom de är **sex–sju siffror långa** och i **meter**, är detta **SWEREF 99 TM (EPSG 3006)**, inte grader.
Om det vore 4326, skulle värdena ligga runt `lat ~ 59.3`, `lon ~ 17.0`.

---

## ✅ **Slutsats för Ortnamn**

För Lantmäteriets *Ortnamn*:

> Använd alltid **SRID 3006** (SWEREF 99 TM).

---

## 💡 **Exempel på korrekt PostGIS-hantering**

Efter att du importerat CSV utan `geom`-kolumnen:

```sql
ALTER TABLE placenames
  ADD COLUMN geom geometry(Point, 3006);

UPDATE placenames
  SET geom = ST_SetSRID(ST_MakePoint(ekoordinat, nkoordinat), 3006);
```

Vill du använda dem i t.ex. webbkarta (WGS84/4326):

```sql
SELECT ST_AsGeoJSON(ST_Transform(geom, 4326)) FROM placenames;
```

---
