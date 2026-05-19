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

---

## Create local version of the database
 
- Create a new MS Access file
- Open VBA Editor (`Alt + F11`)
   - Import modules:
     - File → Import File... → `automated_setup.vba`
     - File → Import File... → `minimal_app_generator.vba`
     - File → Import File... → `minimal_app_runtime.vba`
     - File → Import File... → `UnlinkTables.bas`
  - Run
    1. `LinkAllTngTables` (in `automated_setup.vba`)
    2. `BuildAllForms` (in `minimal_app_generator.vba`)
    3. `RunMaterializationPipeline`
