# DigiDiggie AI Instructions

## Project Overview

DigiDiggie is a digital edition platform for historical Swedish court records (digitala domböcker) with integrated geographic placename data. The project bridges PostgreSQL database backend with MS Access frontend for data entry.

**Key Components:**
- PostgreSQL database with PostGIS (schema: `digidiggie_tng`)
- MS Access application for data entry/editing (linked via ODBC)
- Python utilities for data export and conversion
- VBA code generators for automated Access form creation
- Swedish placename (ortnamn) data integration

## Technology Stack

- **Database**: PostgreSQL 12+ with PostGIS extension
- **Frontend**: Microsoft Access 2016+ (Windows)
- **Backend Tools**: Python 3.14+, `uv` package manager
- **Key Libraries**: pandas, psycopg3, pyodbc, openpyxl, xlsxwriter, sqlalchemy, sqlglot
- **VBA**: Access form generation and automation
- **Infrastructure**: Docker support for PostgreSQL
- **Tools**: MDB tools for Access database operations

## Project Structure

- `src/schema/` - PostgreSQL DDL schemas (digidiggie_en_tng.sql is primary)
- `src/vba_tng_app/` - VBA code generator for Access forms
- `src/pgdb_to_excel.py` - Export PostgreSQL to Excel
- `src/accdb_to_excel.py` - Export Access database to Excel
- `docs/` - Comprehensive documentation (ACCESS-APP-SETUP.md, LINKED-DATABASE.md)
- `data/csv/` - CSV data files for import
- `data/ortnamn/` - Swedish placename data
- `scripts/` - Utility scripts (mdb-to-pg, dump-original)
- `docker/` - Docker Compose setup for PostgreSQL
- `resources/` - SQL scripts for data loading

## Database Schema

**Primary Schema**: `digidiggie_tng`

**Core Tables:**
- `court_case` - Main case records
- `court_case_entry` - Individual entries within cases
- `person` - Person records
- `person_entry` - Links persons to case entries
- `community`, `parish` - Geographic entities
- `source`, `ruling`, `role` - Lookup tables
- `placename` - Swedish placenames with PostGIS geometry

**Schema Conventions:**
- Table names: snake_case (e.g., `court_case_entry`)
- Primary keys: `id` (SERIAL)
- Foreign keys: `{table}_id` (e.g., `court_case_id`)
- Timestamps: `created_at`, `updated_at` (where applicable)
- Soft deletes: Not currently implemented

## Coding Guidelines

### Python Code

- **Python Version**: 3.14+ (strictly enforced)
- **Package Manager**: Use `uv` exclusively (not pip/pipenv/poetry)
- **Style**: Black formatter with 120 character line length
- **Linting**: Pylint with 120 character max line length
- **Type Hints**: Encouraged but not strictly required
- **Imports**: Standard library → third-party → local
- **Database Connections**: Use context managers for all database operations
- **Error Handling**: Explicit error handling with informative messages

**Example Database Connection Pattern:**
```python
import psycopg
from contextlib import contextmanager

@contextmanager
def get_db_connection(host, port, database, user, password):
    conn = psycopg.connect(
        host=host, port=port, database=database, 
        user=user, password=password
    )
    try:
        yield conn
    finally:
        conn.close()
```

### VBA Code

- **Naming Convention**: PascalCase for functions/subs (e.g., `BuildForms_DigiDiggie_TNG`)
- **Comments**: Document purpose and usage at top of each module
- **Error Handling**: Use `On Error GoTo ErrorHandler` pattern
- **Form Generation**: Follow existing patterns in `generate_tng_app.vba`
- **Database Access**: Use ODBC DSN connections (not direct connection strings)

### SQL Code

- **Keywords**: UPPERCASE (SELECT, FROM, WHERE, etc.)
- **Identifiers**: snake_case for tables/columns
- **Schema Qualification**: Always qualify tables with schema name in production code
- **Indentation**: 4 spaces
- **Comments**: Use `--` for single-line, `/* */` for multi-line
- **Migrations**: Create new schema files, don't modify existing ones in git history

### Documentation

- **Format**: Markdown for all documentation
- **File Naming**: UPPERCASE-WITH-DASHES.md for docs (e.g., ACCESS-APP-SETUP.md)
- **Structure**: Always include Table of Contents for docs >100 lines
- **Code Examples**: Use fenced code blocks with language specification
- **Links**: Use relative links for internal project references

## Development Workflow

### Setting Up Development Environment

```bash
# Install Python 3.14 and create venv
uv python install 3.14
uv python pin 3.14
uv sync

# Activate virtual environment
source .venv/bin/activate  # Linux/Mac
.venv\Scripts\activate     # Windows

# Start PostgreSQL (Docker)
make up

# Load schema
psql -U postgres -d digidiggie -f src/schema/digidiggie_en_tng.sql
```

### Common Tasks (via Makefile)

- `make up` - Start PostgreSQL Docker container
- `make down` - Stop PostgreSQL Docker container
- `make dump-to-excel` - Export PostgreSQL to Excel
- `make dump-accdb-to-excel` - Export Access database to Excel
- `make load-test-db` - Load test data into PostgreSQL
- `make dev` - Set up development environment

### Testing Database Changes

1. Make changes to schema in a new SQL file (don't modify existing)
2. Test on local PostgreSQL instance
3. Verify Access application still works (forms load, data entry works)
4. Export test data to Excel to verify integrity
5. Document changes in commit message

## MS Access Application

### Architecture

- **Data Layer**: PostgreSQL tables linked via ODBC (read prefix: no prefix)
- **UI Layer**: Access forms generated by VBA code generator
- **Forms Convention**: `frm_{TableName}` (e.g., `frm_CourtCase`)
- **Queries**: Pass-through queries for complex operations

### ODBC DSN Configuration

- **DSN Name**: `DigiDiggie_TNG` (recommended)
- **Type**: System DSN (preferred) or User DSN
- **Driver**: PostgreSQL Unicode (64-bit or 32-bit matching Access)
- **Schema**: `digidiggie_tng`

### Form Generation

Forms are auto-generated using `src/vba_tng_app/generate_tng_app.vba`:
1. Import VBA module into Access
2. Run `BuildForms_DigiDiggie_TNG()` function
3. All forms are created based on linked table structure

**Do not manually edit generated forms** - make changes to VBA generator instead.

## Data Management

### Placename Data (Ortnamn)

- Source: Swedish National Board Survey (Lantmäteriet)
- Format: GeoPackage (.gpkg) or CSV
- Location: `data/ortnamn/`
- Contains: Swedish placenames with coordinates (SWEREF99 TM)

### CSV Data

- Original data: `data/csv/original/` (Swedish headers)
- Ortnamn data: `data/csv/ortnamn/` (place name reference data)
- Character encoding: UTF-8 with BOM (for Excel compatibility)

### Access Database Files

- `digidiggie_dev.accdb` - Development/test database
- `digidiggie_original.accdb` - Original baseline database
- **Never commit large .accdb files** - Use .gitignore

## Important Notes for AI Assistants

### When Working with This Project

1. **Python Version**: Always use Python 3.14+, enforced in pyproject.toml
2. **Package Manager**: Use `uv run`, `uv add`, `uv sync` - never `pip install`
3. **Line Length**: 120 characters (not 80 or 88)
4. **Schema Changes**: Never suggest modifying `digidiggie_en_tng.sql` directly - create migration scripts
5. **VBA Testing**: Cannot be automated easily - always note manual testing required
6. **ODBC Drivers**: Must match Access architecture (32-bit or 64-bit)
7. **Swedish Characters**: Handle åäöÅÄÖ correctly in all string operations
8. **PostGIS**: Placename table uses geometry type, requires PostGIS functions

### Common User Requests

- **Export data**: Use `pgdb_to_excel.py` or `accdb_to_excel.py`
- **Schema changes**: Create new schema file or migration script
- **Form changes**: Modify VBA generator, regenerate forms
- **Database connection issues**: Check ODBC DSN configuration
- **CSV import**: Ensure UTF-8 encoding, handle Swedish characters
- **Placename integration**: Use PostGIS spatial queries

### Known Issues & Limitations

- MS Access forms must be generated on Windows (VBA limitation)
- ODBC drivers can be tricky - 32-bit vs 64-bit confusion
- Swedish character encoding requires UTF-8 BOM for Excel
- PostgreSQL linked tables in Access have performance limitations
- Some complex queries better as pass-through queries in Access

## References

- **Main Documentation**: [README.md](README.md)
- **Access Setup**: [docs/ACCESS-APP-SETUP.md](docs/ACCESS-APP-SETUP.md)
- **ODBC Linking**: [docs/LINKED-DATABASE.md](docs/LINKED-DATABASE.md)
- **UCanAccess**: [docs/UCANACCESS.md](docs/UCANACCESS.md)
- **TODO List**: [TODO.md](TODO.md)

## Questions to Ask Before Coding

1. **Schema changes**: Does this affect the Access application? Will forms need regeneration?
2. **Python scripts**: Are we using `uv run` with correct Python 3.14?
3. **Database operations**: Are we handling Swedish characters correctly?
4. **Export functions**: Should output be Excel, CSV, or both?
5. **Cross-platform**: Does this work on both Linux (dev) and Windows (Access app)?

## Repository Context

- **Owner**: humlab (Humanities Lab)
- **Current Branch**: dev
- **Default Branch**: main
- **Language**: Mixed (Python, VBA, SQL, Shell)
- **Primary Purpose**: Academic research tool for digital humanities
