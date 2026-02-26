include docker/.env
SHELL := /bin/bash

up:
	@cd docker && docker compose up -d && cd -

down:
	@cd docker && docker compose down && cd -

restart: down up


dev:
	@uv python install 3.14
	@uv python pin 3.14
	@uv add pyodbc pandas xlsxwriter openpyxl requests "psycopg[binary]>=3,<4"
	@uv venv .venv


load-test-db: 
	@./scripts/mdb-to-pg load --database $(TEST_POSTGRES_DB) --schema public --host $(HOST) --user $(TEST_POSTGRES_USER) --port $(TEST_POSTGRES_PORT) data/digidiggie_dev.accdb

load-prod-db: 
	@./scripts/mdb-to-pg load --database $(POSTGRES_DB) --schema digidiggie_tog --host $(HOST) --user $(POSTGRES_USER) --port $(POSTGRES_PORT) data/digidiggie_dev.accdb

dump-to-excel:
	@uv run src/pgdb_to_excel.py --host localhost --port $(POSTGRES_PORT) --database $(POSTGRES_DB) --schema digidiggie_tng --user $(POSTGRES_USER) --output xl.xlsx

dump-accdb-to-excel:
	@uv run src/accdb_to_excel.py --database data/digidiggie_dev.accdb --output data/accdb_dump.xlsx --verbose

