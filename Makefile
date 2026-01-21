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

