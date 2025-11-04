
up:
	@cd docker && docker compose up -d && cd -

down:
	@cd docker && docker compose down && cd -

restart: down up
