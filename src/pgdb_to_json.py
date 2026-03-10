"""Export PostgreSQL database schema to JSON file.

This script connects to a PostgreSQL database and calls the export_schema_as_json
function to export the entire schema as a single JSON document.
"""

import json
import sys
from pathlib import Path
from typing import Optional

import click
import psycopg


@click.command()
@click.option(
    "--host", default="localhost", help="PostgreSQL host address", show_default=True
)
@click.option(
    "--port", default=5432, type=int, help="PostgreSQL port", show_default=True
)
@click.option("--database", "-d", required=True, help="PostgreSQL database name")
@click.option("--user", "-u", required=True, help="PostgreSQL username")
@click.option(
    "--password",
    "-p",
    help="PostgreSQL password (uses .pgpass if not provided)",
)
@click.option(
    "--prompt-password", is_flag=True, help="Prompt for password interactively"
)
@click.option(
    "--schema",
    "-s",
    default="digidiggie_tng",
    help="Database schema to export",
    show_default=True,
)
@click.option(
    "--output",
    "-o",
    type=click.Path(path_type=Path),
    required=True,
    help="Output JSON file path",
)
@click.option("--verbose", "-v", is_flag=True, help="Enable verbose output")
@click.option(
    "--pretty",
    is_flag=True,
    help="Pretty-print JSON output with indentation",
    default=False,
)
def export_schema_to_json(
    host: str,
    port: int,
    database: str,
    user: str,
    password: Optional[str],
    prompt_password: bool,
    schema: str,
    output: Path,
    verbose: bool,
    pretty: bool,
) -> None:
    """Export entire database schema to JSON file.

    Calls the export_schema_as_json() PostgreSQL function to retrieve
    all tables and their data as a single JSON document.

    By default, uses .pgpass for authentication if no password is provided.
    """
    # Prompt for password if --prompt-password flag is set
    if prompt_password and not password:
        password = click.prompt("Password", hide_input=True)

    # Build connection string for psycopg
    # If password is not provided, psycopg will use .pgpass file
    conninfo_parts = [
        f"host={host}",
        f"port={port}",
        f"dbname={database}",
        f"user={user}",
    ]

    if password:
        conninfo_parts.append(f"password={password}")

    conninfo = " ".join(conninfo_parts)

    try:
        # Connect to database
        if verbose:
            click.echo(f"Connecting to database '{database}' at {host}:{port}...")

        with psycopg.connect(conninfo) as conn:
            with conn.cursor() as cur:
                if verbose:
                    click.echo(f"Calling export_schema_as_json('{schema}')...")

                # Call the export_schema_as_json function
                cur.execute(
                    "SELECT export_schema_as_json(%s)",
                    (schema,)
                )
                result = cur.fetchone()

                if not result or result[0] is None:
                    click.echo(f"No data returned from export_schema_as_json('{schema}')", err=True)
                    sys.exit(1)

                json_data = result[0]

                if verbose:
                    click.echo(f"Retrieved schema data for '{schema}'")
                    if isinstance(json_data, dict) and 'tables' in json_data:
                        table_count = len(json_data['tables'])
                        click.echo(f"Found {table_count} table(s) in schema")

        # Write JSON to file
        if verbose:
            click.echo(f"Writing JSON to {output}...")

        output.parent.mkdir(parents=True, exist_ok=True)

        with open(output, 'w', encoding='utf-8') as f:
            if pretty:
                json.dump(json_data, f, indent=2, ensure_ascii=False)
            else:
                json.dump(json_data, f, ensure_ascii=False)

        # Get file size for reporting
        file_size = output.stat().st_size
        size_mb = file_size / (1024 * 1024)

        if size_mb > 1:
            size_str = f"{size_mb:.2f} MB"
        else:
            size_str = f"{file_size / 1024:.2f} KB"

        click.echo(f"✓ Successfully exported schema '{schema}' to {output} ({size_str})")

    except psycopg.Error as e:
        click.echo(f"Database error: {e}", err=True)
        sys.exit(1)
    except Exception as e:  # type: ignore ; # pylint: disable=broad-except
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


if __name__ == "__main__":
    export_schema_to_json()  # pylint: disable=no-value-for-parameter
