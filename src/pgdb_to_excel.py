"""Export PostgreSQL database tables to Excel file.

This script connects to a PostgreSQL database and exports all tables
to separate sheets in an Excel file.
"""

import sys
from pathlib import Path
from typing import Optional

import click
import pandas as pd
import psycopg
from sqlalchemy import create_engine


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
    default="public",
    help="Database schema to export",
    show_default=True,
)
@click.option(
    "--output",
    "-o",
    type=click.Path(path_type=Path),
    required=True,
    help="Output Excel file path",
)
@click.option(
    "--exclude",
    "-e",
    multiple=True,
    help="Tables to exclude (can be specified multiple times)",
)
@click.option(
    "--include",
    "-i",
    multiple=True,
    help="Tables to include (if specified, only these tables will be exported)",
)
@click.option("--verbose", "-v", is_flag=True, help="Enable verbose output")
def export_db_to_excel(
    host: str,
    port: int,
    database: str,
    user: str,
    password: Optional[str],
    prompt_password: bool,
    schema: str,
    output: Path,
    exclude: tuple[str, ...],
    include: tuple[str, ...],
    verbose: bool,
) -> None:
    """Export all tables from PostgreSQL database to Excel file.

    Each table will be exported to a separate sheet in the Excel file.
    Sheet names will match table names (truncated to 31 characters if needed).

    By default, uses .pgpass for authentication if no password is provided.
    """
    # Prompt for password if --prompt-password flag is set
    if prompt_password and not password:
        password = click.prompt("Password", hide_input=True)

    # Build connection strings for both psycopg and SQLAlchemy
    # If password is not provided, psycopg will use .pgpass file
    conninfo_parts = [
        f"host={host}",
        f"port={port}",
        f"dbname={database}",
        f"user={user}",
    ]

    # SQLAlchemy connection string
    if password:
        conninfo_parts.append(f"password={password}")
        sqlalchemy_url = (
            f"postgresql+psycopg://{user}:{password}@{host}:{port}/{database}"
        )
    else:
        sqlalchemy_url = f"postgresql+psycopg://{user}@{host}:{port}/{database}"

    conninfo = " ".join(conninfo_parts)

    try:
        # Connect to database
        if verbose:
            click.echo(f"Connecting to database '{database}' at {host}:{port}...")

        # Use psycopg for getting table list
        with psycopg.connect(conninfo) as conn:
            # Get list of tables
            tables = get_table_list(conn, schema)

            if not tables:
                click.echo(f"No tables found in schema '{schema}'", err=True)
                sys.exit(1)

            # Filter tables based on include/exclude
            tables = filter_tables(tables, include, exclude)

            if not tables:
                click.echo("No tables to export after filtering", err=True)
                sys.exit(1)

            if verbose:
                click.echo(f"Found {len(tables)} table(s) to export:")
                for table in tables:
                    click.echo(f"  - {table}")

        # Create SQLAlchemy engine for pandas operations
        engine = create_engine(sqlalchemy_url)

        # Export tables to Excel
        export_tables_to_excel(engine, schema, tables, output, verbose)

        click.echo(f"✓ Successfully exported {len(tables)} table(s) to {output}")

    except psycopg.Error as e:
        click.echo(f"Database error: {e}", err=True)
        sys.exit(1)
    except Exception as e:  # type: ignore ; # pylint: disable=broad-except
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


def get_table_list(conn: psycopg.Connection, schema: str) -> list[str]:
    """Get list of all tables in the specified schema.

    Args:
        conn: Database connection
        schema: Schema name

    Returns:
        List of table names
    """
    query = """
        SELECT table_name
        FROM information_schema.tables
        WHERE table_schema = %s
          AND table_type = 'BASE TABLE'
        ORDER BY table_name
    """

    with conn.cursor() as cur:
        cur.execute(query, (schema,))
        tables = [row[0] for row in cur.fetchall()]

    return tables


def filter_tables(
    tables: list[str], include: tuple[str, ...], exclude: tuple[str, ...]
) -> list[str]:
    """Filter table list based on include/exclude patterns.

    Args:
        tables: List of all tables
        include: Tables to include (if specified, only these are kept)
        exclude: Tables to exclude

    Returns:
        Filtered list of tables
    """
    # If include list is specified, only keep those tables
    if include:
        tables = [t for t in tables if t in include]

    # Remove excluded tables
    if exclude:
        tables = [t for t in tables if t not in exclude]

    return tables


def export_tables_to_excel(
    engine,
    schema: str,
    tables: list[str],
    output_path: Path,
    verbose: bool,
) -> None:
    """Export tables to Excel file with separate sheets.

    Args:
        engine: SQLAlchemy engine
        schema: Schema name
        tables: List of table names to export
        output_path: Path to output Excel file
        verbose: Enable verbose output
    """
    # Create Excel writer
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for table_name in tables:
            if verbose:
                click.echo(f"Exporting table: {table_name}")

            # Read table data using pandas
            query = f'SELECT * FROM "{schema}"."{table_name}"'

            try:
                df = pd.read_sql_query(query, engine)
                # Excel sheet names are limited to 31 characters
                sheet_name = table_name[:31]

                # Write to Excel
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                if verbose:
                    click.echo(f"  → Exported {len(df)} row(s) to sheet '{sheet_name}'")

            except Exception as e:  # type: ignore ; # pylint: disable=broad-except
                click.echo(f"  ✗ Error exporting table '{table_name}': {e}", err=True)
                continue


if __name__ == "__main__":
    export_db_to_excel()  # type: ignore ; # pylint: disable=no-value-for-parameter
