"""Export MS Access database tables to Excel file.

This script connects to an MS Access database using UCanAccess JDBC driver
and exports all tables to separate sheets in an Excel file.
"""

import re
import sys
import unicodedata
from pathlib import Path
from typing import Optional

import click
import jaydebeapi
import pandas as pd


@click.command()
@click.option(
    "--database",
    "-d",
    type=click.Path(exists=True, path_type=Path),
    required=True,
    help="MS Access database file path (.accdb or .mdb)",
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
@click.option(
    "--ucanaccess-jar",
    type=click.Path(exists=True, path_type=Path),
    help="Path to UCanAccess JAR directory (defaults to ./lib/ucanaccess)",
)
@click.option("--verbose", "-v", is_flag=True, help="Enable verbose output")
def export_accdb_to_excel(
    database: Path,
    output: Path,
    exclude: tuple[str, ...],
    include: tuple[str, ...],
    ucanaccess_jar: Optional[Path],
    verbose: bool,
) -> None:
    """Export all tables from MS Access database to Excel file.

    Each table will be exported to a separate sheet in the Excel file.
    Sheet names will match table names (truncated to 31 characters if needed).

    Requires UCanAccess JDBC driver jars to be available.
    """
    # Set default UCanAccess path if not provided
    if not ucanaccess_jar:
        ucanaccess_jar = Path("./lib/ucanaccess")

    # Verify UCanAccess exists
    if not ucanaccess_jar.exists():
        click.echo(
            f"Error: UCanAccess JAR directory not found at {ucanaccess_jar}",
            err=True,
        )
        click.echo(
            "Please download UCanAccess and extract to ./lib/ucanaccess "
            "or specify path with --ucanaccess-jar",
            err=True,
        )
        sys.exit(1)

    # Build classpath for UCanAccess and dependencies
    jar_files = list(ucanaccess_jar.glob("*.jar"))
    if not jar_files:
        click.echo(
            f"Error: No JAR files found in {ucanaccess_jar}",
            err=True,
        )
        sys.exit(1)

    classpath = ":".join(str(jar) for jar in jar_files)

    if verbose:
        click.echo(f"Using UCanAccess from: {ucanaccess_jar}")
        click.echo(f"Found {len(jar_files)} JAR files")

    try:
        # Connect to Access database
        if verbose:
            click.echo(f"Connecting to database: {database}")

        jdbc_url = f"jdbc:ucanaccess://{database.absolute()}"
        conn = jaydebeapi.connect(
            "net.ucanaccess.jdbc.UcanaccessDriver",
            jdbc_url,
            ["", ""],  # username, password (empty for Access)
            classpath,
        )

        try:
            # Get list of tables
            tables = get_table_list(conn)

            if not tables:
                click.echo("No tables found in database", err=True)
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

            # Export tables to Excel
            export_tables_to_excel(conn, tables, output, verbose)

            click.echo(f"✓ Successfully exported {len(tables)} table(s) to {output}")

        finally:
            conn.close()

    except jaydebeapi.Error as e:
        click.echo(f"Database error: {e}", err=True)
        sys.exit(1)
    except Exception as e:  # pylint: disable=broad-except
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


def get_table_list(conn) -> list[str]:
    """Get list of all user tables in the Access database.

    Args:
        conn: JDBC database connection

    Returns:
        List of table names
    """
    cursor = conn.cursor()

    # Get table metadata - excluding system tables
    cursor.execute(
        """
        SELECT Name FROM MSysObjects 
        WHERE Type=1 AND Flags=0
        ORDER BY Name
    """
    )

    tables = [row[0] for row in cursor.fetchall()]
    cursor.close()

    # Filter out system tables
    tables = [t for t in tables if not t.startswith("MSys") and not t.startswith("~")]

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


def sanitize_name(name: str) -> str:
    """Sanitize table/column name: lowercase, replace spaces with _, remove non-ASCII.

    Args:
        name: Original name to sanitize

    Returns:
        Sanitized name
    """
    # Convert to lowercase
    name = name.lower()

    # Replace whitespace (spaces, tabs, newlines) with underscores
    name = re.sub(r'\s+', '_', name)

    # Remove or replace non-ASCII characters
    # First try to normalize (e.g., ä -> a)
    name = unicodedata.normalize('NFKD', name)
    name = name.encode('ascii', 'ignore').decode('ascii')

    # Replace any remaining non-alphanumeric characters (except underscore) with underscore
    name = re.sub(r'[^a-z0-9_]', '_', name)

    # Remove consecutive underscores
    name = re.sub(r'_+', '_', name)

    # Remove leading/trailing underscores
    name = name.strip('_')

    # Ensure name is not empty
    if not name:
        name = 'unnamed'

    return name


def export_tables_to_excel(
    conn,
    tables: list[str],
    output_path: Path,
    verbose: bool,
) -> None:
    """Export tables to Excel file with separate sheets.

    Args:
        conn: JDBC database connection
        tables: List of table names to export
        output_path: Path to output Excel file
        verbose: Enable verbose output
    """
    # Create Excel writer
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for table_name in tables:
            if verbose:
                click.echo(f"Exporting table: {table_name}")

            try:
                # Read table data using pandas
                query = f"SELECT * FROM [{table_name}]"
                cursor = conn.cursor()
                cursor.execute(query)

                # Fetch data and column names
                columns = [desc[0] for desc in cursor.description]
                data = cursor.fetchall()
                cursor.close()

                # Create DataFrame
                df = pd.DataFrame(data, columns=columns)

                            # Sanitize column names
                df.columns = [sanitize_name(col) for col in df.columns]

                # Sanitize and truncate sheet name (Excel limit is 31 characters)
                sheet_name = sanitize_name(table_name)[:31]

                # Write to Excel
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                if verbose:
                    click.echo(f"  → Exported {len(df)} row(s) to sheet '{sheet_name}'")

            except Exception as e:  # pylint: disable=broad-except
                click.echo(f"  ✗ Error exporting table '{table_name}': {e}", err=True)
                continue


if __name__ == "__main__":
    export_accdb_to_excel()  # pylint: disable=no-value-for-parameter
