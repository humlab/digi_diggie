"""Generate Markdown documentation from PostgreSQL database schema comments.

This script connects to a PostgreSQL database and generates comprehensive
markdown documentation by reading table and column comments, along with
data types and constraints.
"""

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
    help="Database schema to document",
    show_default=True,
)
@click.option(
    "--output",
    "-o",
    type=click.Path(path_type=Path),
    required=True,
    help="Output markdown file path",
)
@click.option("--verbose", "-v", is_flag=True, help="Enable verbose output")
@click.option(
    "--title",
    default="Database Schema Documentation",
    help="Document title",
    show_default=True,
)
def generate_schema_docs(
    host: str,
    port: int,
    database: str,
    user: str,
    password: Optional[str],
    prompt_password: bool,
    schema: str,
    output: Path,
    verbose: bool,
    title: str,
) -> None:
    """Generate markdown documentation from PostgreSQL schema comments.

    Reads table and column comments from the database and generates a
    comprehensive markdown document with data types and descriptions.

    By default, uses .pgpass for authentication if no password is provided.
    """
    # Prompt for password if --prompt-password flag is set
    if prompt_password and not password:
        password = click.prompt("Password", hide_input=True)

    # Build connection string for psycopg
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
                    click.echo(f"Fetching schema information for '{schema}'...")

                # Get schema comment
                schema_comment = get_schema_comment(cur, schema)

                # Get all tables with their comments
                tables = get_tables_with_comments(cur, schema)

                if not tables:
                    click.echo(f"No tables found in schema '{schema}'", err=True)
                    sys.exit(1)

                if verbose:
                    click.echo(f"Found {len(tables)} table(s)")

                # Get columns for each table
                table_columns = {}
                for table_name in tables.keys():
                    columns = get_columns_with_comments(cur, schema, table_name)
                    table_columns[table_name] = columns

                    if verbose:
                        click.echo(f"  - {table_name}: {len(columns)} column(s)")

        # Generate markdown documentation
        if verbose:
            click.echo(f"Generating markdown documentation...")

        markdown = generate_markdown(
            schema, schema_comment, tables, table_columns, title
        )

        # Write to file
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(markdown, encoding="utf-8")

        click.echo(f"✓ Successfully generated documentation: {output}")

    except psycopg.Error as e:
        click.echo(f"Database error: {e}", err=True)
        sys.exit(1)
    except Exception as e:  # type: ignore ; # pylint: disable=broad-except
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


def get_schema_comment(cur: psycopg.Cursor, schema: str) -> Optional[str]:
    """Get the comment for a schema.

    Args:
        cur: Database cursor
        schema: Schema name

    Returns:
        Schema comment or None
    """
    query = """
        SELECT obj_description(oid, 'pg_namespace') AS comment
        FROM pg_namespace
        WHERE nspname = %s
    """
    cur.execute(query, (schema,))
    result = cur.fetchone()
    return result[0] if result and result[0] else None


def get_tables_with_comments(cur: psycopg.Cursor, schema: str) -> dict[str, Optional[str]]:
    """Get all tables in schema with their comments.

    Args:
        cur: Database cursor
        schema: Schema name

    Returns:
        Dictionary mapping table names to their comments
    """
    query = """
        SELECT 
            c.relname AS table_name,
            pg_catalog.obj_description(c.oid, 'pg_class') AS comment
        FROM pg_catalog.pg_class c
        JOIN pg_catalog.pg_namespace n ON n.oid = c.relnamespace
        WHERE n.nspname = %s
          AND c.relkind = 'r'
        ORDER BY c.relname
    """
    cur.execute(query, (schema,))
    return {row[0]: row[1] for row in cur.fetchall()}


def get_columns_with_comments(
    cur: psycopg.Cursor, schema: str, table: str
) -> list[dict[str, Optional[str]]]:
    """Get all columns for a table with their comments and data types.

    Args:
        cur: Database cursor
        schema: Schema name
        table: Table name

    Returns:
        List of column information dictionaries
    """
    query = """
        SELECT 
            a.attname AS column_name,
            pg_catalog.format_type(a.atttypid, a.atttypmod) AS data_type,
            a.attnotnull AS not_null,
            pg_catalog.col_description(c.oid, a.attnum) AS comment,
            pg_get_expr(ad.adbin, ad.adrelid) AS default_value
        FROM pg_catalog.pg_attribute a
        JOIN pg_catalog.pg_class c ON a.attrelid = c.oid
        JOIN pg_catalog.pg_namespace n ON n.oid = c.relnamespace
        LEFT JOIN pg_catalog.pg_attrdef ad ON a.attrelid = ad.adrelid AND a.attnum = ad.adnum
        WHERE n.nspname = %s
          AND c.relname = %s
          AND a.attnum > 0
          AND NOT a.attisdropped
        ORDER BY a.attnum
    """
    cur.execute(query, (schema, table))
    
    columns = []
    for row in cur.fetchall():
        columns.append({
            'name': row[0],
            'type': row[1],
            'not_null': row[2],
            'comment': row[3],
            'default': row[4]
        })
    
    return columns


def generate_markdown(
    schema: str,
    schema_comment: Optional[str],
    tables: dict[str, Optional[str]],
    table_columns: dict[str, list[dict[str, Optional[str]]]],
    title: str,
) -> str:
    """Generate markdown documentation.

    Args:
        schema: Schema name
        schema_comment: Schema description
        tables: Dictionary of table names to comments
        table_columns: Dictionary of table names to their columns
        title: Document title

    Returns:
        Markdown formatted documentation
    """
    lines = []
    
    # Title and header
    lines.append(f"# {title}")
    lines.append("")
    lines.append(f"**Schema:** `{schema}`")
    lines.append("")
    
    if schema_comment:
        lines.append(f"{schema_comment}")
        lines.append("")
    
    # Table of contents
    lines.append("## Table of Contents")
    lines.append("")
    
    # Categorize tables
    lookup_tables = []
    entity_tables = []
    
    for table_name, comment in tables.items():
        if comment and 'lookup' in comment.lower():
            lookup_tables.append(table_name)
        else:
            entity_tables.append(table_name)
    
    if entity_tables:
        lines.append("### Entity Tables")
        lines.append("")
        for table_name in entity_tables:
            lines.append(f" [{table_name}](#{table_name})")
        lines.append("")
    
    if lookup_tables:
        lines.append("### Lookup Tables")
        lines.append("")
        for table_name in lookup_tables:
            lines.append(f" [{table_name}](#{table_name})")
        lines.append("")
    
    lines.append("---")
    lines.append("")
    
    # Entity tables first
    if entity_tables:
        lines.append("## Entity Tables")
        lines.append("")
        for table_name in entity_tables:
            lines.extend(generate_table_section(table_name, tables[table_name], table_columns[table_name]))
    
    # Lookup tables
    if lookup_tables:
        lines.append("## Lookup Tables")
        lines.append("")
        for table_name in lookup_tables:
            lines.extend(generate_table_section(table_name, tables[table_name], table_columns[table_name]))
    
    return "\n".join(lines)


def generate_table_section(
    table_name: str,
    table_comment: Optional[str],
    columns: list[dict[str, Optional[str]]],
) -> list[str]:
    """Generate markdown section for a single table.

    Args:
        table_name: Name of the table
        table_comment: Table description
        columns: List of column information

    Returns:
        List of markdown lines
    """
    lines = []
    
    lines.append(f"### {table_name}")
    lines.append("")
    
    if table_comment:
        lines.append(f"{table_comment}")
        lines.append("")
    
    # Columns table
    lines.append("| Column | Type | Nullable | Description |")
    lines.append("|--------|------|----------|-------------|")
    
    for col in columns:
        col_name = f"`{col['name']}`"
        col_type = f"`{col['type']}`"
        nullable = "✓" if not col['not_null'] else ""
        description = col['comment'] if col['comment'] else "*No description*"
        
        # Escape pipe characters in description
        description = description.replace("|", "\\|")
        
        lines.append(f"| {col_name} | {col_type} | {nullable} | {description} |")
    
    lines.append("")
    
    return lines


if __name__ == "__main__":
    generate_schema_docs()  # pylint: disable=no-value-for-parameter
