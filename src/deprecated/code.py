# import pyodbc
# import psycopg
# import sys
# import logging

# # --- Configuration ---
# # MS Access Configuration
# ACCESS_DB_FILE = r'C:\path\to\your\database.accdb'  # Use 'r' for raw string or double backslashes '\\'
# # Or for .mdb files: r'C:\path\to\your\database.mdb'
# # Find the exact driver name in your ODBC Data Sources (x64 or x32) Admin tool
# # Common names:
# ACCESS_DRIVER = '{Microsoft Access Driver (*.mdb, *.accdb)}'
# # ACCESS_DRIVER = '{Microsoft Access Driver (*.mdb)}' # Older driver

# # PostgreSQL Configuration
# PG_HOST = 'localhost'
# PG_PORT = '5432'
# PG_DATABASE = 'your_postgres_db'
# PG_USER = 'your_postgres_user'
# PG_PASSWORD = 'your_postgres_password'
# PG_SCHEMA = 'public'  # Target schema in PostgreSQL

# # --- Optional Settings ---
# # List of tables to skip (e.g., system tables or tables you don't want)
# SKIP_TABLES = ['MSysObjects', 'MSysACEs', 'MSysQueries', 'MSysRelationships']
# # Convert Access table/column names to lowercase in PostgreSQL?
# CONVERT_TO_LOWERCASE = True
# # How many rows to fetch/insert in each batch
# BATCH_SIZE = 1000
# # If a table already exists in PostgreSQL, what to do: 'skip', 'drop', 'fail'
# TABLE_EXISTS_ACTION = 'skip' # 'drop' will DELETE existing data! Use with caution.

# # --- Logging Setup ---
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# # --- Data Type Mapping (Simplified - Adjust as needed!) ---
# # Add more specific mappings if required (e.g., specific NUMBER subtypes)
# ACCESS_TO_PG_TYPES = {
#     'COUNTER': 'SERIAL', # Often maps to primary key SERIAL
#     'AUTOINCREMENT': 'SERIAL',
#     'IDENTITY': 'SERIAL',
#     'LONG': 'INTEGER',
#     'INTEGER': 'INTEGER',
#     'SHORT': 'SMALLINT',
#     'BYTE': 'SMALLINT', # Smallest integer type
#     'SINGLE': 'REAL',
#     'DOUBLE': 'DOUBLE PRECISION',
#     'CURRENCY': 'NUMERIC(19, 4)', # Or DECIMAL
#     'DECIMAL': 'NUMERIC', # Precision/Scale might need adjustment
#     'NUMBER': 'DOUBLE PRECISION', # Generic NUMBER often float in Access ODBC
#     'TEXT': 'TEXT',
#     'MEMO': 'TEXT',
#     'LONGCHAR': 'TEXT',
#     'VARCHAR': 'VARCHAR', # Length might need adjustment
#     'DATETIME': 'TIMESTAMP',
#     'DATETIME2': 'TIMESTAMP', # SQL Server linked table type
#     'DATE': 'DATE',
#     'TIME': 'TIME',
#     'YESNO': 'BOOLEAN',
#     'BIT': 'BOOLEAN',
#     'OLEOBJECT': 'BYTEA',
#     'LONGBINARY': 'BYTEA',
#     'BINARY': 'BYTEA',
#     'VARBINARY': 'BYTEA',
#     'GUID': 'UUID',
#     'UNIQUEIDENTIFIER': 'UUID',
#     'UNKNOWN': 'TEXT' # Fallback for unknown types
#     # Add Hyperlink -> TEXT or specific handling if needed
# }

# def get_access_tables(access_cursor):
#     """Gets a list of user tables from MS Access."""
#     tables = []
#     # System tables often start with 'MSys' or '~'. User tables are 'TABLE'.
#     for row in access_cursor.tables(tableType='TABLE'):
#         table_name = row.table_name
#         if not table_name.lower().startswith('msys') and not table_name.startswith('~'):
#              if table_name not in SKIP_TABLES:
#                 tables.append(table_name)
#     return tables

# def get_access_columns(access_cursor, table_name):
#     """Gets column names, data types, and nullability for an Access table."""
#     columns = []
#     try:
#         for row in access_cursor.columns(table=table_name):
#             # pyodbc column attributes (indexes might vary slightly based on driver version)
#             # 0: table_cat, 1: table_schem, 2: table_name, 3: column_name, 4: data_type (numeric)
#             # 5: type_name, 6: column_size, 7: buffer_length, 8: decimal_digits, 9: num_prec_radix
#             # 10: nullable, 11: remarks ...
#             col_name = row.column_name
#             type_name = row.type_name.upper() # Normalize type name
#             nullable = row.nullable == pyodbc.SQL_NULLABLE # Check if column allows NULLs
#             columns.append({'name': col_name, 'type': type_name, 'nullable': nullable})
#     except Exception as e:
#         logging.error(f"Error getting columns for table '{table_name}': {e}")
#         # Handle specific errors if needed (e.g., table not found although it should be)
#     return columns

# def get_pg_type(access_type):
#     """Maps Access data type string to PostgreSQL type string."""
#     access_type = access_type.upper()
#     # Handle cases like TEXT(50) -> TEXT
#     if '(' in access_type:
#         access_type = access_type.split('(')[0]
#     return ACCESS_TO_PG_TYPES.get(access_type, 'TEXT') # Default to TEXT if unknown

# def create_pg_table(pg_cursor, table_name, columns):
#     """Creates a table in PostgreSQL based on Access schema."""
#     target_table_name = table_name.lower() if CONVERT_TO_LOWERCASE else table_name
#     pg_cursor.execute(f"SELECT EXISTS (SELECT FROM information_schema.tables WHERE table_schema = %s AND table_name = %s);", (PG_SCHEMA, target_table_name))
#     table_exists = pg_cursor.fetchone()[0]

#     if table_exists:
#         logging.warning(f"Table '{PG_SCHEMA}.{target_table_name}' already exists in PostgreSQL.")
#         if TABLE_EXISTS_ACTION == 'skip':
#             logging.warning("Skipping creation.")
#             return False # Indicate table was skipped
#         elif TABLE_EXISTS_ACTION == 'drop':
#             logging.warning(f"Dropping existing table '{PG_SCHEMA}.{target_table_name}'!")
#             try:
#                 pg_cursor.execute(f'DROP TABLE "{PG_SCHEMA}"."{target_table_name}" CASCADE;') # Use quotes for safety
#                 logging.info(f"Table '{PG_SCHEMA}.{target_table_name}' dropped.")
#             except psycopg.Error as e:
#                  logging.error(f"Failed to drop table '{PG_SCHEMA}.{target_table_name}': {e}")
#                  pg_cursor.connection.rollback() # Rollback the drop attempt
#                  raise # Re-raise the exception to stop the process for this table
#         elif TABLE_EXISTS_ACTION == 'fail':
#             logging.error("Failing migration as table exists.")
#             raise Exception(f"Table '{PG_SCHEMA}.{target_table_name}' already exists.")
#         # If drop was successful, continue to create below

#     col_definitions = []
#     for col in columns:
#         col_name = col['name'].lower() if CONVERT_TO_LOWERCASE else col['name']
#         pg_type = get_pg_type(col['type'])
#         # Special handling for SERIAL - it implies NOT NULL and is often a PRIMARY KEY
#         # Simple approach: Assume COUNTER/AUTOINCREMENT is the PK. More robust: check primary keys separately.
#         if pg_type == 'SERIAL':
#              col_def = f'"{col_name}" SERIAL PRIMARY KEY' # SERIAL implies NOT NULL
#         else:
#             null_clause = "" if not col['nullable'] else "NULL" # NULL is default, explicit NOT NULL needed
#             if not col['nullable']:
#                 null_clause = "NOT NULL"
#             col_def = f'"{col_name}" {pg_type} {null_clause}' # Quote column names

#         col_definitions.append(col_def)

#     create_sql = f'CREATE TABLE "{PG_SCHEMA}"."{target_table_name}" ({", ".join(col_definitions)});'

#     try:
#         logging.info(f"Executing: {create_sql}")
#         pg_cursor.execute(create_sql)
#         logging.info(f"Table '{PG_SCHEMA}.{target_table_name}' created successfully.")
#         return True # Indicate table was created
#     except psycopg.Error as e:
#         logging.error(f"Error creating table '{PG_SCHEMA}.{target_table_name}': {e}")
#         pg_cursor.connection.rollback() # Rollback table creation
#         raise # Stop processing this table

# def transfer_data(access_cursor, pg_cursor, table_name, columns):
#     """Transfers data from an Access table to a PostgreSQL table."""
#     target_table_name = table_name.lower() if CONVERT_TO_LOWERCASE else table_name
#     access_cols = [f"[{col['name']}]" for col in columns] # Use [] for Access names with spaces/special chars
#     pg_cols = [f'"{col["name"].lower() if CONVERT_TO_LOWERCASE else col["name"]}"' for col in columns]
#     placeholders = ", ".join(["%s"] * len(columns))

#     select_sql = f"SELECT {', '.join(access_cols)} FROM [{table_name}]"
#     insert_sql = f'INSERT INTO "{PG_SCHEMA}"."{target_table_name}" ({", ".join(pg_cols)}) VALUES ({placeholders})'

#     logging.info(f"Transferring data for table '{table_name}'...")
#     access_cursor.execute(select_sql)

#     rows_processed = 0
#     batch = []
#     while True:
#         try:
#             rows = access_cursor.fetchmany(BATCH_SIZE)
#             if not rows:
#                 break # No more data

#             # Process rows for insertion (handle potential type issues if needed)
#             processed_rows = []
#             for row in rows:
#                  # Basic type conversion check (more specific handling can be added)
#                  processed_row = []
#                  for i, value in enumerate(row):
#                      # Example: Handle potential encoding issues if reading text fields
#                      # if isinstance(value, bytes) and get_pg_type(columns[i]['type']) == 'TEXT':
#                      #     try:
#                      #         value = value.decode('utf-8') # Or appropriate encoding
#                      #     except UnicodeDecodeError:
#                      #         logging.warning(f"Could not decode value in table {table_name}, col {columns[i]['name']}, row {rows_processed+len(processed_rows)+1}. Inserting as NULL.")
#                      #         value = None
#                      # Add more specific type cleanup if needed here based on errors
#                      processed_row.append(value)
#                  processed_rows.append(tuple(processed_row)) # executemany needs list of tuples


#             pg_cursor.executemany(insert_sql, processed_rows)
#             rows_processed += len(processed_rows)
#             logging.info(f"  {rows_processed} rows transferred for '{table_name}'...")

#             # Commit periodically for large tables
#             if rows_processed % (BATCH_SIZE * 10) == 0:
#                 pg_cursor.connection.commit()
#                 logging.info(f"  Committed transaction for '{table_name}'.")


#         except pyodbc.Error as e:
#             logging.error(f"Error fetching data from Access table '{table_name}': {e}")
#             raise # Stop processing this table
#         except psycopg.Error as e:
#             logging.error(f"Error inserting batch into PostgreSQL table '{target_table_name}': {e}")
#             logging.error(f"  Problematic INSERT statement likely: {insert_sql}")
#             # Log problematic data if possible (be careful with sensitive data)
#             # logging.error(f"  Problematic data (first row of batch): {processed_rows[0] if processed_rows else 'N/A'}")
#             pg_cursor.connection.rollback() # Rollback the failed batch
#             # Option 1: Stop for this table
#             raise
#             # Option 2: Skip batch and continue (data loss!)
#             # logging.warning("Skipping current batch due to error.")
#             # continue
#         except Exception as e:
#              logging.error(f"An unexpected error occurred during data transfer for '{table_name}': {e}")
#              raise

#     pg_cursor.connection.commit() # Final commit for the table
#     logging.info(f"Finished transferring data for table '{table_name}'. Total rows: {rows_processed}")


# # --- Main Execution ---
# access_conn = None
# pg_conn = None

# try:
#     # Connect to MS Access
#     access_conn_str = (
#         f'DRIVER={ACCESS_DRIVER};'
#         f'DBQ={ACCESS_DB_FILE};'
#         # Add UID and PWD if your Access DB is password protected:
#         # f'UID=your_access_user;'
#         # f'PWD=your_access_password;'
#     )
#     logging.info(f"Connecting to MS Access: {ACCESS_DB_FILE}")
#     access_conn = pyodbc.connect(access_conn_str)
#     access_cursor = access_conn.cursor()
#     logging.info("MS Access connection successful.")

#     # Connect to PostgreSQL
#     logging.info(f"Connecting to PostgreSQL: db='{PG_DATABASE}' host='{PG_HOST}' user='{PG_USER}'")
#     pg_conn = psycopg.connect(
#         host=PG_HOST,
#         port=PG_PORT,
#         database=PG_DATABASE,
#         user=PG_USER,
#         password=PG_PASSWORD
#     )
#     # Optional: Set schema search path if not using public or want to be explicit
#     # pg_conn.cursor().execute(f"SET search_path TO {PG_SCHEMA}, public;")
#     pg_cursor = pg_conn.cursor()
#     logging.info("PostgreSQL connection successful.")

#     # Get list of tables from Access
#     tables_to_process = get_access_tables(access_cursor)
#     logging.info(f"Found {len(tables_to_process)} user tables in Access: {tables_to_process}")

#     # Process each table
#     for table_name in tables_to_process:
#         logging.info(f"--- Processing table: {table_name} ---")

#         # Get column definitions from Access
#         columns = get_access_columns(access_cursor, table_name)
#         if not columns:
#             logging.warning(f"Could not get columns for table '{table_name}'. Skipping.")
#             continue
#         logging.info(f"  Found {len(columns)} columns in '{table_name}'.")

#         # Create table in PostgreSQL
#         try:
#             table_created_or_exists = create_pg_table(pg_cursor, table_name, columns)
#             if table_created_or_exists is False and TABLE_EXISTS_ACTION == 'skip':
#                 logging.info(f"  Skipping data transfer for existing table '{table_name}'.")
#                 continue # Move to the next table if skipping

#             # Transfer data
#             transfer_data(access_cursor, pg_cursor, table_name, columns)

#             # Commit successful table transfer
#             pg_conn.commit()
#             logging.info(f"--- Successfully processed table: {table_name} ---")

#         except Exception as e:
#             logging.error(f"!!! Failed to process table '{table_name}': {e} !!!")
#             pg_conn.rollback() # Rollback any partial changes for this table

#     logging.info("=== Migration script finished ===")

# except pyodbc.Error as e:
#     sqlstate = e.args[0]
#     if sqlstate == 'IM002':
#          logging.error("ODBC Driver Error: Data source name not found, and no default driver specified.")
#          logging.error("Ensure the MS Access ODBC driver is installed (correct 32/64 bit version) and the driver name in the script is correct.")
#          logging.error(f"Attempted driver name: {ACCESS_DRIVER}")
#     elif sqlstate == '08001':
#          logging.error(f"ODBC Connection Error: Unable to connect to Access database '{ACCESS_DB_FILE}'. Check file path and permissions.")
#     else:
#         logging.error(f"Database connection error (pyodbc): {e}")
#     sys.exit(1)
# except psycopg.Error as e:
#     logging.error(f"PostgreSQL connection or execution error: {e}")
#     sys.exit(1)
# except Exception as e:
#     logging.error(f"An unexpected error occurred: {e}")
#     sys.exit(1)
# finally:
#     if access_cursor:
#         access_cursor.close()
#     if access_conn:
#         access_conn.close()
#         logging.info("MS Access connection closed.")
#     if pg_cursor:
#         pg_cursor.close()
#     if pg_conn:
#         # If an error occurred mid-transaction, rollback before closing
#         try:
#             # Check connection status if possible (varies by psycopg version)
#              if pg_conn.status == psycopg.extensions.STATUS_IN_TRANSACTION:
#                  pg_conn.rollback()
#                  logging.warning("Rolled back final PostgreSQL transaction due to errors.")
#         except AttributeError: # Older psycopg might not have status attribute easily checkable like this
#              pass # Cannot reliably check status, close anyway
#         except Exception as final_e:
#              logging.error(f"Error during final rollback/close: {final_e}")

#         pg_conn.close()
#         logging.info("PostgreSQL connection closed.")