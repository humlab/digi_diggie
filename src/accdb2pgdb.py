import sqlglot

def convert_access_to_postgresql(access_sql: str):
    """
    Convert SQL from Microsoft Access to PostgreSQL format.
    
    Args:
        access_sql (str): SQL query in Microsoft Access format.
    
    Returns:
        str: SQL query in PostgreSQL format.
    """
    # Convert Access SQL to PostgreSQL
    pg_sql = sqlglot.transpile(access_sql, read="msaccess", write="postgresql")[0]
    
    # Return the converted SQL
    return pg_sql

if __name__ == "__main__":
    
    # Convert to PostgreSQL
    access_sql = open("schema_ddl.sql", encoding="utf-8").read()
    # Read the Access SQL file
    pg_sql = convert_access_to_postgresql(access_sql)

    with open("converted_sql.sql", "w", encoding="utf-8") as f:
        f.write(pg_sql)
    
    # Print the converted SQL
    print("Converted PostgreSQL SQL:")
    print(pg_sql)
