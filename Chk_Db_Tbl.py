import pymysql

host='sscdb01.convoyops.com'
port=3306
database='sos_logginsys2 '
user='sos-read'
password='sos-read-210930'


def get_mysql_tables(host, user, password, database):
    try:
        # Connect to MySQL server
        connection = pymysql.connect(
            host=host,
            port=3306,
            user=user,
            password=password,
            database=database
        )
        
        cursor = connection.cursor()

        # Query to get the list of tables with schema
        query = """
        SELECT 
            TABLE_SCHEMA, 
            TABLE_NAME
        FROM 
            INFORMATION_SCHEMA.TABLES
        WHERE 
            TABLE_SCHEMA = %s;
        """
        cursor.execute(query, (database,))
        
        tables = cursor.fetchall()

        # Display tables
        for table in tables:
            print(f"Schema: {table[0]}, Table: {table[1]}")

    except pymysql.MySQLError as err:
        print(f"Error: {err}")
    finally:
        if connection:
            cursor.close()
            connection.close()

# Example usage:
get_mysql_tables(host=host, user=user, password=password, database=database)
