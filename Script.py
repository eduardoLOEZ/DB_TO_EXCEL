import pymysql
import pandas as pd
from openpyxl import Workbook
import os


#TRAER TODAS LA TABLAS DE LA DB
def get_table(cursor):
    cursor.execute("SHOW TABLES")
    return [table[0] for table in cursor.fetchall()]


#VERIFICAR QUE EXISTA LA TABLA QUE INGRESA EL USUARIO
#LA COMPARAMOS CON LA LISTA DE TABLAS QUE IMPORTAMOS 
def check_table_exists(table_name, existing_tables):
    return table_name in existing_tables  #true or false


#EXPORTAR LA QUERY SQL A DATAFRAME Y DE DF A EXCEL
def export_db_to_excel(table_name, conn):
    sql_query = f"SELECT * FROM {table_name}"
    df = pd.read_sql_query(sql_query, conn)

    workbook = Workbook()
    sheet = workbook.active
    
    for index, row in df.iterrows():
        sheet.append(row.tolist())

    documents_folder  = os.path.expanduser('~\\Documents')
     #GUARDAR EN LOS DOCS DE WINDOWS
    excel_file_path = os.path.join(documents_folder, f'{table_name}.xlsx')
    workbook.save(excel_file_path)




def main()-> None:
    #CONEXION A DB
    db_config= {
        'host': '127.0.0.1',
        "port": 3306,
        'user': 'root',
        'password': 'amarillo200',
        'database': 'prueba_dbtoexcel'
    }

    nombre_tabla = input("introduce el nombre de la tabla que deseas exportar: ")

    conn = pymysql.connect(**db_config)
    cursor = conn.cursor()

    existing_tables = get_table(cursor)

    #SI DA UN FALSE, DARA EL MSG ERROR
    if not check_table_exists(nombre_tabla, existing_tables):
        print(f"Error: La tabla '{nombre_tabla}' no existe en la base de datos.")
    else:
        export_db_to_excel(nombre_tabla, conn)
        print("tabla de la DB guardada exitosamente!!")

    
    cursor.close()
    conn.close()

if __name__== "__main__":
    main()

