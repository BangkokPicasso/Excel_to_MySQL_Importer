import pandas as pd
import os.path
import pymysql

def get_xlsx_data_and_to_df(File_Path, sheet_names_list):
    if os.path.isfile(File_Path):
        print("  File Found  " + File_Path)
        df = pd.read_excel(File_Path, sheet_name=sheet_names_list)
        return df
    else:
        print("404 NOT FOUND  " + File_Path)   
        return       

def execute_sql(SQL, data):

    cnx = pymysql.connect(
        host="127.0.0.1", 
        user="UserName", 
        password="password", 
        db="DBName", 
        charset="utf8"
    )
    cursor = cnx.cursor()    
    if data:
        cursor.executemany(SQL, data) 
    else:
        cursor.execute(SQL) 
    cnx.commit()
    cursor.close()
    cnx.close()

    return    

def main():

     # Dictionary mapping Excel sheet names to MySQL table names
    table_name_dict = {
        # "xlsx SheetName": "SchemaName.TableName"
        "SPC": "Schema1.settings_spc",
        "sales": "Schema2.sales",
    }

    # Path to the Excel file
    File_Path = "C:\\Users\\user1\\Desktop\\data.xlsx"

    sheet_names_list = list(table_name_dict.keys())
    # Read data from the Excel file into a dictionary of DataFrames
    df = get_xlsx_data_and_to_df(File_Path, sheet_names_list)

    for sheet_name in df:
        data = [tuple(row) for row in df[sheet_name].values.tolist()]
        # Generate placeholders for SQL insertion
        values_s = "%s," * len(df[sheet_name].columns)
        columns_str = ",".join([f"`{col}`" for col in df[sheet_name].columns])
        
        # SQL insert command
        # assuming xlsx column names == table column names
        sql_insert = f"""
            INSERT INTO {table_name_dict[sheet_name]}
            ({columns_str})
            VALUES
            ({values_s[:-1]})    
        """
        print(f"Inserting {len(data)} rows into {table_name_dict[sheet_name]}...")
        execute_sql(sql_insert, data)

    return



if __name__ == '__main__':
    main()