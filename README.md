# Excel_to_MySQL_Importer

This project provides a Python script to read data from an Excel file and insert it into specified MySQL database tables. The script uses `pandas` for reading Excel files and `pymysql` for interacting with the MySQL database. This is useful for data migration tasks or batch data insertion from Excel files to MySQL databases.

## Features

- Reads data from specified sheets in an Excel file using `pandas`.
- Connects to a MySQL database using `pymysql`.
- Inserts the data into corresponding MySQL tables.
- Supports bulk insertion for efficiency.

## How to Use

1. **Install Required Libraries**: Ensure you have `pymysql` and `pandas` installed.
   ```sh
   pip install pymysql pandas
   ```
2. **Configure Database Connection**: Update the connection parameters (host, user, password, db) in the execute_sql function.

3. **Map Excel Sheets to MySQL Tables**: Update the table_name_dict dictionary in the main function to map Excel sheet names to MySQL table names.
   
4. **Specify the Path to Your Excel File**: Update the File_Path variable in the main function to point to your Excel file location.
   ```sh
   File_Path = "C:\\Path\\To\\Your\\data.xlsx"
   ```
5. **Run the Script**: Execute the script to read data from the Excel file and insert it into the MySQL tables.
   ```sh
   python Excel_to_MySQL_Importer.py
   ```
## Example
In this script, the Excel file data.xlsx located on the desktop is read. The data from the sheets SPC and sales are inserted into the MySQL tables Schema1.settings_spc and Schema2.sales, respectively.

## Related Repositories
* [export_mysql_data_to_xlsx](https://github.com/BangkokPicasso/export_mysql_data_to_xlsx): Export MySQL Data to Excel
* [df_to_mysql_table_statement_generator](https://github.com/BangkokPicasso/df_to_mysql_table_statement_generator): Converting a Pandas DataFrame into a MySQL CREATE TABLE statement
  
## Prerequisites
* Python 3.x
* MySQL server
* Necessary Python libraries: pymysql, pandas
  
## Notes
* Ensure your MySQL server is running and accessible.
* Update the SQL insert command if the Excel column names differ from the MySQL table column names.
