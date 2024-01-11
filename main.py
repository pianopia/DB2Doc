import os
import mysql.connector
import pandas as pd
from dotenv import load_dotenv
import openpyxl

# .envファイルから環境変数を読み込む
load_dotenv()

# MySQLの接続情報を環境変数から取得
db_config = {
    'host': os.getenv('MYSQL_HOST'),
    'user': os.getenv('MYSQL_USER'),
    'password': os.getenv('MYSQL_PASSWORD'),
    'database': os.getenv('MYSQL_DATABASE'),
}

# MySQLに接続
conn = mysql.connector.connect(**db_config)
cursor = conn.cursor()

# テーブル一覧を取得
cursor.execute("SHOW TABLES;")
tables = cursor.fetchall()

# Excelファイルを作成
excel_writer = pd.ExcelWriter('db_design.xlsx', engine='openpyxl')

# 各テーブルの情報をExcelに書き込む
for table in tables:
    table_name = table[0]
    
    # テーブルのカラム情報を取得
    cursor.execute(f"DESCRIBE {table_name};")
    columns = cursor.fetchall()
    
    # カラム情報をDataFrameに変換
    column_df = pd.DataFrame(columns, columns=['Field', 'Type', 'Null', 'Key', 'Default', 'Extra'])
    
    # DataFrameをExcelに書き込み
    column_df.to_excel(excel_writer, sheet_name=table_name, index=False)

# Excelファイルを保存
# excel_writer.save()

# Close the ExcelWriter object
excel_writer.close()

# MySQLとの接続を閉じる
cursor.close()
conn.close()

