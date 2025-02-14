import subprocess
import pandas as pd
import io
import os
from datetime import datetime

def list_mdb_tables(mdb_path):
    tables = subprocess.check_output(["mdb-tables", "-1", mdb_path]).decode().split("\n")
    return [t.strip() for t in tables if t.strip()]

def read_mdb_table(mdb_path, table_name):
    output = subprocess.check_output(["mdb-export", mdb_path, table_name]).decode()
    return pd.read_csv(io.StringIO(output))

def filter_main_table(df, time_limit):
    df["TimeStamp"] = pd.to_datetime(df["TimeStamp"], format="%H:%M:%S", errors="coerce").dt.time  
    time_limit = datetime.strptime(time_limit, "%H:%M:%S").time()
    df_filtered = df[df["TimeStamp"].isna() | (df["TimeStamp"] <= time_limit)]
    page_counts = df_filtered["PageNum"].value_counts()
    valid_pages = page_counts[page_counts > 1].index.tolist()
    return df_filtered[df_filtered["PageNum"].isin(valid_pages)]

def export_tables_to_csv(mdb_path, output_folder_csv, time_limit):
    if not os.path.exists(output_folder_csv):
        os.makedirs(output_folder_csv)
    
    tables = list_mdb_tables(mdb_path)
    print(f"Found {len(tables)} tables in the database.")
    
    for table in tables:
        print(f"Exporting table: {table}")
        df = read_mdb_table(mdb_path, table)
        output_file = os.path.join(output_folder_csv, f"{table}.csv")
        df.to_csv(output_file, index=False)
        print(f"Saved {table} to {output_file}")
        
        if table == "Main":
            filtered_df = filter_main_table(df, time_limit)
            filtered_output_file = os.path.join(output_folder_csv, "Main_filtered.csv")
            filtered_df.to_csv(filtered_output_file, index=False)
            print(f"Saved filtered Main table to {filtered_output_file}")

if __name__ == "__main__":
    mdb_file = "/path/to/your/mdb_file.mdb" #For example: "mnt/c/User1/files/MK091624.mdb"
    current_directory = f"{os.getcwd()}/output_csv"
    time_limit="00:40:00"
    export_tables_to_csv(mdb_file, current_directory, time_limit)