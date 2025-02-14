import os
import time
import pandas as pd
from comtypes.client import CreateObject
from comtypes.gen import Access
import subprocess

def get_current_directory_paths():
    """Gets current directory paths for both Windows and WSL formats."""
    current_directory = os.getcwd()
    drive, path = os.path.splitdrive(current_directory)
    drive_letter = drive.replace(':', '').lower()
    current_directory_linux = f"/mnt/{drive_letter}{path.replace('\\', '/')}"
    
    return current_directory, current_directory_linux

def run_wsl_script(current_directory_linux):
    """Runs the script in Ubuntu to generate CSV files."""
    print("Running script in WSL to generate CSV files...")
    subprocess.run(["wsl", "python3", f"{current_directory_linux}/generate_csv_files.py"], check=True)
    print("Script executed successfully.")

def get_available_mdb_filename(base_name="my_database.mdb"):
    """Finds an available .mdb filename by incrementing a counter."""
    current_directory = f"{os.getcwd()}"
    file_path = os.path.join(current_directory, base_name)
    
    if not os.path.exists(file_path):
        return file_path
    
    name, ext = os.path.splitext(base_name)
    counter = 1
    while True:
        new_filename = f"{name}_{counter}{ext}"
        new_file_path = os.path.join(current_directory, new_filename)
        if not os.path.exists(new_file_path):
            return new_file_path
        counter += 1

def create_new_mdb_file(filename="my_database.mdb"):
    """Creates a new .mdb file inside the /newMdbFiles directory, creating it if necessary."""
    
    folder_path = os.path.join(os.getcwd(), "newMdbFiles")
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    file_path = get_available_mdb_filename(os.path.join(folder_path, filename))
    
    access = CreateObject('Access.Application')
    DBEngine = access.DBEngine
    db = DBEngine.CreateDatabase(file_path, Access.DB_LANG_GENERAL)
    db.Close()
    
    print(f"MDB file created at: {file_path}")
    return file_path

def insert_csv_into_mdb(mdb_path, csv_folder):
    """Reads CSV files and inserts them into the MDB."""
    access = CreateObject('Access.Application')
    DBEngine = access.DBEngine
    db = DBEngine.OpenDatabase(mdb_path)
    
    csv_files = {
        "Info": "Info.csv",
        "MPEGS": "MPEGS.csv",
        "Main": "Main_filtered.csv"
    }
    
    for table_name, csv_file in csv_files.items():
        csv_path = os.path.join(csv_folder, csv_file)
        
        if not os.path.exists(csv_path):
            print(f"Warning: {csv_path} not found, skipping...")
            continue
        
        df = pd.read_csv(csv_path)
        columns = ", ".join([f"[{col}] TEXT" for col in df.columns])
        create_table_sql = f"CREATE TABLE {table_name} ({columns})"
        db.Execute(create_table_sql)
        
        for _, row in df.iterrows():
            values = ", ".join([f"'{str(val).replace("'", "''")}'" for val in row])
            insert_sql = f"INSERT INTO {table_name} VALUES ({values})"
            db.Execute(insert_sql)
        
        print(f"Table {table_name} created and inserted into {mdb_path}")
    
    db.Close()
    print("Process completed.")

if __name__ == "__main__":
    current_directory, current_directory_linux = get_current_directory_paths()
    output_csv_folder = os.path.join(current_directory, "output_csv")

    run_wsl_script(current_directory_linux)
    time.sleep(2)
    
    new_mdb_path = create_new_mdb_file("my_database.mdb")
    insert_csv_into_mdb(new_mdb_path, output_csv_folder)
