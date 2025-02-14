# README: .MDB File Reduction Project

## Description
This project generates a smaller version of a `.mdb` (Microsoft Access Database) file by filtering the `Main` table based on a time limit. The code runs on **Windows**, but uses **Ubuntu WSL** for data extraction via `mdb-tools`.

## Prerequisites
- **Windows 10/11** with WSL enabled.
- **Ubuntu** installed on WSL.
- **Python 3.10+** installed on both Windows and WSL.

### Verify WSL Installation
1. Open **PowerShell** as administrator.
2. Run:
   ```bash
   wsl -l -v
   ```
   - Confirm `Ubuntu` appears and is `Running`.
   - If not installed, use:
     ```bash
     wsl --install -d Ubuntu
     ```
     Then, set Ubuntu as your default distribution of WSL by running:
     ```bash
     wsl --set-default Ubuntu
     ```

### Install `mdb-tools` on Ubuntu WSL:
```bash
sudo apt update
sudo apt install mdbtools
```

---
## Project Installation
### 1. Clone the Repository
```bash
git clone https://github.com/felixsuarez0727/make_mdb_files_smaller.git

cd mdb_small
```

### 2. Create Virtual Environment and Install Dependencies
In **Windows (PowerShell):**
```bash
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

---
## Main Files
- `main.py`: Main script to be run on Windows.
- `generate_csv_files.py`: WSL script to extract and filter data.
- `requirements.txt`: Required packages for both environments.

---
## Execution
### 1. Set your .mdb file path:
On the script `generate_csv_files.py`, in the line 44, set the path of the .mdb file that you want to get smaller:
```bash
mdb_file = "/path/to/your/mdb_file.mdb" 
```

### 2. Run Main Script (Windows):
```bash
python main.py
```
This will:
- Call the WSL script to generate filtered CSVs.
- Create a new `.mdb` file with filtered data.

### 2. View Output
The resulting `.mdb` file will be saved in the project directory, within the `nevMdbFilesFolder` under the name of `my_database.mdb`. If you run the `main.py` script multiple times, then you will get different files called `my_database.mdb`, `my_database_1.mdb`, and so on.

---
## Important Notes
- **Why Windows Execution:** `.mdb` files can only be created on Windows using `comtypes` and `Access.Application`. Linux cannot create `.mdb` files.
- **Why WSL for Extraction:** `mdb-tools` is only available on Linux, so extraction is performed via Ubuntu on WSL.

---
## Author
- Name: Felix Suarez
- Contact: felix@depoiq.com

---


