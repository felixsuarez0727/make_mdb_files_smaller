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

### Install `mdb-tools` on Ubuntu WSL:
```bash
sudo apt update
sudo apt install mdbtools
```

---
## Project Installation
### 1. Clone the Repository
```bash
git clone https://github.com/youruser/your-repo.git
cd your-repo
```

### 2. Create Virtual Environment and Install Dependencies
In **Windows (PowerShell):**
```bash
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

In **Ubuntu WSL:**
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

---
## Main Files
- `main.py`: Main script to be run on Windows.
- `generate_csv_files.py`: WSL script to extract and filter data.
- `requirements.txt`: Required packages for both environments.

---
## Execution
### 1. Run Main Script (Windows):
```bash
python main.py
```
This will:
- Call the WSL script to generate filtered CSVs.
- Create a new `.mdb` file with filtered data.

### 2. View Output
The resulting `.mdb` file will be saved in the project directory.

---
## `requirements.txt` File
Create it at the project root with:
```bash
touch requirements.txt
```
Add:
```
pandas==2.1.0
comtypes==1.2.0
```
In **Ubuntu WSL**, also install:
```bash
pip install mdbtools-py
```

---
## Important Notes
- **Why Windows Execution:** `.mdb` files can only be created on Windows using `comtypes` and `Access.Application`. Linux cannot create `.mdb` files.
- **Why WSL for Extraction:** `mdb-tools` is only available on Linux, so extraction is performed via Ubuntu on WSL.

---
## Author
- Name: Your Name
- Contact: youremail@example.com

---
## License
This project is licensed under the MIT License. See `LICENSE` for details.

