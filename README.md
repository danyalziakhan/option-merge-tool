# Option Merge Tool

A Python utility to join two columns from an Excel file into a third column using a custom separator.  
It also supports formatting based on a template Excel file, NaN row removal, and duplicate row filtering.

Includes both a **GUI** (using [DearPyGui](https://github.com/hoffstadt/DearPyGui)) and a **CLI** interface for flexibility.

---

## ✨ Features

- 📄 Read Excel `.xlsx` input files with customizable column mappings
- 🔗 Join values from two text columns using a separator (e.g. comma)
- 🧹 Drop rows with NaN values in a specified column
- 📌 Remove duplicate rows based on specified column combinations
- 📋 Format output using a provided Excel template file
- 🖥️ GUI mode for user-friendly interaction
- ⚙️ CLI mode for automation and scripting

---

## 📦 Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/danyalziakhan/option-merge-tool.git
   cd option-merge-tool
   ```

2. Install [uv](https://github.com/astral-sh/uv) package manager:

   ```bash
   # On macOS and Linux.
   curl -LsSf https://astral.sh/uv/install.sh | sh
   ```

   ```powershell
   # On Windows (PowerShell).
   irm https://astral.sh/uv/install.ps1 | iex
   ```

   Or install via pip:

   ```bash
   pip install uv
   ```

3. Create virtual environment and sync dependencies:

   ```bash
   uv venv
   uv pip install -r pyproject.toml
   uv sync
   ```

---

## 🚀 Usage

### ✅ CLI Mode

Use the CLI version for scripting or automated pipelines:

```bash
python run.py ^
  --input_file INPUT_FILE.xlsx ^
  --template_file TEMPLATE_FILE_DB.xlsx ^
  --first_column "원본 상품명" ^
  --second_column "원가\n[필수]" ^
  --join_by "," ^
  --output_column "옵션상세명칭(1)\n[사방넷]" ^
  --column_to_dropna "원본 상품명" ^
  --columns_to_drop_dulicates "원본 상품명,옵션상세명칭(1),물류처ID,모델NO"
```

Or run on Windows:

```bash
run.bat
```

---

### 🖼️ GUI Mode

Launch a GUI with file pickers and form inputs for interactive use:

```bash
python run.py --gui ^
  --input_file INPUT_FILE.xlsx ^
  --template_file TEMPLATE_FILE_DB.xlsx ^
  --first_column "원본 상품명" ^
  --second_column "원가\n[필수]" ^
  --join_by "," ^
  --output_column "옵션상세명칭(1)\n[사방넷]" ^
  --column_to_dropna "원본 상품명" ^
  --columns_to_drop_dulicates "원본 상품명,옵션상세명칭(1),물류처ID,모델NO"
```

Or run on Windows using:

```bash
run_gui.bat
```

---

## 🗂 Output

- Resulting Excel files saved in the `output/` directory
- Logs are saved in `logs/` directory by date

---

## 🛠 Requirements

- Python 3.10 or 3.11
- Dependencies defined in `pyproject.toml`

---

## 📁 Project Structure

```
option-merge-tool/
│
├── run.py               # Main program entry point
├── run.bat              # Windows CLI launcher
├── run_gui.bat          # Windows GUI launcher
├── pyproject.toml
├── README.md
│
└── option_merge_tool/
    ├── excel.py
    ├── gui.py
    ├── log.py
    ├── merge.py
    ├── non_gui.py
    ├── settings.py

```

---

## 🤝 License

MIT License. Free to use, modify, and distribute.
