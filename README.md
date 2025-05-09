# Excel File Splitter

**Excel File Splitter** is a Python automation tool that divides a large Excel spreadsheet (e.g., 150,000+ rows) into multiple smaller `.xlsx` files, each with a styled header and preserved layout.

This tool is useful for handling large datasets, reducing file size, and improving manageability for clients or end-users.

---

## 🚀 Features

- Handles large `.xlsx` files with complex headers
- Preserves:
  - Merged header cells
  - Font and cell styles
  - Column widths and row heights
- Filters out invalid rows based on 12-digit ID validation
- Supports CLI usage with flexible configuration
- Includes dry-run mode to preview output structure without saving files

---

## 🛠 Requirements

- Python 3.9+
- `openpyxl`

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## 📦 Usage

```bash
python excel_splitter.py -i path/to/large_file.xlsx -p 10000 -o output_folder
```

### Arguments:

| Argument           | Description |
|--------------------|-------------|
| `-i`, `--input`     | Path to the input `.xlsx` file (required) |
| `-p`, `--per-file`  | Number of rows per output file **(mutually exclusive with `--num-files`)** |
| `-n`, `--num-files` | Total number of output files to create |
| `-o`, `--output-dir`| Directory to save the output files (default: current directory) |
| `-s`, `--sheet-name`| Name of the worksheet to split (default: second sheet in workbook) |
| `--prefix`          | Prefix for output filenames (default: `split`) |
| `--dry-run`         | Preview how many files would be generated, without saving anything |

---

## ✅ Example

Split a file into 10,000-row chunks:

```bash
python excel_splitter.py -i ./products.xlsx -p 10000 -o ./chunks
```

Simulate and report how the file will be split:

```bash
python excel_splitter.py -i ./products.xlsx -n 15 --dry-run
```

---

## 📂 Output

Creates output files like:

```
split_01.xlsx
split_02.xlsx
...
split_N.xlsx
```

Each file includes the original header, layout, and a subset of the data rows.

---

## 📄 License

This project is licensed under the MIT License.
