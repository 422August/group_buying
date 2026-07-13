# group_buying

A set of **group-buying helper tools** built with **Python + HTML**:

1. **Order Sheet Generator** (`表單.py`)
2. **Order Message Generator** (`訊息.py`)
3. **Quantity Statistics Tool** (`統計.py`, `數量統計.html`)

---

## Overview

This repository contains tools to streamline common group-buy workflows:

- Parse order comments into structured tables
- Generate per-customer order messages from Excel
- Count order quantities from raw comment text

---

## Features

### 1) Order Sheet Generator (`表單.py`)
- Parses Facebook group-buy comments
- Supports automatic accumulation for repeated customer entries
- Handles mixed formats like `Name + number` and text+quantity patterns
- Displays merged order data in a GUI table (Tkinter Treeview)
- Exports results to Excel (`.xlsx`)

### 2) Order Message Generator (`訊息.py`)
- Loads order data from Excel (`.xlsx`, `.xls`)
- Generates personalized order messages per customer
- Search by customer name
- Sort customers by product ordering status
- Copy single message to clipboard
- Edit and save message text in-app
- Export all messages to a `.txt` file

### 3) Quantity Statistics Tool (`統計.py` / `數量統計.html`)
- **Product mode**: count by product code (e.g., `A+1`, `B*2`)
- **Total mode**: sum all `+n` / `*n` values
- Supports full-width/half-width normalization (NFKC)
- Quick copy of output result

---

## Project Structure

```text
group_buying/
├── README.md
├── requirements.txt
├── 表單.py
├── 訊息.py
├── 統計.py
└── 數量統計.html
```

---

## Requirements

- Python 3.10+ (recommended)
- pip

---

## Installation

```bash
git clone https://github.com/August980422/group_buying.git
cd group_buying
python -m venv venv
```

Activate virtual environment:

- macOS / Linux
  ```bash
  source venv/bin/activate
  ```
- Windows
  ```bash
  venv\Scripts\activate
  ```

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Dependencies (`requirements.txt`)

- `customtkinter==5.2.2`
- `openpyxl==3.1.5`
- `pandas==2.2.3`
- `pyperclip==1.9.0`

---

## Usage

### A. Run Order Sheet Generator
```bash
python 表單.py
```

### B. Run Order Message Generator
```bash
python 訊息.py
```

### C. Run Quantity Statistics (Desktop)
```bash
python 統計.py
```

### D. Run Quantity Statistics (Web)
Open this file directly in your browser:

```text
數量統計.html
```

---

## Input Examples (Statistics Tool)

- Product counting: `A+1`, `B 2`, `C*3`
- Total counting: sums numbers after `+` or `*`

---

## Notes

- Some filenames are in Chinese (e.g., `表單.py`, `訊息.py`). Use a UTF-8 compatible environment.
- Excel column structure affects message-generation results; keep a consistent sheet format.
- `訊息.py` uses system clipboard via `pyperclip`.

---

## License

MIT

---

## Author

GitHub: [@August980422](https://github.com/August980422)
