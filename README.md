# Quotation Index Generator — Tổng hợp Báo giá

A lightweight, no-database tool that scans a construction price-estimation project
folder and automatically compiles a formatted Excel summary (**"DANH MỤC BÁO GIÁ"**)
from the **names** of the folders and files inside it. Built for the price-estimation
team at Cubic Architects.

The program reads only folder and file **names** — it never opens the contents of the
PDF, Excel, Word, or image files. That makes it fast, fully deterministic, and
privacy-preserving: the large quotation files (often several GB) never need to be
opened, copied, or moved.

---

## How it works

1. You name the quotation files in each category folder following a simple convention.
2. You run the tool (double-click a `.bat`, or call it from the command line).
3. The tool walks the project folder, parses the names, and writes a two-sheet Excel
   workbook: a print-ready report and a data/QA sheet.

The naming convention *is* the data schema — there is no separate database.

---

## Repository contents

| File | Purpose |
|------|---------|
| `bao_gia_index.py` | The engine: parses folder/file names and builds the Excel workbook. |
| `Chay_Tong_hop_bao_gia.bat` | Windows one-click runner (finds Python, installs the dependency, runs the tool, opens the result). |
| `cau_hinh.txt` | Plain-text configuration (self-documenting, in Vietnamese). Optional — the tool falls back to built-in defaults if it is missing. |
| `Huong_dan_dat_ten.docx` | The end-user guide for the team, in Vietnamese, written for non-technical staff. |
| `Tong_hop_bao_gia.xlsx` | A sample of the generated output. |
| `README.md` | This file. |

---

## Requirements

- **Python 3.8+**
- **openpyxl** (`pip install openpyxl`) — the only third-party dependency.
- Windows is recommended for the one-click `.bat`; the Python script itself runs on
  any OS.

---

## Setup & usage

### Option A — one-click (Windows)

1. Copy `bao_gia_index.py`, `Chay_Tong_hop_bao_gia.bat`, and (optionally)
   `cau_hinh.txt` into the **project folder** — the folder that contains the numbered
   category subfolders (`1. …`, `2. …`, … `12. …`).
2. Double-click `Chay_Tong_hop_bao_gia.bat`.
3. The result, `Tong_hop_bao_gia.xlsx`, is written next to it and opens automatically.

The first run on a new machine needs Python installed (with "Add Python to PATH"
ticked) and an internet connection so the runner can install `openpyxl`. Subsequent
runs need neither.

### Option B — command line

```bash
python bao_gia_index.py "<path to project folder>" "<output file.xlsx>"
```

Example:

```bash
python bao_gia_index.py "D:/Projects/03. BAO GIA" "Tong_hop_bao_gia.xlsx"
```

Re-running always rebuilds the report from the current folder state, so you can add,
remove, or rename files and run again at any time.

---

## Naming conventions

### Category folders

A folder is treated as a category **only** when its name starts with a number followed
by a separator (`.`, `-`, `_`, or `)`):

```
3. Ket cau
4 - Kien truc
10_PCCC
```

- The leading **number** sets the order; the text after it is for human reading only
  (the display name comes from the configuration).
- Numbers **1–12** map to the twelve configured categories.
- Numbers **outside 1–12** (e.g. `00.`, `13.`, `14.`) are still read and appended as
  extra categories at the end, using the folder's own name.
- A leading number immediately followed by another digit is treated as a **date**, not
  a category (e.g. `4.6.2026 Summary` is ignored, never read as category 4).
- Any folder **without** a leading number — including names beginning with `@` or `!`,
  or plain text — is **ignored entirely**. Reference material and large folders can sit
  in the project untouched.
- If a folder number 1–12 is missing, the report still shows the full I–XII skeleton;
  the missing one is just an empty header row.
- If two folders share the same number, both are merged into the report **and** a
  warning is written to the QA sheet (nothing is silently dropped).

### Quotation files

Files to be included are named:

```
Subcategory @ Partner @ Status.ext
```

| Part | Meaning | Required |
|------|---------|----------|
| Subcategory | The work item / material (e.g. `Chống thấm`, `Cửa gỗ`) | Yes |
| Partner | The supplier / contractor (e.g. `Công ty Dcons`) | No (may be blank) |
| Status | Selected / not selected | No |

Notes:

- The `@` character is the field separator. Spaces around it are optional.
- **Subcategory** and **Partner** are printed verbatim, so they must be typed with
  correct Vietnamese diacritics.
- A leading `Báo giá` / `Bảng giá` / `BG` on the subcategory is stripped automatically.
- `Cong ty`, `Cty`, `C.ty`, `Congty` (any case/diacritics) are normalized to `Công ty`.
- Only these file types are read: `pdf, xls, xlsx, xlsm, doc, docx, jpg, jpeg, png,
  ppt, pptx`. Everything else (`.zip`, technical files, `Thumbs.db`, …) is ignored.
- A file of a readable type that does **not** match the convention is listed in the QA
  sheet so it can be corrected.

### Status shorthand

Typed in the file name → printed in the report:

| Printed | Accepted spellings |
|---------|--------------------|
| `Chọn` | `chọn, chon, c, co, x, v, ok` |
| `Không chọn` | `không chọn, khong chon, ko chon, ko, k, không, ko dùng, k dùng` |
| *(blank)* | omit the status field |

---

## Output

A single `.xlsx` workbook with two sheets:

- **`Báo cáo`** — the client-facing report. Title block (DANH MỤC BÁO GIÁ + DỰ ÁN /
  CÔNG TRÌNH / ĐỊA ĐIỂM), a blue header row, and the indexed table. Categories use
  Roman numerals (I–XII); items within a category are numbered with live `=A+1`
  formulas. Times New Roman throughout. Print setup is baked in: **A4, portrait, fit to
  one page wide, horizontally centered, fixed margins, with the column header repeated
  on every printed page.**
- **`Dữ liệu & Kiểm tra`** — a flat data table (one row per file read), the list of
  skipped files, and warnings (duplicate folder numbers, possible duplicate
  subcategories from inconsistent spelling, etc.). This sheet can be deleted if not
  needed.

---

## Configuration (`cau_hinh.txt`)

A plain-text file the team can edit in Notepad without touching the Python code. It is
optional; if absent, the built-in defaults are used. Sections:

| Section | Controls |
|---------|----------|
| `[THONG TIN]` | Report title and the project / building / location lines (persist across runs). |
| `[DANH MUC]` | The category list: number, display name, and whether it is a special I/II folder. |
| `[CHON]` / `[KHONG CHON]` | The words accepted as "Chọn" / "Không chọn". |
| `[BO TIEN TO]` | Prefixes stripped from the start of subcategory names. |
| `[DUOI FILE]` | Which file extensions are read. |
| `[TUY CHON]` | Toggles: auto-convert "Cong ty" → "Công ty"; auto-capitalize the first letter of a subcategory. |

To adapt the tool to a project with different disciplines, edit `[DANH MUC]`.

---

## Limitations & notes

- The tool produces no standalone `.exe`; each machine needs Python installed once. (A
  PyInstaller build can be added if a zero-install executable is wanted.)
- Vietnamese diacritics in the subcategory and partner fields are the user's
  responsibility — the tool cannot guess them. Category display names are always correct
  because they come from the configuration.
- When run from a mapped/network drive, that drive must be available on the machine.

---

## License

Released under the [MIT License](LICENSE) — see the `LICENSE` file for the full text.
You are free to use, copy, modify, and distribute this software; it is provided "as is",
without warranty of any kind.
