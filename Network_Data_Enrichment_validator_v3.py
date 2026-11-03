"""
Branch Device / Network Interface Details - Excel Validator / Corrector
------------------------------------------------------------------------
On running, asks the user to select an Excel file and corrects the values
as per the "Description" column of the format sheet:

    Column Name        Description
    LinkUsage          First letter Capital, rest small
    State              All Capital
    Country            First letter Capital, rest small
    Circle             All Capital
    Region             First letter Capital, rest small
    BranchCode         5 digit (leading zeros added)
    Environment        First letter Capital, rest small
    CountryFO          First letter Capital, rest small
    ISP                All Capital
    Bandwidth          Same as given (16 Mbps)
    All other columns  N/A (left untouched)

Extra rules:
  * Region / State containing any digit  -> cell RED (BranchCode is never red)
  * Device       "Router" / "router"     -> "IP router"
  * Environment  word "branch" coming after the first word gets a capital
                 letter, e.g. "Critical branch" -> "Critical Branch"
  * BranchCode   only made 5 digit by adding leading zeros (no colour)

Two sheet layouts are supported (detected automatically):
  1) Vertical  : headings "Column Name | Value | Description" in a row,
                 one field per row (rules are read from the Description column)
  2) Horizontal: field names as column headings in one row (row 1), one
                 record per row below it

Output: a new file "<original name>_corrected.xlsx" saved in the SAME
        folder as the input file (original is untouched).
        Corrected cells are highlighted YELLOW (except BranchCode), cells
        with numbers in Region / State are highlighted RED, and a change
        log is printed.

Requirements:  pip install openpyxl
"""

import os
import re
import sys
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter


def norm(text):
    """Normalise a header for matching (case/space insensitive)."""
    return re.sub(r"\s+", "", str(text)).lower() if text is not None else ""


# --------------------------------------------------------------------------
# Rules (taken from the Description column)
#   "capfirst"    -> First letter Capital, rest small
#   "upper"        -> All Capital
#   "code5"        -> 5 digit (leading zeros added)
#   "bandwidth"    -> Same as given, e.g. "16 Mbps"
#   "device"       -> "Router"/"router" becomes "IP router"
#   N/A columns are not listed (left untouched)
# --------------------------------------------------------------------------
RULES = {
    "LinkUsage": "capfirst",
    "State": "upper",
    "Country": "capfirst",
    "Circle": "upper",
    "Region": "capfirst",
    "BranchCode": "code5",
    "Environment": "capfirst",
    "CountryFO": "capfirst",
    "ISP": "upper",
    "Bandwidth": "bandwidth",
    "Device": "device",
}

# Words that must start with a capital letter when they come AFTER the first
# word of the value (the first word always gets a capital letter anyway).
#   { column heading : [words] }     e.g. "Critical branch" -> "Critical Branch"
WORD_CAPITALIZE = {
    "Environment": ["branch"],
}
WORD_CAPITALIZE_NORM = {norm(k): v for k, v in WORD_CAPITALIZE.items()}

# Columns that must be highlighted RED if they contain any number.
# They are found BY HEADING NAME. Only if a heading is missing from a
# horizontal sheet, the column letter below is used (and only when that
# column does not already carry another known heading, e.g. BranchCode).
NUMBER_CHECK_COLUMNS = ["Region", "State"]
NUMBER_CHECK_LETTERS = {"Region": "G", "State": "K"}

# If a heading is not found in a horizontal sheet, these letters are used
FALLBACK_LETTERS = {
    "Device": "N",
    "Environment": "O",
}

# All known headings (used to detect a horizontal sheet / header row)
KNOWN_HEADERS = [
    "WANLinkIP", "LinkUsage", "DeviceIP", "State", "Country", "City",
    "Circle", "BranchDataRegion", "BranchEmailId", "Region", "BranchCode",
    "BranchName", "Environment", "Network", "Module", "CIName", "CountryFO",
    "ISP", "CircuitId", "Bandwidth", "DOC", "Link-Maintenance", "LinkMntST",
    "LinkMntET", "SRCBranchCode", "action", "Device",
]

HIGHLIGHT = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # yellow = corrected
RED_FILL = PatternFill(fill_type="solid", fgColor="FFFF0000", bgColor="FFFF0000")     # red = number found

RULES_NORM = {norm(k): v for k, v in RULES.items()}
KNOWN_NORM = {norm(h) for h in KNOWN_HEADERS}
NUMBER_CHECK_NORM = {norm(h) for h in NUMBER_CHECK_COLUMNS}


# ----------------------------- value fixers -------------------------------
def fix_bandwidth(value):
    """'16mbps', '16 MBPS', 16 -> '16 Mbps'"""
    text = str(value).strip()
    m = re.match(r"^(\d+(?:\.\d+)?)\s*([a-zA-Z]*)$", text)
    if not m:
        return text  # unknown pattern, leave as is
    number, unit = m.group(1), m.group(2).lower()
    if number.endswith(".0"):
        number = number[:-2]
    unit_map = {"": "Mbps", "m": "Mbps", "mb": "Mbps", "mbps": "Mbps",
                "k": "Kbps", "kb": "Kbps", "kbps": "Kbps",
                "g": "Gbps", "gb": "Gbps", "gbps": "Gbps"}
    return f"{number} {unit_map.get(unit, unit.capitalize())}"


def fix_branch_code(value):
    """Make BranchCode a 5 digit text value (zero padded)."""
    text = str(value).strip()
    if text.endswith(".0"):          # numeric cell like 5559.0
        text = text[:-2]
    if text.isdigit() and len(text) < 5:
        return text.zfill(5)
    return text


def fix_device(value):
    """'Router' / 'router' (with or without spaces) -> 'IP router'"""
    text = str(value).strip()
    if text.lower() == "router":
        return "IP router"
    return value


def has_number(value):
    """True if the cell value contains any digit."""
    if value is None:
        return False
    return any(ch.isdigit() for ch in str(value))


def capitalize_words(text, col_name=""):
    """First letter Capital, rest small. Words listed in WORD_CAPITALIZE
    (e.g. 'branch' in Environment) also get a capital when not first word."""
    result = text.capitalize()
    special = WORD_CAPITALIZE_NORM.get(norm(col_name), [])
    if special:
        words = result.split(" ")
        for i in range(1, len(words)):
            if words[i].lower() in special:
                words[i] = words[i].capitalize()
        result = " ".join(words)
    return result


def apply_rule(rule, value, col_name=""):
    if value is None or str(value).strip() == "":
        return value
    text = str(value).strip()
    if rule == "capfirst":
        return capitalize_words(text, col_name)
    if rule == "upper":
        return text.upper()
    if rule == "code5":
        return fix_branch_code(value)
    if rule == "bandwidth":
        return fix_bandwidth(value)
    if rule == "device":
        return fix_device(value)
    return value


def parse_description(desc, col_name):
    """
    Convert the text of the Description column into a rule.
    Returns a rule name, "none" (N/A) or None (blank / not understood).
    """
    if desc is None:
        return None
    d = re.sub(r"\s+", " ", str(desc)).strip().lower()
    if not d:
        return None
    if d in ("n/a", "na", "-"):
        return "none"
    if re.search(r"\b5\s*digit", d):
        return "code5"
    if "all capital" in d or "all caps" in d or "uppercase" in d:
        return "upper"
    if "letter" in d and "capital" in d:      # also catches typo "Firtst letter"
        return "capfirst"
    if "same as given" in d:
        return "bandwidth" if norm(col_name) == "bandwidth" else "none"
    return None


# --------------------------- cell processing -------------------------------
def process_cell(ws, cell, col_name, rule, check_number, log):
    """Apply the rule to one cell, highlight changes / numbers. Returns 1 if red."""
    red = 0
    if rule:
        old = cell.value
        new = apply_rule(rule, old, col_name)
        if new != old and str(new) != str(old):
            cell.value = new
            if rule == "code5":
                cell.number_format = "@"   # keep as text, no colour
            else:
                cell.fill = HIGHLIGHT
            log.append((ws.title, cell.coordinate, col_name, old, new))

    # number check: Region / State -> RED (done last so red wins over yellow)
    # BranchCode is never highlighted red.
    if check_number and norm(col_name) != "branchcode" and has_number(cell.value):
        cell.fill = RED_FILL
        red = 1
        log.append((ws.title, cell.coordinate, col_name, cell.value,
                    "!! CONTAINS NUMBER - highlighted red"))
    return red


# ----------------------- layout 1: vertical sheet --------------------------
def find_vertical_header(ws, max_rows=10):
    """Find the row holding 'Column Name | Value | Description'."""
    for r in range(1, min(ws.max_row, max_rows) + 1):
        cols = {norm(c.value): c.column for c in ws[r] if c.value is not None}
        if "columnname" in cols and "value" in cols:
            return r, cols["columnname"], cols["value"], cols.get("description")
    return None


def process_vertical(ws, info, log):
    header_row, name_col, value_col, desc_col = info
    print("      Layout: vertical (Column Name | Value | Description)", flush=True)
    red_count = 0
    for r in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(row=r, column=name_col).value
        if name is None or str(name).strip() == "":
            continue
        name = str(name).strip()
        n = norm(name)
        desc = ws.cell(row=r, column=desc_col).value if desc_col else None

        parsed = parse_description(desc, name)
        if parsed == "none":
            rule = RULES_NORM.get(n) if n == "device" else None
        elif parsed:
            rule = parsed
        else:                                   # description blank / unknown
            rule = RULES_NORM.get(n)

        cell = ws.cell(row=r, column=value_col)
        red_count += process_cell(ws, cell, name, rule,
                                  n in NUMBER_CHECK_NORM, log)
    print(f"      Cells highlighted RED in '{ws.title}': {red_count}", flush=True)


# ---------------------- layout 2: horizontal sheet -------------------------
def find_header_row(ws, max_rows=10):
    """Return row number that looks like the heading row, else None."""
    for r in range(1, min(ws.max_row, max_rows) + 1):
        hits = sum(1 for c in ws[r] if norm(c.value) in KNOWN_NORM)
        if hits >= 5:
            return r
    return None


def process_horizontal(ws, header_row, log):
    print("      Layout: horizontal (headings in a row)", flush=True)

    # map column index -> (heading, rule), and columns to check for numbers
    col_rules = {}
    number_cols = {}
    for cell in ws[header_row]:
        n = norm(cell.value)
        rule = RULES_NORM.get(n)
        if rule:
            col_rules[cell.column] = (str(cell.value).strip(), rule)
        if n in NUMBER_CHECK_NORM:
            number_cols[cell.column] = str(cell.value).strip()

    def letter_is_free(letter):
        """True if that column does not already hold another known heading."""
        head = ws.cell(row=header_row, column=column_index_from_string(letter)).value
        return norm(head) not in KNOWN_NORM

    # fall back to fixed column letters (N, O) if heading missing
    for name, letter in FALLBACK_LETTERS.items():
        idx = column_index_from_string(letter)
        if name in RULES and idx not in col_rules and letter_is_free(letter) \
                and norm(name) not in {norm(v[0]) for v in col_rules.values()}:
            col_rules[idx] = (name, RULES[name])

    # Region / State number check: by heading; column letter only if the
    # heading is missing and that column is not another known heading
    found = {norm(v) for v in number_cols.values()}
    for name, letter in NUMBER_CHECK_LETTERS.items():
        if norm(name) not in found and letter_is_free(letter):
            number_cols[column_index_from_string(letter)] = name

    checked = ", ".join(f"{get_column_letter(i)} ({n})"
                        for i, n in sorted(number_cols.items())) or "none found"
    print(f"      Number check (red) on columns: {checked}", flush=True)

    red_count = 0
    for row in ws.iter_rows(min_row=header_row + 1):
        for cell in row:
            col_name, rule = col_rules.get(cell.column, (None, None))
            check = cell.column in number_cols
            if rule or check:
                name = col_name or number_cols.get(cell.column, "")
                red_count += process_cell(ws, cell, name, rule, check, log)
    print(f"      Cells highlighted RED in '{ws.title}': {red_count}", flush=True)


# ------------------------------ workbook -----------------------------------
def process_workbook(path):
    print("Loading workbook, please wait...", flush=True)
    wb = load_workbook(path)
    log = []
    sheets_done = 0
    total_sheets = len(wb.worksheets)

    for sheet_no, ws in enumerate(wb.worksheets, start=1):
        vertical = find_vertical_header(ws)
        header_row = None if vertical else find_header_row(ws)
        if not vertical and header_row is None:
            print(f"  [{sheet_no}/{total_sheets}] Skipping sheet '{ws.title}' "
                  f"(expected headings not found)", flush=True)
            continue
        sheets_done += 1
        print(f"  [{sheet_no}/{total_sheets}] Processing sheet '{ws.title}' ...",
              flush=True)
        if vertical:
            process_vertical(ws, vertical, log)
        else:
            process_horizontal(ws, header_row, log)

    print("Saving corrected file, please wait...", flush=True)
    # save in the SAME folder where the input file is kept
    in_path = os.path.abspath(path)
    in_dir = os.path.dirname(in_path)
    base = os.path.splitext(os.path.basename(in_path))[0]
    out_path = os.path.join(in_dir, base + "_corrected.xlsx")
    wb.save(out_path)
    return out_path, log, sheets_done


def select_file():
    """Open a file-selection dialog; fall back to console input."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Select Branch Device details Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        root.destroy()
        return path
    except Exception:
        return input("Enter full path of the Excel file: ").strip().strip('"')


def main():
    print("=" * 60)
    print(" Branch Device Validator")
    print("=" * 60)

    path = select_file()
    if not path:
        print("No file selected. Exiting.")
        sys.exit(0)
    if not os.path.isfile(path):
        print(f"File not found: {path}")
        sys.exit(1)

    print(f"Selected file : {path}")
    print("\n>>> SCRIPT IS RUNNING - please wait, do not close this window <<<\n",
          flush=True)
    try:
        out_path, log, sheets = process_workbook(path)
    except PermissionError:
        print("Cannot read/write file. Close it in Excel and try again.")
        sys.exit(1)

    if sheets == 0:
        print("No sheet with the expected column headings was found.")
        sys.exit(1)

    print(f"\nSheets checked: {sheets}")
    if log:
        print(f"\n{len(log)} correction(s)/warning(s):")
        for sheet, coord, col, old, new in log:
            print(f"  [{sheet}] {coord:<6} {col:<12} {old!r}  ->  {new!r}")
    else:
        print("All values already match the required format.")

    print(f"\nCorrected file saved as: {out_path}")
    print("\n>>> SCRIPT EXECUTION FINISHED - you can close this window <<<")


if __name__ == "__main__":
    main()
    try:
        input("\nPress Enter to exit...")
    except EOFError:
        pass
