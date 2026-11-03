"""
Branch Device Details - Excel Validator / Corrector
----------------------------------------------------
On running, asks the user to select an Excel file.
For every sheet whose row 1 holds the known column headings, each data
row (row 2 onwards) is checked against the formatting rules from the
"Description" column of the format sheet, and corrected in place.

Extra rules:
  * Region (col G) / State (col K) containing any digit -> cell RED
  * Device (col N)  "Router" / "router"                 -> "IP router"
  * Environment (col O) "Critical branch"               -> "Critical Branch"

Output: a new file "<original name>_corrected.xlsx" (original is untouched)
        Changed cells are highlighted in yellow, cells with numbers in
        Region/State are highlighted in red, and a change log is printed.

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
#   "bandwidth"    -> Same as given format, e.g. "16 Mbps"
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

# Exceptions: if the value matches (case-insensitive), the normal rule is NOT
# applied and this exact text is written instead.
#   { column heading : { value in lower case : text to write } }
EXCEPTIONS = {
    "Environment": {"critical branch": "Critical Branch"},
}
EXCEPTIONS_NORM = {norm(k): v for k, v in EXCEPTIONS.items()}

# Columns that must be highlighted RED if they contain any number.
# They are checked BY COLUMN LETTER (G = Region, K = State) and ALSO by heading
# name, so the check works even if the heading text/position is different.
NUMBER_CHECK_COLUMNS = ["Region", "State"]
NUMBER_CHECK_LETTERS = {"Region": "G", "State": "K"}

# If a heading is not found in the sheet, these column letters are used
FALLBACK_LETTERS = {
    "Region": "G",
    "State": "K",
    "Device": "N",
    "Environment": "O",
}

# All known headings (used to detect the correct sheet / header row)
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


def apply_rule(rule, value, col_name=""):
    if value is None or str(value).strip() == "":
        return value
    text = str(value).strip()
    # exception values (e.g. Environment "Critical branch") skip the normal rule
    exc = EXCEPTIONS_NORM.get(norm(col_name), {})
    if text.lower() in exc:
        return exc[text.lower()]
    if rule == "capfirst":
        return text.capitalize()
    if rule == "upper":
        return text.upper()
    if rule == "code5":
        return fix_branch_code(value)
    if rule == "bandwidth":
        return fix_bandwidth(value)
    if rule == "device":
        return fix_device(value)
    return value


def find_header_row(ws, max_rows=10):
    """Return row number that looks like the heading row, else None."""
    for r in range(1, min(ws.max_row, max_rows) + 1):
        hits = sum(1 for c in ws[r] if norm(c.value) in KNOWN_NORM)
        if hits >= 5:
            return r
    return None


def process_workbook(path):
    print("Loading workbook, please wait...", flush=True)
    wb = load_workbook(path)
    log = []
    sheets_done = 0
    total_sheets = len(wb.worksheets)

    for sheet_no, ws in enumerate(wb.worksheets, start=1):
        header_row = find_header_row(ws)
        if header_row is None:
            print(f"  [{sheet_no}/{total_sheets}] Skipping sheet '{ws.title}' "
                  f"(expected headings not found)", flush=True)
            continue
        sheets_done += 1
        print(f"  [{sheet_no}/{total_sheets}] Processing sheet '{ws.title}' ...",
              flush=True)

        # map column index -> rule, and set of columns to check for numbers
        col_rules = {}
        number_cols = {}
        for cell in ws[header_row]:
            n = norm(cell.value)
            rule = RULES_NORM.get(n)
            if rule:
                col_rules[cell.column] = (str(cell.value).strip(), rule)
            if n in NUMBER_CHECK_NORM:
                number_cols[cell.column] = str(cell.value).strip()

        # fall back to fixed column letters (N, O ...) if heading missing
        for name, letter in FALLBACK_LETTERS.items():
            idx = column_index_from_string(letter)
            if name in RULES and idx not in col_rules \
                    and norm(name) not in {norm(v[0]) for v in col_rules.values()}:
                col_rules[idx] = (name, RULES[name])

        # ALWAYS check Region = column G and State = column K for numbers
        # (in addition to any column whose heading is Region / State)
        for name, letter in NUMBER_CHECK_LETTERS.items():
            number_cols.setdefault(column_index_from_string(letter), name)

        checked = ", ".join(f"{get_column_letter(i)} ({n})"
                            for i, n in sorted(number_cols.items()))
        print(f"      Number check (red) on columns: {checked}", flush=True)
        red_count = 0

        for row in ws.iter_rows(min_row=header_row + 1):
            for cell in row:
                # ---- formatting corrections ----
                if cell.column in col_rules:
                    col_name, rule = col_rules[cell.column]
                    old = cell.value
                    new = apply_rule(rule, old, col_name)
                    if new != old and str(new) != str(old):
                        cell.value = new
                        if rule == "code5":
                            cell.number_format = "@"   # keep as text
                        cell.fill = HIGHLIGHT
                        log.append((ws.title, cell.coordinate, col_name, old, new))

                    # warn if BranchCode still not 5 digits
                    if rule == "code5":
                        v = str(cell.value).strip() if cell.value is not None else ""
                        if v and not (v.isdigit() and len(v) == 5):
                            log.append((ws.title, cell.coordinate, col_name, old,
                                        "!! NOT 5 DIGIT - check manually"))

                # ---- number check: Region (G) / State (K) -> RED ----
                if cell.column in number_cols and has_number(cell.value):
                    cell.fill = RED_FILL
                    red_count += 1
                    log.append((ws.title, cell.coordinate, number_cols[cell.column],
                                cell.value, "!! CONTAINS NUMBER - highlighted red"))

        print(f"      Cells highlighted RED in '{ws.title}': {red_count}", flush=True)

    print("Saving corrected file, please wait...", flush=True)
    base, ext = os.path.splitext(path)
    out_path = base + "_corrected.xlsx"
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
