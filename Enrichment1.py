#Enrichment 1

import os
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# ============================================================
# CONFIGURATION
# ============================================================

CORRECT_STATES = [
    "AMARAVATI",
    "BENGALURU",
    "BHOPAL",
    "BHUBANESWAR",
    "CHANDIGARH",
    "CHENNAI",
    "DELHI",
    "GANDHINAGAR",
    "HYDERABAD",
    "JAIPUR",
    "KOLKATA",
    "LUCKNOW",
    "MAHARASHTRA",
    "MUMBAI METRO",
    "NORTH-EASTERN",
    "PATNA",
    "THIRUVANANTHAPURAM",
]

LEVENSHTEIN_THRESHOLD = 3


# Red fill - equivalent to VBA vbRed
RED_FILL = PatternFill(
    fill_type="solid",
    fgColor="FF0000"
)


# ============================================================
# LEVENSHTEIN DISTANCE
# ============================================================

def levenshtein(s1: str, s2: str) -> int:

    s1 = str(s1)
    s2 = str(s2)

    rows = len(s1) + 1
    cols = len(s2) + 1

    matrix = [[0] * cols for _ in range(rows)]

    for i in range(rows):
        matrix[i][0] = i

    for j in range(cols):
        matrix[0][j] = j

    for i in range(1, rows):

        for j in range(1, cols):

            if s1[i - 1] == s2[j - 1]:
                cost = 0
            else:
                cost = 1

            matrix[i][j] = min(
                matrix[i - 1][j] + 1,
                matrix[i][j - 1] + 1,
                matrix[i - 1][j - 1] + cost
            )

    return matrix[-1][-1]


# ============================================================
# GET CORRECT STATE
# ============================================================

def get_correct_state(input_text: str) -> str:

    if input_text is None:
        return ""

    input_text = str(input_text).strip().upper()

    if not input_text:
        return ""

    best_match = ""
    min_dist = 999

    for correct_value in CORRECT_STATES:

        distance = levenshtein(
            input_text,
            correct_value
        )

        if distance < min_dist:
            min_dist = distance
            best_match = correct_value

    if min_dist <= LEVENSHTEIN_THRESHOLD:
        return best_match

    return ""


# ============================================================
# IP VALIDATION
# ============================================================

def is_valid_ip(ip) -> bool:

    if ip is None:
        return False

    ip = str(ip).strip()

    if not ip:
        return False

    parts = ip.split(".")

    if len(parts) != 4:
        return False

    for part in parts:

        if not part.isdigit():
            return False

        value = int(part)

        if value < 0 or value > 255:
            return False

    return True


# ============================================================
# SAFE STRING
# ============================================================

def clean_string(value) -> str:

    if value is None:
        return ""

    return str(value).strip()


# ============================================================
# MARK CELL RED
# ============================================================

def mark_invalid(cell):

    cell.fill = RED_FILL


# ============================================================
# VALIDATE DATA
# ============================================================

def validate_data(ws):

    # Find last row using column A
    last_row = ws.max_row

    processed_rows = 0
    invalid_cells = 0

    # --------------------------------------------------------
    # Process rows starting from row 2
    # --------------------------------------------------------

    for row in range(2, last_row + 1):

        processed_rows += 1

        # ====================================================
        # 1. WANLinkIP - A
        # ====================================================

        cell = ws[f"A{row}"]

        if not is_valid_ip(cell.value):
            mark_invalid(cell)
            invalid_cells += 1

        # ====================================================
        # 2. DeviceIP - B
        # ====================================================

        cell = ws[f"B{row}"]

        if not is_valid_ip(cell.value):
            mark_invalid(cell)
            invalid_cells += 1

        # ====================================================
        # 3. LinkUsage - C
        # ====================================================

        cell = ws[f"C{row}"]

        value = clean_string(cell.value).lower()

        if value == "primary":
            cell.value = "Primary"

        elif value == "secondary":
            cell.value = "Secondary"

        elif value == "tertiary":
            cell.value = "Tertiary"

        else:
            mark_invalid(cell)
            invalid_cells += 1

        # ====================================================
        # 4. ISP - D
        # ====================================================

        cell = ws[f"D{row}"]

        value = clean_string(cell.value)

        normalized_isp = value.lower().replace(" ", "")

        isp_map = {
            "airtel": "AIRTEL",
            "bsnl": "BSNL",
            "hughes": "HUGHES",
            "mtnl": "MTNL",
            "pgcil": "PGCIL",
            "railtel": "RAILTEL",
            "rjio": "RJIO",
            "sify": "SIFY",
            "tcl": "TCL",
            "vodafone": "VODAFONE",
        }

        if normalized_isp in isp_map:
            cell.value = isp_map[normalized_isp]

        else:
            mark_invalid(cell)
            invalid_cells += 1

        # ====================================================
        # 5. BranchCode - F
        # ====================================================

        cell = ws[f"F{row}"]

        if cell.value is not None and clean_string(cell.value):

            try:

                number = int(float(cell.value))

                # Equivalent to VBA Format(value, "00000")
                cell.value = f"{number:05d}"

            except (ValueError, TypeError):

                mark_invalid(cell)
                invalid_cells += 1

        # ====================================================
        # 6. Region - H
        # ====================================================

        cell = ws[f"H{row}"]

        value = clean_string(cell.value).lower()

        region_map = {
            "east": "East",
            "west": "West",
            "north": "North",
            "south": "South",
        }

        if value in region_map:
            cell.value = region_map[value]

        else:
            mark_invalid(cell)
            invalid_cells += 1

        # ====================================================
        # 7. Circle - I
        # ====================================================

        cell = ws[f"I{row}"]
        if isinstance(cell.value, str) and cell.value.islower():
            value = clean_string(cell.value)
            cell.value = cell.value.upper()
        if value.upper() == "BANGLORE":
            cell.value = "BENGALURU"

        # ====================================================
        # 8. State - J
        # ====================================================

        cell = ws[f"J{row}"]

        value = clean_string(cell.value)

        if value:
            cell.value = value.upper()

        # ====================================================
        # 9. Country - K
        # ====================================================

        cell = ws[f"K{row}"]

        value = clean_string(cell.value)

        if value.upper() != "INDIA":

            mark_invalid(cell)
            invalid_cells += 1

        else:

            cell.value = "India"

        # ====================================================
        # 10. CountryFo - N
        # ====================================================

        cell = ws[f"N{row}"]

        value = clean_string(cell.value)

        if value.upper() != "INDIA":

            mark_invalid(cell)
            invalid_cells += 1

        else:

            cell.value = "India"

        # ====================================================
        # 11. Environment - U
        # ====================================================

        cell = ws[f"U{row}"]

        value = clean_string(cell.value).lower()

        environment_map = {
            "branch": "Branch",
            "critical branch": "Critical Branch",
            "contact centre": "Contact Centre",
            "corporate centre": "CORPORATE CENTRE",
            "ec,atm & other branches": "EC,ATM & Other Branches",
        }

        if value in environment_map:
            cell.value = environment_map[value]

        else:
            mark_invalid(cell)
            invalid_cells += 1

    return processed_rows, invalid_cells


# ============================================================
# FILE SELECTION
# ============================================================

def select_excel_file():

    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[
            ("Excel files", "*.xlsx"),
            ("Excel Macro-Enabled", "*.xlsm"),
            ("All files", "*.*"),
        ],
    )

    root.destroy()

    return file_path


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 60)
    print("EXCEL DATA VALIDATOR - OPENPYXL VERSION")
    print("=" * 60)

    # --------------------------------------------------------
    # Select Excel file
    # --------------------------------------------------------

    file_path = select_excel_file()

    if not file_path:

        print("No file selected.")
        return

    try:

        print(f"\nSelected file:")
        print(file_path)

        # ----------------------------------------------------
        # Load workbook
        # ----------------------------------------------------

        keep_vba = file_path.lower().endswith(".xlsm")

        wb = load_workbook(
            file_path,
            keep_vba=keep_vba
        )

        # ----------------------------------------------------
        # First worksheet
        # ----------------------------------------------------

        ws = wb.worksheets[0]

        print(f"\nWorksheet: {ws.title}")
        print("Validating data...")

        # ----------------------------------------------------
        # Validate
        # ----------------------------------------------------

        processed_rows, invalid_cells = validate_data(ws)

        # ----------------------------------------------------
        # Output filename
        # ----------------------------------------------------

        directory = os.path.dirname(file_path)
        filename = os.path.basename(file_path)

        name, extension = os.path.splitext(filename)

        output_file = os.path.join(
            directory,
            f"{name}_validated{extension}"
        )

        # ----------------------------------------------------
        # Save
        # ----------------------------------------------------

        wb.save(output_file)

        print("\nValidation completed successfully.")
        print(f"Output file:")
        print(output_file)

        print(f"\nRows processed: {processed_rows}")
        print(f"Invalid cells: {invalid_cells}")

        # ----------------------------------------------------
        # Show completion message
        # ----------------------------------------------------

        root = tk.Tk()
        root.withdraw()

        messagebox.showinfo(
            "Validation Completed",
            f"Validation completed successfully!\n\n"
            f"Output file:\n{output_file}\n\n"
            f"Rows processed: {processed_rows}\n"
            f"Invalid cells: {invalid_cells}"
        )

        root.destroy()

    except PermissionError:

        message = (
            "Unable to save the output file.\n\n"
            "Please make sure the Excel file is not open "
            "in Microsoft Excel."
        )

        print(message)

        root = tk.Tk()
        root.withdraw()

        messagebox.showerror(
            "Permission Error",
            message
        )

        root.destroy()

    except Exception as e:

        print("\nERROR:")
        print(str(e))

        root = tk.Tk()
        root.withdraw()

        messagebox.showerror(
            "Validation Error",
            f"An error occurred:\n\n{str(e)}"
        )

        root.destroy()


if __name__ == "__main__":
    main()
