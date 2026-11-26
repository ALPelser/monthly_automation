import pdfplumber
import pandas as pd
import re
import sys

# ==============================
# CONFIGURATION
# ==============================

PDF_FILE = "EM28328.pdf"
EXCEL_FILE = "vehicles.xlsx"
OUTPUT_FILE = "vehicles_updated.xlsx"

# If you want to skip pages before the odo section entirely:
START_PAGE = 35  # set to e.g. 5 if odo always starts on page 6

# ==============================
# REGEX DEFINITIONS
# ==============================

PLATE_REGEX = r"""
(
    CL\s?\d{5}         # Western Cape
    |
    [A-Z]{2}\s?\d{2,3}\s?[A-Z]{2}\s?GP  # Gauteng new
    |
    [A-Z]{3}\s?\d{3}\s?GP               # Gauteng old
)
"""

ODO_ROW_REGEX = r"""
\d{2}/\d{2}/\d{4}      # Date
\s+\d{2}:\d{2}         # Time
\s+(\d{5,7})           # ODO (capture this)
"""

# ==============================
# HELPERS
# ==============================

def normalize_plate(plate: str) -> str:
    return re.sub(r"\s+", "", plate).upper()

def exit_with_message(msg):
    print(msg)
    input("Press Enter to exit...")
    sys.exit(1)

# ==============================
# MAIN
# ==============================

def main():
    print("Starting odometer update...\n")

    # Load Excel
    try:
        df = pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        exit_with_message(f"ERROR: {EXCEL_FILE} not found")

    if "Vehicle" not in df.columns:
        exit_with_message("ERROR: Excel must contain a 'Vehicle' column")

    df["Vehicle_norm"] = df["Vehicle"].apply(normalize_plate)

    # Open PDF
    try:
        pdf = pdfplumber.open(PDF_FILE)
    except FileNotFoundError:
        exit_with_message(f"ERROR: {PDF_FILE} not found")

    # Process pages
    for page_index, page in enumerate(pdf.pages[START_PAGE:], start=START_PAGE + 1):
        text = page.extract_text() or ""

        # Extract license plate
        plate_match = re.search(PLATE_REGEX, text, re.VERBOSE)
        if not plate_match:
            print(f"Page {page_index}: No license plate found")
            continue

        pdf_plate = normalize_plate(plate_match.group())

        # Extract odometer readings ONLY from Fuel & Oil section
        fuel_section = False
        odo_values = []

        for line in text.splitlines():
            if "Fuel & Oil" in line:
                fuel_section = True
                continue

            if not fuel_section:
                continue

            match = re.search(ODO_ROW_REGEX, line, re.VERBOSE)
            if match:
                odo_values.append(int(match.group(1)))

        if not odo_values:
            print(f"{pdf_plate}: No fuel odo readings found")
            continue

        last_odo = odo_values[-1]

        # Update Excel
        if pdf_plate not in df["Vehicle_norm"].values:
            print(f"{pdf_plate}: Not found in Excel")
            continue

        df.loc[df["Vehicle_norm"] == pdf_plate, "Month End Odo"] = last_odo
        print(f"{pdf_plate}: Updated to {last_odo}")

    pdf.close()

    # Save output
    df.drop(columns=["Vehicle_norm"], inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)

    print("\nDone!")
    print(f"Updated file saved as: {OUTPUT_FILE}")
    input("Press Enter to exit...")

# ==============================
# ENTRY POINT
# ==============================

if __name__ == "__main__":
    main()
