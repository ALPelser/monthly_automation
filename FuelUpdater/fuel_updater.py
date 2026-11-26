import PyPDF2
import openpyxl
import re

def extract_fuel_data_from_pdf(pdf_path, start_page=1, end_page=2):
    """
    Extract vehicle and fuel data from PDF.
    Returns a dictionary with vehicle names as keys and fuel amounts as values.
    
    Args:
        pdf_path: Path to PDF file
        start_page: First page to extract (1-indexed, default 1)
        end_page: Last page to extract (1-indexed, default 2)
    """
    fuel_data = {}
    
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        
        # Process specified pages (convert to 0-indexed)
        for page_num in range(start_page - 1, min(end_page, len(pdf_reader.pages))):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            
            # Split into lines and process
            lines = text.split('\n')
            
            for line in lines:
                # Skip header lines and summary lines
                if 'REG. NO.' in line or 'CARD A/C' in line or 'VEH.' in line or 'COST CENTRE' in line:
                    continue
                if 'END OF REPORT' in line or 'BILLING ACCOUNT' in line or 'EXPENSE SUMMARY' in line:
                    continue
                if 'PAGE' in line and len(line.strip().split()) < 5:
                    continue
                    
                # Look for lines starting with card numbers (format: 0008299826 or similar)
                parts = line.strip().split()
                if not parts or not parts[0].isdigit() or len(parts[0]) < 8:
                    continue
                
                if len(parts) < 6:
                    continue
                
                # Extract registration number (parts[2] + parts[3])
                try:
                    reg_part1 = parts[2]
                    reg_part2 = parts[3]
                    
                    # Combine registration parts and remove spaces
                    reg_no = (reg_part1 + reg_part2).replace(' ', '')
                    
                    # Count how many comma-containing numbers we find
                    # The FIRST one should be FUEL
                    # Format: CARD TYPE REG1 REG2 VEHICLE_MAKE DRIVER FUEL OIL TYRES REPAIRS MISC TOLL E-TOLL SERVICE FIN TOTAL
                    
                    comma_values = []
                    for i in range(4, len(parts)):
                        part = parts[i]
                        if ',' in part:
                            test_val = part.replace(',', '').replace('.', '')
                            if test_val.isdigit():
                                comma_values.append(part)
                    
                    # The FIRST comma value is FUEL
                    # If there are NO comma values, fuel is empty (0)
                    if len(comma_values) > 0:
                        fuel_amount = comma_values[0]
                    else:
                        fuel_amount = "0"
                    
                    # Replace comma with period for decimal
                    fuel_amount = fuel_amount.replace(',', '.')
                    
                    fuel_data[reg_no] = float(fuel_amount)
                        
                except (ValueError, IndexError) as e:
                    continue
    
    return fuel_data

def update_excel_with_fuel_data(excel_path, output_path, fuel_data, vehicle_col='A', fuel_col='B', start_row=2):
    """
    Update Excel sheet with fuel data from PDF and save to a new file.
    
    Args:
        excel_path: Path to original Excel file
        output_path: Path to save updated Excel file
        fuel_data: Dictionary with vehicle names and fuel amounts
        vehicle_col: Column letter for vehicle names (default 'A')
        fuel_col: Column letter for fuel amounts (default 'B')
        start_row: First row with data (default 2, assumes row 1 is header)
    """
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active  # Use active sheet, or specify: ws = wb['SheetName']
    
    updated_count = 0
    not_found = []
    
    # Iterate through Excel rows
    for row in range(start_row, ws.max_row + 1):
        vehicle_cell = f'{vehicle_col}{row}'
        fuel_cell = f'{fuel_col}{row}'
        
        vehicle_name = ws[vehicle_cell].value
        
        if vehicle_name:
            vehicle_name = str(vehicle_name).strip()
            
            # Check if vehicle exists in PDF data
            if vehicle_name in fuel_data:
                ws[fuel_cell] = fuel_data[vehicle_name]
                updated_count += 1
                print(f"Updated {vehicle_name}: {fuel_data[vehicle_name]}")
            else:
                not_found.append(vehicle_name)
    
    # Save to new workbook
    wb.save(output_path)
    
    print(f"\n✓ Successfully updated {updated_count} vehicles")
    print(f"✓ Saved to: {output_path}")
    
    if not_found:
        print(f"\n⚠ Vehicles not found in PDF ({len(not_found)}):")
        for vehicle in not_found:
            print(f"  - {vehicle}")
    
    # Check for PDF vehicles not in Excel
    excel_vehicles = set(not_found) | set([str(ws[f'{vehicle_col}{row}'].value).strip() 
                                           for row in range(start_row, ws.max_row + 1) 
                                           if ws[f'{vehicle_col}{row}'].value])
    pdf_only = set(fuel_data.keys()) - excel_vehicles
    
    if pdf_only:
        print(f"\n⚠ Vehicles in PDF but not in Excel ({len(pdf_only)}):")
        for vehicle in pdf_only:
            print(f"  - {vehicle}: {fuel_data[vehicle]}")

# Main execution
if __name__ == "__main__":
    try:
        # Configuration
        PDF_PATH = "EM28328.pdf"      # Change to your PDF file path
        EXCEL_PATH = "vehicles.xlsx"       # Change to your original Excel file path
        OUTPUT_PATH = "vehicles_updated.xlsx"  # New file to save results
        
        # Page configuration (pages where fuel data appears)
        START_PAGE = 8  # First page with fuel data
        END_PAGE = 9    # Last page with fuel data
        
        # Column configuration (adjust as needed)
        VEHICLE_COLUMN = 'A'  # Column with vehicle names
        FUEL_COLUMN = 'B'     # Column to update with fuel amounts
        START_ROW = 2         # First row with data (2 assumes row 1 is header)
        
        print("Starting PDF to Excel fuel data update...\n")
        
        # Extract data from PDF
        print(f"1. Extracting data from PDF pages {START_PAGE}-{END_PAGE}...")
        fuel_data = extract_fuel_data_from_pdf(PDF_PATH, START_PAGE, END_PAGE)
        print(f"   Found {len(fuel_data)} vehicles in PDF\n")
        
        # Update Excel
        print("2. Updating Excel sheet...")
        update_excel_with_fuel_data(
            EXCEL_PATH,
            OUTPUT_PATH,
            fuel_data, 
            vehicle_col=VEHICLE_COLUMN,
            fuel_col=FUEL_COLUMN,
            start_row=START_ROW
        )
        
        print("\n✓ Process completed!")
        
    except FileNotFoundError as e:
        print(f"\n❌ ERROR: File not found!")
        print(f"Details: {e}")
        print(f"\nMake sure these files are in the same folder as this program:")
        print(f"  - {PDF_PATH}")
        print(f"  - {EXCEL_PATH}")
        
    except Exception as e:
        print(f"\n❌ ERROR: Something went wrong!")
        print(f"Error type: {type(e).__name__}")
        print(f"Details: {e}")
        
    finally:
        input("\nPress Enter to close...")