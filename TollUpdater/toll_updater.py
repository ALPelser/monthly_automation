import PyPDF2
import openpyxl

def extract_toll_data_from_pdf(pdf_path, start_page=1, end_page=2):
    """
    Extract vehicle and combined toll data from PDF using character positions.
    Returns a dictionary with vehicle names as keys and combined toll amount as values.
    
    Args:
        pdf_path: Path to PDF file
        start_page: First page to extract (1-indexed, default 1)
        end_page: Last page to extract (1-indexed, default 2)
    """
    toll_data = {}
    
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        
        # Process specified pages (convert to 0-indexed)
        for page_num in range(start_page - 1, min(end_page, len(pdf_reader.pages))):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            
            # Split into lines and process
            lines = text.split('\n')
            
            # First, find the header line to determine column positions
            toll_col_start = None
            toll_col_end = None
            etoll_col_start = None
            etoll_col_end = None
            
            for line in lines:
                if 'TOLL' in line and 'E-TOLL' in line and 'SERVICE' in line:
                    # Found the header line
                    # Find positions of TOLL and E-TOLL columns
                    toll_idx = line.find('TOLL')
                    etoll_idx = line.find('E-TOLL')
                    service_idx = line.find('SERVICE')
                    
                    # Expand the column ranges slightly to capture full numbers
                    # TOLL column: start a bit earlier to catch numbers that might start before header
                    toll_col_start = max(0, toll_idx - 5)  # Start 5 chars before
                    toll_col_end = etoll_idx
                    
                    # E-TOLL column: also expand the range
                    etoll_col_start = etoll_idx - 5  # Start 5 chars before
                    etoll_col_end = service_idx
                    
                    print(f"Found column positions - TOLL: {toll_col_start}-{toll_col_end}, E-TOLL: {etoll_col_start}-{etoll_col_end}")
                    break
            
            if toll_col_start is None:
                print("Warning: Could not find TOLL and E-TOLL column headers")
                continue
            
            for line in lines:
                # Skip header lines and summary lines
                if 'REG. NO.' in line or 'CARD A/C' in line or 'VEH.' in line or 'COST CENTRE' in line:
                    continue
                if 'END OF REPORT' in line or 'BILLING ACCOUNT' in line or 'EXPENSE SUMMARY' in line:
                    continue
                if 'PAGE' in line or 'TOLL' in line:
                    continue
                
                # Check if line is long enough to contain toll columns
                if len(line) < etoll_col_end:
                    continue
                    
                # Look for lines starting with card numbers
                parts = line.strip().split()
                if not parts or not parts[0].isdigit() or len(parts[0]) < 8:
                    continue
                
                if len(parts) < 4:
                    continue
                
                # Extract registration number (parts[2] + parts[3])
                try:
                    reg_part1 = parts[2]
                    reg_part2 = parts[3]
                    
                    # Combine registration parts and remove spaces
                    reg_no = (reg_part1 + reg_part2).replace(' ', '')
                    
                    # Extract TOLL and E-TOLL using character positions
                    # Get a wider range to ensure we capture full numbers
                    toll_text = line[toll_col_start:toll_col_end].strip()
                    etoll_text = line[etoll_col_start:etoll_col_end].strip()
                    
                    # Parse TOLL value - look for all numbers with commas
                    toll_amount = 0.0
                    if toll_text:
                        # Split and find the number with comma
                        toll_parts = toll_text.split()
                        for part in toll_parts:
                            # Check if it contains a comma (decimal separator)
                            if ',' in part:
                                # Remove any non-numeric characters except comma and digits
                                cleaned = ''.join(c for c in part if c.isdigit() or c == ',')
                                if cleaned:
                                    toll_amount = float(cleaned.replace(',', '.'))
                                    break
                    
                    # Parse E-TOLL value - look for all numbers with commas
                    etoll_amount = 0.0
                    if etoll_text:
                        # Split and find the number with comma
                        etoll_parts = etoll_text.split()
                        for part in etoll_parts:
                            # Check if it contains a comma (decimal separator)
                            if ',' in part:
                                # Remove any non-numeric characters except comma and digits
                                cleaned = ''.join(c for c in part if c.isdigit() or c == ',')
                                if cleaned:
                                    etoll_amount = float(cleaned.replace(',', '.'))
                                    break
                    
                    # Combine TOLL and E-TOLL into a single value
                    combined_toll = toll_amount + etoll_amount
                    
                    toll_data[reg_no] = combined_toll
                        
                except (ValueError, IndexError) as e:
                    continue
    
    return toll_data

def update_excel_with_toll_data(excel_path, output_path, toll_data, vehicle_col='A', toll_col='B', start_row=2):
    """
    Update Excel sheet with combined toll data from PDF and save to a new file.
    
    Args:
        excel_path: Path to original Excel file
        output_path: Path to save updated Excel file
        toll_data: Dictionary with vehicle names and combined toll amounts
        vehicle_col: Column letter for vehicle names (default 'A')
        toll_col: Column letter for combined toll amounts (default 'B')
        start_row: First row with data (default 2, assumes row 1 is header)
    """
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    
    updated_count = 0
    not_found = []
    
    # Iterate through Excel rows
    for row in range(start_row, ws.max_row + 1):
        vehicle_cell = f'{vehicle_col}{row}'
        toll_cell = f'{toll_col}{row}'
        
        vehicle_name = ws[vehicle_cell].value
        
        if vehicle_name:
            vehicle_name = str(vehicle_name).strip()
            
            # Check if vehicle exists in PDF data
            if vehicle_name in toll_data:
                combined_toll = toll_data[vehicle_name]
                ws[toll_cell] = combined_toll
                updated_count += 1
                print(f"Updated {vehicle_name}: {combined_toll}")
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
    pdf_only = set(toll_data.keys()) - excel_vehicles
    
    if pdf_only:
        print(f"\n⚠ Vehicles in PDF but not in Excel ({len(pdf_only)}):")
        for vehicle in pdf_only:
            print(f"  - {vehicle}: {toll_data[vehicle]}")

# Main execution
if __name__ == "__main__":
    # Configuration
    PDF_PATH = "EM28328.pdf"           # Your PDF file path
    EXCEL_PATH = "vehicles.xlsx"           # Your original Excel file path
    OUTPUT_PATH = "vehicles_updated.xlsx"  # New file to save toll results
    
    # Page configuration (pages where toll data appears)
    START_PAGE = 8  # First page with toll data
    END_PAGE = 9    # Last page with toll data
    
    # Column configuration (adjust as needed)
    VEHICLE_COLUMN = 'A'  # Column with vehicle names
    TOLL_COLUMN = 'B'     # Column to update with combined toll amounts (TOLL + E-TOLL)
    START_ROW = 2         # First row with data (2 assumes row 1 is header)
    
    print("Starting PDF to Excel toll data update...\n")
    
    # Extract data from PDF
    print(f"1. Extracting toll data from PDF pages {START_PAGE}-{END_PAGE}...")
    toll_data = extract_toll_data_from_pdf(PDF_PATH, START_PAGE, END_PAGE)
    print(f"   Found {len(toll_data)} vehicles in PDF\n")
    
    # Update Excel
    print("2. Updating Excel sheet...")
    update_excel_with_toll_data(
        EXCEL_PATH,
        OUTPUT_PATH,
        toll_data, 
        vehicle_col=VEHICLE_COLUMN,
        toll_col=TOLL_COLUMN,
        start_row=START_ROW
    )
    
    print("\n✓ Process completed!")