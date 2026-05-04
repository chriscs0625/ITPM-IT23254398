import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side
from pathlib import Path
import sys

file_path = Path("Manual Test Cases for Option 2.xlsx")

if not file_path.exists():
    print("Excel file not found!")
    sys.exit(1)

try:
    # Load the completely raw data file created by pandas
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    ws.title = "Manual Test Cases"
    
    # Styling definitions
    wrap_left = Alignment(wrap_text=True, vertical='center', horizontal='left')
    wrap_center = Alignment(wrap_text=True, vertical='center', horizontal='center')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    bold_font = Font(bold=True)
    
    # Set proper column widths so nothing is cut off
    widths = {'A': 12, 'B': 25, 'C': 40, 'D': 45, 'E': 15, 'F': 15, 'G': 30}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width
        
    # Apply formatting to all cells
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border
            
            # First row is Header
            if cell.row == 1:
                cell.font = bold_font
                cell.alignment = wrap_center
            else:
                # Center align TC ID, Feature, Actual Output, Status
                if cell.column in [1, 2, 5, 6]:
                    cell.alignment = wrap_center
                else:
                    cell.alignment = wrap_left

    # Freeze the top header row
    ws.freeze_panes = "A2"

    wb.save(file_path)
    print("Successfully formatted the Excel file! It is no longer squished.")

except PermissionError:
    print("PERMISSION ERROR: The Excel file is currently OPEN. Please close Microsoft Excel, and run again.")
except Exception as e:
    print(f"Error: {e}")
