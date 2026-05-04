import openpyxl

data = [
    # Feature 1: Document conversion
    ("Pos_0001", "Document conversion", "Valid DOCX file", "Successfully converts DOCX to PDF", "", "Not Executed", "Supported valid file format"),
    ("Neg_0001", "Document conversion", "Corrupted DOCX file", "Error message indicating invalid or corrupted file", "", "Not Executed", "File format validation works"),
    ("Neg_0002", "Document conversion", "Unsupported EXE file", "Error message indicating wrong or unsupported file format", "", "Not Executed", "Strict format extension restrictions"),
    
    # Feature 2: PDF editing
    ("Pos_0002", "PDF editing", "Valid PDF and add text", "Text is added and PDF saves successfully", "", "Not Executed", "Normal operation conditions met"),
    ("Neg_0003", "PDF editing", "Password-protected PDF", "Prompt to unlock or clear error message", "", "Not Executed", "Security enforcement handles encryption"),
    ("Neg_0004", "PDF editing", "0 byte PDF file", "Error message for empty file", "", "Not Executed", "Empty file handling"),
    
    # Feature 3: Image resizing
    ("Pos_0003", "Image resizing", "Valid JPG, 800x600 size inputs", "Image resized to exactly 800x600", "", "Not Executed", "Valid numeric dimensions accepted"),
    ("Neg_0005", "Image resizing", "Width input: -100", "Validation error requiring positive values", "", "Not Executed", "Input boundary constraints"),
    ("Neg_0006", "Image resizing", "Dimension field with alphabet 'abc'", "Validation error or ignored input", "", "Not Executed", "Data type validation on numeric field"),
    
    # Feature 4: Cropping
    ("Pos_0004", "Cropping", "Valid crop area selection box", "Image cropped strictly to selected area", "", "Not Executed", "Standard behaviour works"),
    ("Neg_0007", "Cropping", "Crop area outside image boundaries", "Cannot select outside bounds / Snap to edge", "", "Not Executed", "UI enforces spatial boundaries"),
    ("Neg_0008", "Cropping", "0x0 area selection", "Submit disabled or error prompt", "", "Not Executed", "Minimum crop dimensions configured"),
    
    # Feature 5: Compression
    ("Pos_0005", "Compression", "Valid 5MB PNG file", "Output file size strictly less than 5MB", "", "Not Executed", "Compression algorithm runs properly"),
    ("Neg_0009", "Compression", "0 byte image file", "Error message for invalid/empty image payload", "", "Not Executed", "Empty file validation triggers"),
    ("Neg_0010", "Compression", "Image exceeding maximum allowed system size (e.g., 5GB)", "Error stating file too large", "", "Not Executed", "Server resource limits checked"),
    
    # Feature 6: Image format conversion
    ("Pos_0006", "Image format conversion", "Valid PNG file, target 'JPG' chosen", "Image securely downloaded as valid JPG", "", "Not Executed", "Format library handles convert"),
    ("Neg_0011", "Image format conversion", "No target format selected", "Conversion disabled, warning to select format", "", "Not Executed", "Mandatory field check active"),
    ("Neg_0012", "Image format conversion", "Fake image (txt file manually renamed to .png)", "Error reading or parsing image file", "", "Not Executed", "File MIME type strict check"),
    
    # Feature 7: Meme generation
    ("Pos_0007", "Meme generation", "Valid text on top and bottom", "Meme generated with appropriate text overlay", "", "Not Executed", "Text overlay functional"),
    ("Neg_0013", "Meme generation", "Extremely long text string (10,000 chars)", "Text truncated or length limit error displayed", "", "Not Executed", "Max character limits limit size"),
    ("Neg_0014", "Meme generation", "Unsupported binary or zalgo characters", "Characters stripped safely or error returned", "", "Not Executed", "Safe encoding validation"),
    
    # Feature 8: Color picker
    ("Pos_0008", "Color picker", "Click valid pixel within image bounds", "Correct HEX code displayed instantly", "", "Not Executed", "Color extraction parses correct RGBA"),
    ("Neg_0015", "Color picker", "Click outside image on webpage background", "No color picked / Picker ignores background", "", "Not Executed", "Canvas bounds checking works"),
    ("Neg_0016", "Color picker", "Picker action triggered while image is still loading", "Picker disabled or shows loading cursor", "", "Not Executed", "Async DOM state handling"),
    
    # Feature 9: Image rotation
    ("Pos_0009", "Image rotation", "Click rotate 90 degrees right", "Image appears properly rotated 90 deg", "", "Not Executed", "Transform function applies properly"),
    ("Neg_0017", "Image rotation", "Manual angle input 'XYZ'", "Validation error asking for integer", "", "Not Executed", "Input data type validation"),
    ("Neg_0018", "Image rotation", "Corrupted image uploaded before clicking rotate", "General error loading image initially", "", "Not Executed", "Prerequisite dependencies evaluated"),
    
    # Feature 10: Image flipping
    ("Pos_0010", "Image flipping", "Click flip horizontally", "Image natively mirrored horizontally", "", "Not Executed", "Transform CSS applies cleanly"),
    ("Neg_0019", "Image flipping", "Rapid consecutive spam clicks (10+ times) on flip", "System debounces without browser crash", "", "Not Executed", "Debounce and queue rate limiting"),
    ("Neg_0020", "Image flipping", "Flip action on already removed image", "Error message or UI element resets", "", "Not Executed", "State validity and synchronization"),
    
    # Extra 6 Cases
    ("Pos_0011", "Document conversion", "Multi-page PDF to DOCX conversion", "All pages converted seamlessly into single Doc", "", "Not Executed", "Pagination handlers complete"),
    ("Neg_0021", "Document conversion", "Disconnect internet during ongoing conversion", "Graceful network failure warning message", "", "Not Executed", "Frontend network connectivity error handling"),
    ("Pos_0012", "PDF editing", "Draw visible rectangle and annotation on page", "Shapes and text saved correctly on output PDF", "", "Not Executed", "Draw annotation feature robust"),
    ("Neg_0022", "PDF editing", "Try to delete all pages in a PDF document", "Warning message or prevent deletion of last page", "", "Not Executed", "Minimum 1 page baseline rule in PDF spec"),
    ("Pos_0013", "Image resizing", "'Keep aspect ratio' checked, edit width", "Height automatically recalculates appropriately", "", "Not Executed", "Aspect ratio math executes immediately"),
    ("Neg_0023", "Image resizing", "Resize to arbitrary massive value (99,999,999 pixels)", "Application error indicating limit threshold reached", "", "Not Executed", "Sanity scaling limits enforce security against mem abuse")
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Test Cases"

headers = [
    "TC ID", "Application Feature Tested", "Input",
    "Expected output", "Actual output", "Status",
    "Assumption for Expected Output"
]
ws.append(headers)

for row in data:
    ws.append(row)

# Format headers
for cell in ws[1]:
    cell.font = openpyxl.styles.Font(bold=True)

# Adjust column widths slightly
column_widths = {'A': 10, 'B': 25, 'C': 35, 'D': 40, 'E': 15, 'F': 15, 'G': 40}
for col, width in column_widths.items():
    ws.column_dimensions[col].width = width

wb.save("Manual Test Cases for Option 2.xlsx")
print("Saved to Manual Test Cases for Option 2.xlsx")
