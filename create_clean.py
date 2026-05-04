import pandas as pd
from pathlib import Path

headers = [
    "TC ID", 
    "Application Feature Tested", 
    "Input", 
    "Expected output", 
    "Actual output", 
    "Status", 
    "Assumption for Expected Output"
]

# Exactly 36 Test Cases - strictly following the assignment guidelines
data = [
    # 1. Document conversion (3 cases)
    ("Pos_0001", "Document conversion", "Upload valid DOCX file and click convert", "File successfully converts to PDF and download link appears", "TBD", "Not Executed", "NA"),
    ("Neg_0001", "Document conversion", "Upload an unsupported HTML file", "System rejects the file and shows an 'Invalid format' message", "TBD", "Not Executed", "NA"),
    ("Neg_0002", "Document conversion", "Click convert without uploading any file", "System shows an error message prompting to upload a file first", "TBD", "Not Executed", "NA"),

    # 2. PDF editing (3 cases)
    ("Pos_0002", "PDF editing", "Add a text box to the PDF and type 'Test'", "Text box appears on the PDF and displays 'Test'", "TBD", "Not Executed", "NA"),
    ("Neg_0003", "PDF editing", "Try to save a PDF without making any changes", "System should indicate that no changes were made to save", "TBD", "Not Executed", "NA"),
    ("Neg_0004", "PDF editing", "Place a text box outside the PDF page canvas", "Text box is constrained to the PDF page boundaries", "TBD", "Not Executed", "NA"),

    # 3. Image resizing (3 cases + 1 extra = 4)
    ("Pos_0003", "Image resizing", "Enter width 500 and height 500", "Image preview resizes to the specified 500x500 dimensions", "TBD", "Not Executed", "NA"),
    ("Pos_0004", "Image resizing", "Select 'Maintain aspect ratio' and change width", "Height automatically updates to maintain the original aspect ratio", "TBD", "Not Executed", "NA"),
    ("Neg_0005", "Image resizing", "Enter alphabetical characters ('abc') in the width field", "Input is rejected, and only numeric values are allowed", "TBD", "Not Executed", "NA"),
    ("Neg_0006", "Image resizing", "Enter a negative value (-100) for height", "System displays a validation error for negative dimensions", "TBD", "Not Executed", "NA"),

    # 4. Cropping (3 cases + 1 extra = 4)
    ("Pos_0005", "Cropping", "Draw a valid crop rectangle and click apply", "Image is cropped exactly to the defined rectangle", "TBD", "Not Executed", "NA"),
    ("Pos_0006", "Cropping", "Select predefined 1:1 format and apply crop", "Crop boundaries lock into a perfect square", "TBD", "Not Executed", "NA"),
    ("Neg_0007", "Cropping", "Attempt to drag crop handles completely off the image", "Crop handles automatically snap back to the image edges", "TBD", "Not Executed", "NA"),
    ("Neg_0008", "Cropping", "Click apply crop without drawing any crop box", "System warns the user to select an area to crop first", "TBD", "Not Executed", "NA"),

    # 5. Compression (3 cases + 1 extra = 4)
    ("Pos_0007", "Compression", "Adjust compression slider to 30%", "Preview updates and estimated file size is appropriately reduced", "TBD", "Not Executed", "NA"),
    ("Pos_0008", "Compression", "Click reset after compressing", "Image preview completely reverts to original uncompressed quality", "TBD", "Not Executed", "NA"),
    ("Neg_0009", "Compression", "Manually type 200% into compression text box", "System caps the value at the maximum 100% allowed limit", "TBD", "Not Executed", "NA"),
    ("Neg_0010", "Compression", "Drag the slider down to 0%", "System enforces a minimum of 1% compression logic", "TBD", "Not Executed", "NA"),

    # 6. Image format conversion (3 cases)
    ("Pos_0009", "Image format conversion", "Select PNG as target for a JPG image", "Image intelligently converts and downloads as a standard PNG", "TBD", "Not Executed", "NA"),
    ("Neg_0011", "Image format conversion", "Select target format same as source format (PNG to PNG)", "System disables the convert button or shows 'Already in this format' message", "TBD", "Not Executed", "NA"),
    ("Neg_0012", "Image format conversion", "Click convert without selecting a target format", "System prompts the user to select an output format from the dropdown", "TBD", "Not Executed", "NA"),

    # 7. Meme generation (3 cases + 1 extra = 4)
    ("Pos_0010", "Meme generation", "Type 'Hello' in top text and 'World' in bottom text", "White text with black border overlays the image preview accordingly", "TBD", "Not Executed", "NA"),
    ("Neg_0013", "Meme generation", "Click generate meme before uploading the background image", "System displays 'Image required' popup or warning", "TBD", "Not Executed", "Meme generation requires image input"),
    ("Neg_0014", "Meme generation", "Submit purely blank spaces in the text field", "Validation correctly prevents generating with empty spaces", "TBD", "Not Executed", "NA"),
    ("Neg_0015", "Meme generation", "Enter excessively long string without spaces", "Text safely wraps or scales down to stay within image boundaries", "TBD", "Not Executed", "NA"),

    # 8. Color picker (3 cases + 1 extra = 4)
    ("Pos_0011", "Color picker", "Click a blue pixel on the image", "Color picker displays the correct HEX/RGB code for the clicked blue area", "TBD", "Not Executed", "NA"),
    ("Pos_0012", "Color picker", "Click the copy button next to the HEX code", "The HEX code is successfully copied to the user's clipboard", "TBD", "Not Executed", "NA"),
    ("Neg_0016", "Color picker", "Click outside the image entirely", "Color picker ignores the click and the HEX field remains unchanged", "TBD", "Not Executed", "NA"),
    ("Neg_0017", "Color picker", "Manually type an invalid code 'XYZ123' in the HEX box", "Input is rejected and reverts to the last recorded valid color", "TBD", "Not Executed", "NA"),

    # 9. Image rotation (3 cases + 1 extra = 4)
    ("Pos_0013", "Image rotation", "Click 'Rotate 90° Right'", "Image perfectly rotates 90 degrees clockwise", "TBD", "Not Executed", "NA"),
    ("Neg_0018", "Image rotation", "Try to interact with rotation controls without an image uploaded", "Rotation buttons are visibly disabled and unclickable", "TBD", "Not Executed", "Image is required before rotation"),
    ("Neg_0019", "Image rotation", "Input special characters '$$$' into a custom rotation angle field", "The numerical field natively blocks non-numeric inputs", "TBD", "Not Executed", "NA"),
    ("Neg_0020", "Image rotation", "Try to input angle of 500 degrees", "Angle auto-corrects to 140 degrees (500 modulo 360) or restricts input", "TBD", "Not Executed", "NA"),

    # 10. Image flipping (3 cases)
    ("Pos_0014", "Image flipping", "Click 'Flip Horizontal'", "Image instantly mirrors effectively across the vertical axis", "TBD", "Not Executed", "NA"),
    ("Neg_0021", "Image flipping", "Click the flip buttons continuously when no image is present", "Application takes no actions and displays no glitching", "TBD", "Not Executed", "Image is required before flipping"),
    ("Neg_0022", "Image flipping", "Double click the 'Flip Vertical' button rapidly", "Sequence correctly queues and flips image back to original state without hanging", "TBD", "Not Executed", "NA"),
]

df = pd.DataFrame(data, columns=headers)

# Save strictly as requested
csv_path = Path("Manual Test Cases for Option 2.csv")
xlsx_path = Path("Manual Test Cases for Option 2.xlsx")

df.to_csv(csv_path, index=False)
df.to_excel(xlsx_path, index=False)

print("Files completely rebuilt from scratch!")
