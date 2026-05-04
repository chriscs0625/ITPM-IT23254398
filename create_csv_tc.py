import csv
from pathlib import Path

# Headers
headers = [
    "TC ID", 
    "Application Feature Tested", 
    "Input", 
    "Expected output", 
    "Actual output", 
    "Status", 
    "Assumption for Expected Output"
]

# Data for 36 test cases based on UI behavior (no backend/security/etc.)
raw_data = [
    # 1. Document Conversion (1 Pos, 2 Neg)
    ("P", "Document conversion", "Valid DOCX file upload", "File is converted to PDF format successfully", "TBD", "Not Executed", "Valid file provided"),
    ("N", "Document conversion", "Corrupted DOCX file", "UI shows error message 'Invalid file'", "TBD", "Not Executed", "NA"),
    ("N", "Document conversion", "DOCX file exceeding 50MB size limit", "UI shows error message 'File too large'", "TBD", "Not Executed", "Limit is 50MB"),

    # 2. PDF Editing (1 Pos, 2 Neg)
    ("P", "PDF editing", "Add text annotation to PDF", "Text appears correctly on the PDF canvas", "TBD", "Not Executed", "NA"),
    ("N", "PDF editing", "Upload password-protected PDF", "Prompt asking for password appears rather than crashing", "TBD", "Not Executed", "NA"),
    ("N", "PDF editing", "Save edits on 0-page empty PDF", "UI displays error 'Empty document cannot be edited'", "TBD", "Not Executed", "NA"),

    # 3. Image Resizing (2 Pos, 3 Neg) - 2 Extras
    ("P", "Image resizing", "Enter valid dimensions 800x600 px", "Image is resized exactly to 800x600 px rendering clearly", "TBD", "Not Executed", "NA"),
    ("N", "Image resizing", "Enter negative width (e.g., -100) in input", "Input field rejects value or shows validation error", "TBD", "Not Executed", "NA"),
    ("N", "Image resizing", "Enter alphabet characters 'abc' in height", "Input field prevents typing letters", "TBD", "Not Executed", "NA"),
    ("P", "Image resizing", "Select 50% scale from dropdown", "Image scales down to exactly half its original size properly", "TBD", "Not Executed", "NA"),
    ("N", "Image resizing", "Enter an excessive scale of 5000%", "Validation error showing 'Scale must be under 1000%' in UI", "TBD", "Not Executed", "NA"),

    # 4. Cropping (2 Pos, 3 Neg) - 2 Extras
    ("P", "Cropping", "Draw crop box entirely inside image and apply", "Image is cropped precisely to the selected region on UI", "TBD", "Not Executed", "NA"),
    ("N", "Cropping", "Drag crop box partially outside image bounds", "Crop box is properly restricted to image edges", "TBD", "Not Executed", "NA"),
    ("N", "Cropping", "Set crop dimensions to 0x0 via handles", "Crop action is disabled or shows error message", "TBD", "Not Executed", "NA"),
    ("P", "Cropping", "Select 1:1 aspect ratio constraint and drag", "Crop box maintains a perfect square shape during drag", "TBD", "Not Executed", "NA"),
    ("N", "Cropping", "Input negative crop coordinates manually", "Warning tooltip that coordinates must be positive", "TBD", "Not Executed", "NA"),

    # 5. Compression (1 Pos, 3 Neg) - 1 Extra
    ("P", "Compression", "Upload JPG and select 'High Compression'", "File size metric drops, visual stays acceptable on display", "TBD", "Not Executed", "NA"),
    ("N", "Compression", "Upload a non-image file (.txt)", "UI rejects upload promptly with 'Invalid image format' error", "TBD", "Not Executed", "NA"),
    ("N", "Compression", "Select 0% (or empty) quality slider", "Slider snaps to minimum 1% or shows validation warning", "TBD", "Not Executed", "NA"),
    ("N", "Compression", "Upload an already highly compressed 1KB image", "Notifies user 'File is already maximally compressed'", "TBD", "Not Executed", "NA"),

    # 6. Image format conversion (1 Pos, 2 Neg)
    ("P", "Image format conversion", "Upload PNG and convert to JPG", "Converted file downloads as valid JPG", "TBD", "Not Executed", "NA"),
    ("N", "Image format conversion", "Upload unsupported format (.xyz)", "Error message 'Unsupported file format' shown on UI", "TBD", "Not Executed", "NA"),
    ("N", "Image format conversion", "Convert animated GIF to PNG", "User is warned that animation will be lost on format change", "TBD", "Not Executed", "NA"),

    # 7. Meme generation (1 Pos, 2 Neg)
    ("P", "Meme generation", "Enter top and bottom text on image template", "Text is overlaid properly with white font and black stroke", "TBD", "Not Executed", "NA"),
    ("N", "Meme generation", "Enter exceptionally long text string", "Text auto-scales down or wraps without breaking UI layout", "TBD", "Not Executed", "NA"),
    ("N", "Meme generation", "Enter purely blank spaces in text inputs", "No text rendered, or validation 'Text cannot be empty'", "TBD", "Not Executed", "NA"),

    # 8. Color picker (1 Pos, 2 Neg)
    ("P", "Color picker", "Click on a solid red area in the image", "Color picker correctly displays hex code #FF0000", "TBD", "Not Executed", "NA"),
    ("N", "Color picker", "Use color picker outside the image area bounds", "Action is safely ignored; no color picked", "TBD", "Not Executed", "NA"),
    ("N", "Color picker", "Click on fully transparent PNG pixel", "Displays transparent code or warns 'No color detected'", "TBD", "Not Executed", "NA"),

    # 9. Image rotation (1 Pos, 2 Neg)
    ("P", "Image rotation", "Click 'Rotate 90° Right' button once", "Image rotates exactly 90 degrees clockwise", "TBD", "Not Executed", "NA"),
    ("N", "Image rotation", "Input '90.5' in custom angle field", "Input field safely rejects decimal values", "TBD", "Not Executed", "NA"),
    ("N", "Image rotation", "Input angle exceeding maximum (e.g. 99999)", "Angle is neatly rounded modulo 360", "TBD", "Not Executed", "NA"),

    # 10. Image flipping (2 Pos, 2 Neg) - 1 Extra
    ("P", "Image flipping", "Click 'Flip Horizontal' button", "Image visuals mirror correctly horizontally immediately", "TBD", "Not Executed", "NA"),
    ("N", "Image flipping", "Click flip buttons before uploading an image", "Buttons appear disabled, greyed out, or unclickable", "TBD", "Not Executed", "NA"),
    ("N", "Image flipping", "Rapid double-click on vertical flip", "UI does not freeze and evaluates both flips natively", "TBD", "Not Executed", "NA"),
    ("P", "Image flipping", "Click 'Flip Vertical' button on a 1x1 image", "UI updates normally handling extreme bounds correctly", "TBD", "Not Executed", "NA"),
]

file_path = Path("Manual Test Cases for Option 2.csv").resolve()

with file_path.open("w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(headers)
    
    pos_case_id = 1
    neg_case_id = 1
    
    for row in raw_data:
        t_type = row[0]
        if t_type == "P":
            tc_id = f"Pos_{pos_case_id:04d}"
            pos_case_id += 1
        else:
            tc_id = f"Neg_{neg_case_id:04d}"
            neg_case_id += 1
        
        writer.writerow([tc_id] + list(row[1:]))

print(f"CSV successfully created at: {file_path}")