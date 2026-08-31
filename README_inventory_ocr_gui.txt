
JPG OCR to Inventory Template

Purpose
This desktop Python app reads JPG images, extracts OCR text, and writes that text into the Description column of an Excel inventory import template.

How it works
- Loads your existing .xlsx template
- Uses the Data sheet by default
- Looks for the header named Description on row 1
- Writes one OCR result per image into successive rows starting at row 2
- Preserves the rest of the workbook and keeps the Instructions sheet

Default mapping for your sample template
- Sheet: Data
- Header row: 1
- Start row: 2
- Target column header: Description

Files included
- batch_jpg_to_inventory_description_gui.py
- requirements_inventory_ocr.txt

Requirements
1. Python 3.10 or newer
2. Tesseract OCR installed on your computer
3. Python packages in requirements_inventory_ocr.txt

Install packages
pip install -r requirements_inventory_ocr.txt

Install Tesseract
Windows:
- Install Tesseract OCR
- Typical path:
  C:\Program Files\Tesseract-OCR\tesseract.exe

Mac:
brew install tesseract

Run
Windows:
python batch_jpg_to_inventory_description_gui.py

Mac:
python3 batch_jpg_to_inventory_description_gui.py

Basic use
1. Pick your Excel template workbook
2. Pick a folder of JPG files or choose files manually
3. Pick the output workbook name
4. If needed, set the Tesseract executable path
5. Click Create Filled Workbook

Notes
- OCR text is written into the Description column only
- One image = one row
- If there are more images than existing blank rows, the app inserts more rows
- Best results come from clean, straight, high contrast images

If OCR quality is poor
- Try PSM 6 first
- Try cleaner images
- Try flattening line breaks if you want a single paragraph in each cell
