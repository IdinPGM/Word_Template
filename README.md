**Read Me**

Purpose: This script is designed to efficiently create multiple customized Word documents based on a single template and data stored in an Excel file.

**Prerequisites:**

Python 3.x installed.
The following Python libraries installed:
Bash

pip install python-docx pandas openpyxl
A Word document template named template.docx in the same directory as the script. Placeholders for data from the Excel file should be enclosed in double curly braces (e.g., {{เรื่อง}}, {{เรียน}}).
An Excel file named input.xlsx in the same directory as the script. The first row of the Excel file should contain headers that match the placeholders in the Word template.

**How to Use:**
1. Ensure that the template.docx and input.xlsx files are in the same directory as the Python script.
2. Run the Python script (fill_word_template.py).
3. The script will read the data from input.xlsx and generate new Word documents (e.g., output_row1.docx, output_row2.docx) in the same directory, with the placeholders filled with the corresponding data.

**Output:**
The script will generate one Word document for each row in the input.xlsx file, with the placeholders in the template.docx replaced by the data from that row. The output files will be named with a prefix (default: output.docx) followed by _row and the row number from the Excel file.
