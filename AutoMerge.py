import os
from datetime import datetime
from mailmerge import MailMerge
from docx2pdf import convert
import csv

TEMPLATE_DOCUMENT = "./Input/CertTemplateW25.docx"
LIST_CSV = "./Input/AriSaltarelli-W25-ClassList - McDonald.csv"
OUTPUT_FOLDER  = "./Output/"

# Create the output folder if it doesn't exist
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

data = []

with open(LIST_CSV, newline='', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        print(row)
        data.append({
            'Student_Name': row.get('Student Name', ''),
            'Rank': row.get('Rank', ''),
            'Test_Date': row.get('Test Date', ''),
            'Instructor_Name': row.get('Instructor Name', ''),
            'Instructor_Rank': row.get('Instructor Rank', '')
        })

with MailMerge(TEMPLATE_DOCUMENT) as document:
    document.merge_pages(data)
    merged_docx_path = os.path.join(OUTPUT_FOLDER, "MergedResult.docx")
    document.write(merged_docx_path)

merged_pdf_path = os.path.join(OUTPUT_FOLDER, "MergedResult.pdf")
try:
    convert(merged_docx_path, merged_pdf_path)
    print("Merge complete")
except Exception as e:
    print("Error converting to PDF: ", e)


