import os
from datetime import datetime
from mailmerge import MailMerge
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
        data.append({
            'Student_Name': row.get('Student_Name', ''),
            'Rank': row.get('Rank', ''),
            'Test_Date': row.get('Test_Date', ''),
            'Instructor_Name': row.get('Instructor_Name', ''),
            'Instructor_Rank': row.get('Instructor_Rank', '')
        })

with MailMerge(TEMPLATE_DOCUMENT) as document:
    document.merge_pages(data)
    merged_docx_path = "./Output/MergedResult.docx"
    document.write(merged_docx_path)

    
