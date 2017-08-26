# Turn the application spreadsheet into a printable .docx file
# Requirements: openpyxl (python-pptx), docx (python-docx) (can pip install)
## Requires a spreadsheet with name file_name.xlsx (specified below) and
## sheet name sheet_name (specified below). The first row should be a header
## with names such as First Name, Last Name, Why do you want to do BOLT, etc
## To include the first N columns, set num_relevant_cols to N (specified below)
# any string decoding errors will need to be bugfixed by editing the code
# run with python application-cleaner.py
## first column CANNOT BE BLANK otherwise the entire row will be skipped
## (this is to deal with Excel automatically adding blank rows at the bottom)

import openpyxl
from docx import Document

# set parameters here
file_name = '2017-application.xlsx'
sheet_name = 'All Accepted'
num_relevant_cols = 10

# load the application spreadsheet
ss = openpyxl.load_workbook(file_name)
sheet = ss.get_sheet_by_name(sheet_name)
vals = list(sheet.values)

# create the document
doc = Document()

# for each row...
for val in vals[1:]:
    # if the first column in this row is blank, skip it entirely
    if val[0] is None:
        continue

    string = ''
    for i in range(num_relevant_cols):
        # a is corresponding header tag
        a = vals[0][i]
        # b is the response
        b = val[i]
        string += '%s: %s\n' % (a, b)

    doc.add_paragraph(string)
    doc.add_page_break()

doc.save('Printable Apps.docx')
