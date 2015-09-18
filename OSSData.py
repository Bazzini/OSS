from xlrd import open_workbook
# tkinter allows GUI-based file dialog. Spelling is important. Not Tkinter or tKinter.
from tkinter.filedialog import askopenfilename  # This opens a dialog box.

__author__ = 'Other Laptop'

filename = askopenfilename()
book = open_workbook(filename)
sheet = book.sheet_by_index(0)

# read header values into the list
keys = [sheet.cell(0, col_index).value for col_index in range(sheet.ncols)]

student_list = []
for row_index in range(1, sheet.nrows):
    d = {keys[col_index]: sheet.cell(row_index, col_index).value
         for col_index in range(sheet.ncols)}
    student_list.append(d)

count = 0
studwithteach = []
for student in student_list:
    if student['StudentTeacherID'] == 'Rosemarie Klimasko':
        studwithteach.append(student['FirstName'] + ' ' + student['LastName'])

parentwithteach = []
for student in student_list:
    if student['StudentTeacherID'] == 'Rosemarie Klimasko':
        parentwithteach.append(student['P1Email'])
        if student['P2Email'] and student['P2Email'].strip():
            parentwithteach.append(student['P2Email'])
pass