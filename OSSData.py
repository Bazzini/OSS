from xlrd import open_workbook
import os
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
    d['FullName'] = d['FirstName'] + ' ' + d['LastName']
    student_list.append(d)

studwithteach = []
for student in student_list:
    if student['StudentTeacherID'] == 'Rosemarie Klimasko':
        studwithteach.append(student['FirstName'] + ' ' + student['LastName'])

parentwithteach = []
for student in student_list:
    if student['StudentTeacherID'] == 'Rosemarie Klimasko':
        parentwithteach.append(student['P1Email'])
        if student['P2Email'] and student['P2Email'].strip():  # if P2Email not blank
            parentwithteach.append(student['P2Email'])

teacherlist = []
for student in student_list:
    teacherlist.append(student['StudentTeacherID'])

teachers = set(teacherlist)

# Group class

filename = askopenfilename(title='Pick a Group Class file')
filedir, dfilename = os.path.split(filename)  # this extracts the directory and the specific  name used.
dfilename, extension = os.path.splitext(dfilename)  # this extracts the extension used.

group_files = [f for f in os.listdir(filedir) if f.endswith(extension)]

for groupfilename in group_files:
    groupbook = open_workbook(groupfilename)
    groupsheet = groupbook.sheet_by_index(0)
    firstcol = groupsheet.col(0)

    # Find the Group Class column
    classrow = [a for a, b in enumerate(firstcol) if 'Group Class' in firstcol[a].value][0]
    groupclass = groupsheet.cell(classrow, 1).value
    # Find the Teacher column
    teacherrow = [a for a, c in enumerate(firstcol) if 'Teacher' in firstcol[a].value][0]
    teacher = groupsheet.cell(teacherrow, 1).value
    # Find where the headers are for all students in list
    headerrow = [a for a, d in enumerate(firstcol) if 'LastName' in firstcol[a].value][0]

    keys = [sheet.cell(headerrow, col_index).value for col_index in range(sheet.ncols)]

    group_list = []

    for row_index in range(headerrow + 1, sheet.nrows):
        d = {keys[col_index]: sheet.cell(row_index, col_index).value
             for col_index in range(sheet.ncols)}
        d['FullName'] = d['FirstName'] + ' ' + d['LastName']
        d['InstrClass'] = groupclass
        d['GroupInstrTeacher'] = teacher
        group_list.append(d)

pass
