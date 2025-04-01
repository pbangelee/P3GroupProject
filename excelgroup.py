#Group 10:
#Kenton Harris, Nathan Saez, Aaron Shumway, Angelee Marshall, Jennica Olsen
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font


myWorkbook = openpyxl.load_workbook("Poorly_Organized_Data_1.xlsx")

currWs = myWorkbook.active

for row in currWs.iter_rows(min_row=2, values_only=True, min_col=1, max_col=3):
    if row[0] not in myWorkbook.sheetnames:
        myWorkbook.create_sheet(row[0])
        myWorkbook.active = myWorkbook[str(row[0])]
        currWs = myWorkbook.active
        currWs["A1"] = "Last Name"
        currWs["B1"] = "First Name"
        currWs["C1"] = "Student ID"
        currWs["D1"] = "Grade"
        myWorkbook.active = myWorkbook["Grades"]
        currWs = myWorkbook.active

for row in currWs.iter_rows(min_row=2, values_only=True, min_col=1, max_col=3):
    studentInfo = row[1].split("_")
    studentInfo.append(row[2])
    myWorkbook.active = myWorkbook[row[0]]
    myWorkbook.active.append(studentInfo)

# Create new worksheets for each class (e.g., a sheet for Algebra, a sheet for Calculus, etc.)

#currWs = myWork
#myWorksheet.create_sheet.title("Alegebra")
#myWorksheet.create_sheet("Calculus")
#myWorksheet.create_sheet("Stats")
#myWorksheet.create_sheet("Geometry")
#myWorksheet.create_sheet("Trigonometry")




# In each sheet, create columns for last name, first name, student ID, and grade with the student data for that class placed there.

#currWs["A12"] = "Lastt Name"

#currWs[] = "First Name"

#currWs[] = "Student ID"

#currWs[] = "Grade"

# A filter should be placed over the 4 aforementioned columns in each sheet.
# Additionally, each sheet should have some simple summary information about each class using functions in columns F (the titles) and G (the data). It should show:
# The highest grade, The lowest grade, The mean grade, The median grade, The number of students in the class
# Some simple formatting (bolding headers) and changing the width of the columns.
# The width of the columns for A,B,C,D,F,G must each be set to the number of characters in the header + 5. 
# For example the column D header is “Grade” which has 5 characters, so the width of column D should be 10, etc.
# Save the results as a new Excel file named “formatted_grades.xlsx”





myWorkbook.save(filename="FixedSheet.xlsx")
myWorkbook.close()