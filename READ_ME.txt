=====================
How to run the program:

1. Write the names of students and class inside 'StudentWorkshop_SampleExcelSheet'
2. Ensure that Excel sheet is closed before running the Python program.
3. Run 'ExcelHelperTool.py' in your terminal/IDE/tool of choice.
4. Re-open the Excel sheet to see the update.

Note: 
- Numbers of students will be taken into account for each calculation, and there is no need
to remove existing workshop information inside the excel sheet.

=====================
Changing variables in Excel and Python

1. The program is currently linked to the name of the Excel file 
(StudentWorkshop_SampleExcelSheet.xlsx) and the sheet name (Gesamtübersicht). If they must 
be changed, make sure to update 'filename' and 'sheetname' inside the python script.

2. Changes to which column and row the student names/class start must be updated in the 
python script. The default is column A,B, C and row 4. 

3. You could change the name of the workshops inside the text file 'workshop_names'. However,
beware that the workshop names on the right side of the Excel sheet does not reflect this 
change. 

4. The order of workshop names in 'workshop_names.txt' is very important, as for now, the 
order of class in the overview in Excel only work for the default order inside the text file.