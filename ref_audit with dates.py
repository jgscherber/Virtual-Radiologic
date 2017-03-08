import openpyxl as xl
import sys, datetime
from RadInfo import *



# use pyinstaller "ref_audit with dates.py" -F -n ref_audit.exe
# -- Function Code --            
##try:
       
audit = unicode(open("note.txt","r").read(), errors='replace')
audit = audit.split('\n')[1:]


r1_loc = [] # name
r2_loc = [] # file date 
r3_loc = [] # close date
r4_loc = [] # case name
r5_loc = [] # method closed
r6_loc = [] # investigations 
r7_loc = [] # disciplinary action

# find the start and end of each section
for line in range(len(audit)):
    audit[line] = audit[line].rstrip(' ')
    
    if audit[line] == 'R1':
        r1_loc.append(line) # holds the locations of all R1s
    elif audit[line] == 'R2':
        r2_loc.append(line)
    elif audit[line] == 'R3':
        r3_loc.append(line)
    elif audit[line] == 'R4':
        r4_loc.append(line)
    elif audit[line] == 'R5':
        r5_loc.append(line)
    elif audit[line] == 'R6':
        r6_loc.append(line)
    elif audit[line] == 'R7':
        r7_loc.append(line)
allRads = []
total_lines = 0
numberRads = len(r1_loc)
for rad in range(numberRads): # len of R1 is number of rads

    # get each section of information
    current = RadInfo(audit[(r1_loc[rad]+1)]) # name
    current.setFileDates(audit[(r2_loc[rad]+1):r3_loc[rad]]) # 1st: file date
    current.setClosedDates(audit[(r3_loc[rad]+1):r4_loc[rad]]) # 2nd: close date
    current.setCaseNames(audit[(r4_loc[rad]+1):r5_loc[rad]]) # 3rd: case name
    current.setMethodCloseds(audit[(r5_loc[rad]+1):r6_loc[rad]]) # 4th: method closed
    current.setInvestigations(audit[(r6_loc[rad]+1):r7_loc[rad]]) # 5th: investigations
    
    if rad == len(r1_loc)-1: # if last rad entry
        current.setDisciplinaryActions(audit[(r7_loc[rad]+1):len(audit)]) # 6th: disciplinary action
    else:
        current.setDisciplinaryActions(audit[(r7_loc[rad]+1):r1_loc[rad+1]])
     


    # add to full rad list
    allRads.append(current)
    total_lines = total_lines + current.longest
# end for (next rad)

# put into WB
    
startingRow = 2
wb = xl.Workbook()
sheet = wb.active
sheet['A1'].value = 'Physician'
sheet['B1'].value = 'Date Filed'
sheet['C1'].value = 'Date Closed'
sheet['D1'].value = 'Malpractice Case'
sheet['E1'].value = 'Malpractice Status'
sheet['F1'].value = 'Investigations'
sheet['G1'].value = 'Disciplinary Actions'

for item in range(numberRads):

    for row in range(startingRow, startingRow + allRads[item].longest):
        sheet['A{0}'.format(row)].value = allRads[item].name
        if (row - startingRow) < len(allRads[item].fileDates):
            sheet['B{0}'.format(row)].value = allRads[item].fileDates[row - startingRow]
            sheet['B{0}'.format(row)].number_format = 'm/d/yyyy'
        if (row - startingRow) < len(allRads[item].closedDates):
            sheet['C{0}'.format(row)].value = allRads[item].closedDates[row - startingRow]
            sheet['C{0}'.format(row)].number_format = 'm/d/yyyy'
        if (row - startingRow) < len(allRads[item].caseNames):
            sheet['D{0}'.format(row)].value = allRads[item].caseNames[row - startingRow]
        if (row - startingRow) < len(allRads[item].methodCloseds):
            sheet['E{0}'.format(row)].value = allRads[item].methodCloseds[row - startingRow]
        if (row - startingRow) < len(allRads[item].investigations):
            sheet['F{0}'.format(row)].value = allRads[item].investigations[row - startingRow]
        if (row - startingRow) < len(allRads[item].disciplinaryActions):
            sheet['G{0}'.format(row)].value = allRads[item].disciplinaryActions[row - startingRow]
               
    startingRow += allRads[item].longest
# formatting
sheet.column_dimensions['A'].width = 35
sheet.column_dimensions['B'].width = 35
sheet.column_dimensions['C'].width = 35
sheet.column_dimensions['D'].width = 35
sheet.column_dimensions['E'].width = 35
sheet.column_dimensions['F'].width = 35
sheet.column_dimensions['G'].width = 35
sheet.auto_filter.ref = "A1:G{0}".format(startingRow-1)
sheet.freeze_panes = 'A2'

wb.save('Audit_Format.xlsx')
##except Exception as e:
##    print("Error:  {0}\nLine:   {1}".format(e, sys.exc_info()[2].tb_lineno))
##    raw_input("Press enter to continue...")
