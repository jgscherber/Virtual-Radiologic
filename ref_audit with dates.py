import openpyxl as xl
import sys, datetime



# use pyinstaller "ref_audit with dates.py" -F -n ref_audit.exe
# -- Function Code --            
try:
       
    audit = unicode(open("note.txt","r").read(), errors='replace')
    audit = audit.split('\n')[1:]
    

    r1_loc = [] # name
    r2_loc = [] # case name 
    r3_loc = [] # file date
    r4_loc = [] # close date
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
    total = []
    total_lines = 0
    
    for rad in range(len(r1_loc)): # len of R1 is number of rads

        # get each section of information
        current = []
        current.append(audit[(r1_loc[rad]+1)]) # 0th element is name
        current.append(audit[(r2_loc[rad]+1):r3_loc[rad]]) # 1st: case name
        current.append(audit[(r3_loc[rad]+1):r4_loc[rad]]) # 2nd: file date
        current.append(audit[(r4_loc[rad]+1):r5_loc[rad]]) # 3rd: close date
        current.append(audit[(r5_loc[rad]+1):r6_loc[rad]]) # 4th: method closed
        current.append(audit[(r6_loc[rad]+1):r7_loc[rad]]) # 5th: investigations
        
        if rad == len(r1_loc)-1: # if last rad entry
            current.append(audit[(r7_loc[rad]+1):len(audit)]) # 6th: disciplinary action
        else:
            current.append(audit[(r7_loc[rad]+1):r1_loc[rad+1]])
        lastIndex = len(current) - 1
        possibleBlank = len(current[lastIndex]) - 1
        
        # remove empty and blank lines
        bad_entry = [None, u"", u"N/A", u"NA", u"N/A ", u"No", u"No "]
        while possibleBlank > -1:
            
            if current[lastIndex][possibleBlank] in bad_entry:
                del current[lastIndex][possibleBlank]

            possibleBlank -= 1
        possibleBlank = len(current[lastIndex-1]) - 1
        while possibleBlank > -1:
            if current[lastIndex-1][possibleBlank] in bad_entry:
                del current[lastIndex-1][possibleBlank]

            possibleBlank -= 1
        
        # find the number of names that will be needed
        longest = sorted([len(current[1]), len(current[2]), len(current[3]),
                          len(current[4]),len(current[5]),len(current[6])])[-1]
        current.append(longest) # last entry in current
    

    
        # add to full rad list
        total.append(current)
        total_lines = total_lines + longest
    # end for (next rad)

    # put into WB
        
    working_row = 2
    wb = xl.Workbook()
    sheet = wb.active
    sheet['A1'].value = 'Physician'
    sheet['B1'].value = 'Date Filed'
    sheet['C1'].value = 'Date Closed'
    sheet['D1'].value = 'Malpractice Case'
    sheet['E1'].value = 'Malpractice Status'
    sheet['F1'].value = 'Investigations'
    sheet['G1'].value = 'Disciplinary Actions'
    lengthLoc = len(total[0]) - 1 
    for item in range(len(total)):

        for row in range(working_row, working_row + total[item][lengthLoc]):
            sheet['A{0}'.format(row)].value = total[item][0]
            if (row - working_row) < len(total[item][1]):
                if(total[item][1][row - working_row] != ''):
                    sheet['B{0}'.format(row)].value = datetime.datetime.strptime( \
                        total[item][1][row - working_row], '%m/%d/%Y')
                    sheet['B{0}'.format(row)].number_format = 'm/d/yyyy'
                else:
                    sheet['B{0}'.format(row)].value = total[item][1][row - working_row]
            if (row - working_row) < len(total[item][2]):
                if(total[item][2][row - working_row] != ''):
                    sheet['C{0}'.format(row)].value = datetime.datetime.strptime( \
                        total[item][2][row - working_row], '%m/%d/%Y')
                    sheet['C{0}'.format(row)].number_format = 'm/d/yyyy'
                else:
                    sheet['C{0}'.format(row)] = total[item][2][row - working_row]
            if (row - working_row) < len(total[item][3]):
                sheet['D{0}'.format(row)].value = total[item][3][row - working_row]
            if (row - working_row) < len(total[item][4]):
                sheet['E{0}'.format(row)].value = total[item][4][row - working_row]
            if (row - working_row) < len(total[item][5]):
                sheet['F{0}'.format(row)].value = total[item][5][row - working_row]
            if (row - working_row) < len(total[item][6]):
                sheet['G{0}'.format(row)].value = total[item][6][row - working_row]
                   
        working_row += total[item][lengthLoc]
    # formatting
    sheet.column_dimensions['A'].width = 35
    sheet.column_dimensions['B'].width = 35
    sheet.column_dimensions['C'].width = 35
    sheet.column_dimensions['D'].width = 35
    sheet.column_dimensions['E'].width = 35
    sheet.column_dimensions['F'].width = 35
    sheet.column_dimensions['G'].width = 35
    sheet.auto_filter.ref = "A1:G{0}".format(working_row-1)
    sheet.freeze_panes = 'A2'

    wb.save('Audit_Format.xlsx')
except Exception as e:
    print("Error:  {0}\nLine:   {1}".format(e, sys.exc_info()[2].tb_lineno))
    raw_input("Press enter to continue...")
