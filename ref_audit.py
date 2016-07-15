import openpyxl as xl

       
audit = unicode(open("note2.txt","r").read(), errors='replace')
audit = audit.split('\n')[1:]

r1_loc = []
r2_loc = []
r3_loc = []
r4_loc = []
r5_loc = []
for line in range(len(audit)):
    if audit[line] == 'R1':
        r1_loc.append(line)
    elif audit[line] == 'R2':
        r2_loc.append(line)
    elif audit[line] == 'R3':
        r3_loc.append(line)
    elif audit[line] == 'R4':
        r4_loc.append(line)
    elif audit[line] == 'R5':
        r5_loc.append(line)
total = []
total_lines = 0
bad_entry = [None, u"", u"N/A", u"NA", u"N/A ", u"No", u"No "]
for loc in range(len(r1_loc)):
    current = []
    current.append(audit[(r1_loc[loc]+1)])
    current.append(audit[(r2_loc[loc]+1):r3_loc[loc]])
    current.append(audit[(r3_loc[loc]+1):r4_loc[loc]])
    current.append(audit[(r4_loc[loc]+1):r5_loc[loc]])
    if loc == len(r1_loc)-1:
        current.append(audit[(r5_loc[loc]+1):len(audit)])
    else:
        current.append(audit[(r5_loc[loc]+1):r1_loc[loc+1]])
    blank = len(current[4]) -1
    while blank > -1:
        
        if current[4][blank] in bad_entry:
            del current[4][blank]
##        else:
##            break
        blank -= 1
    blank = len(current[3]) -1
    while blank > -1:
        if current[3][blank] in bad_entry:
            del current[3][blank]
##        else:
##            break
        blank -= 1   
    longest = sorted([len(current[1]), len(current[2]), len(current[3]),
                      len(current[4])])[-1]
    current.append(longest)
    total.append(current)
    total_lines = total_lines + longest

# put into WB
    
working_row = 2
wb = xl.Workbook()
sheet = wb.active
sheet['A1'].value = 'Physician'
sheet['B1'].value = 'Malpractice'
sheet['C1'].value = 'Status'
sheet['D1'].value = 'Investigations'
sheet['E1'].value = 'Disciplinary'

for item in range(len(total)):

    for row in range(working_row, working_row + total[item][5]):
        sheet['A{0}'.format(row)].value = total[item][0]
        if (row - working_row) < len(total[item][1]):
            sheet['B{0}'.format(row)].value = total[item][1][row - working_row]
        if (row - working_row) < len(total[item][2]):
            sheet['C{0}'.format(row)].value = total[item][2][row - working_row]
        if (row - working_row) < len(total[item][3]):
            sheet['D{0}'.format(row)].value = total[item][3][row - working_row]
        if (row - working_row) < len(total[item][4]):
            sheet['E{0}'.format(row)].value = total[item][4][row - working_row]
               
    working_row = working_row + total[item][5]

# clean out blank lines
print(working_row)

wb.save('Audit_Format.xlsx')
