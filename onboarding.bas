Attribute VB_Name = "Module1"
Sub Information()
Attribute Information.VB_ProcData.VB_Invoke_Func = "R\n14"
Dim counter As Long
Dim template_row As Long
Dim working As Worksheet
Dim WorkingName As String
Dim LastWorkRow As Long




' ___creating list of worksheet names___
counter = Worksheets.Count
Sheets.Add After:=Sheets(counter)
Sheets(counter + 1).Name = "Summary"
Sheets("Summary").Range("A1").Value = "Physicians"
Sheets("Summary").Range("B1").Value = "% Requested"
Sheets("Summary").Range("C1").Value = "% Received"
Sheets("Summary").Range("D1").Value = "% Uploaded"
For i = 1 To counter
    If Sheets(i).Name <> "Template" Then
        Sheets("Summary").Range("A" & CStr(i + 1)).Value = Sheets(i).Name
    Else
        template_row = i + 1
    End If
Next i
Sheets("Summary").Rows(template_row).Delete
'   For sheets(2-counter):
For i = 2 To counter
    WorkingName = Sheets("Summary").Range("A" & CStr(i)).Value
    Set working = Sheets(WorkingName)
'   First determin location of all header columns
    LastWorkRow = CStr(working.UsedRange.SpecialCells(xlCellTypeLastCell).Row)
    
'   Legal: between "Legal Documents" and "State Licenses"
'   State Licenses: variable between rows "State Licenses" and "Certificates"
'   Certificates: variable between rows "Certificates" and "Verifications of Certificates"

    
Next i
Columns("A:D").AutoFit

End Sub
