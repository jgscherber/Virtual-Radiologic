Attribute VB_Name = "Module1"
Sub Information()
Attribute Information.VB_ProcData.VB_Invoke_Func = "R\n14"
Dim counter As Long
Dim template_row As Long


' ___creating list of worksheet names___
counter = Worksheets.Count
Sheets.Add After:=Sheets(counter)
Sheets(counter + 1).Name = "Summary"
Sheets("Summary").Range("A1").Value = "Physicians"

For i = 1 To counter
    If Sheets(i).Name <> "Template" Then
        Sheets("Summary").Range("A" & CStr(i + 1)).Value = Sheets(i).Name
    Else
        template_row = i + 1
    End If
Next i
Sheets("Summary").Rows(template_row).Delete

End Sub
