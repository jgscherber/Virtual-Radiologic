Attribute VB_Name = "Module1"
Sub Information()
Attribute Information.VB_ProcData.VB_Invoke_Func = "R\n14"
Dim counter As Long
Dim template_row As Long
Dim Working As Worksheet
Dim WorkingName As String
Dim LastWorkRow As Long
Dim HeaderRows As Collection
Dim ReqE, TotalE As Long



'Dim LegalHeaderRow, StateHeaderRow, CertHeaderRow As Long
'Dim VerifCertHeaderRow, AddHeaderRow, EduCertHeaderRow As Long
'Dim PremedHeaderRow, MedHeaderRow, PostGradHeaderRow As Long
'Dim ExamHeaderRow, WorkHeaderRow, HospHeaderRow As Long
'Dim ReportHeaderRow, MilHeaderRow, RefHeaderRow, PointHeaderRow As Long

' check for summary worksheet, if exists, deletes it

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
'   Iterate over worksheets (change 2 back to counter when done testing)
For i = 16 To 17
    ReqE = 0
    TotalE = 0
'   Clear dict to be blank for new worksheet
    Set HeaderRows = New Collection
'   Get worksheet from summary page, set to "Working"
    WorkingName = Sheets("Summary").Range("A" & CStr(i)).Value
    Set Working = Sheets(WorkingName)
'   Determine last row of data
    LastWorkRow = Working.UsedRange.SpecialCells(xlCellTypeLastCell).Row
'   Determine row numbers of all header columns
    For j = 1 To LastWorkRow
        If Working.Range("A" & CStr(j)).Value Like "*Legal Documents*" Then
            HeaderRows.Add j, "LegalHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value = "State Licenses" Then
            HeaderRows.Add j, "StateHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value = "Certificates" Then
            HeaderRows.Add j, "CertHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Verification of Certificates*" Then
            HeaderRows.Add j, "VerifCertHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Additional Information/Documents*" Then
            HeaderRows.Add j, "AddHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Education Certificates*" Then
            HeaderRows.Add j, "EduCertHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value = "Premed" Then
            HeaderRows.Add j, "PremedHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value = "  Medical School " Then
            HeaderRows.Add j, "MedHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Post Graduate Training*" Then
            HeaderRows.Add j, "PostGradHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Exam Records*" Then
            HeaderRows.Add j, "ExamHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Work History*" Then
            HeaderRows.Add j, "WorkHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Hospital Affiliations*" Then
            HeaderRows.Add j, "HospHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Insurance (Past 10 years)*" Then
            HeaderRows.Add j, "HeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Reports/Malpractice*" Then
            HeaderRows.Add (j + 1), "ReportHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Military*" Then
            HeaderRows.Add j, "MilHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value = "References" Then
            HeaderRows.Add j, "RefHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Additional Items - Point Person*" Then
            HeaderRows.Add j, "PointAddPHeaderRow"
        End If
    Next j
    
'   Legal: between "Legal Documents" and "State Licenses"
    TotalE = HeaderRows("StateHeaderRow") - (HeaderRows("LegalHeaderRow") + 1)
    For leg = HeaderRows("LegalHeaderRow") + 1 To HeaderRows("StateHeaderRow") - 1
        If (IsEmpty(Working.Range("B" & CStr(leg)).Value) = False) Or Working.Range("B" & CStr(leg)).Interior.ColorIndex = 1 Then
            ReqE = ReqE + 1
        End If
    Next leg
    Sheets("Summary").Range("B" & CStr(i)).Value = Round((ReqE / TotalE) * 100)
        
        
'   State Licenses: variable between rows "State Licenses" and "Certificates"
'   Certificates: variable between rows "Certificates" and "Verifications of Certificates"
'Dim Item As Variant
'For Each Item In HeaderRows
'    MsgBox CStr(Item)
'Next Item
    
Next i
Columns("A:D").AutoFit


End Sub
