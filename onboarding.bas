Attribute VB_Name = "Module1"
Public Function TypeCounter(HeaderRows, Working, TopHeaderRow, BottomHeaderRow, ReqRcvUp)
Dim Total As Long
Dim Req As Long
Dim Percent As Long

Req = 0
Total = 0
'Total = HeaderRows(BottomHeaderRow) - (HeaderRows(TopHeaderRow) + 1)
For leg = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
    If Working.Range(ReqRcvUp & CStr(leg)).Interior.ColorIndex <> 1 Then
        Total = Total + 1
    End If
    If (IsEmpty(Working.Range(ReqRcvUp & CStr(leg)).Value) = False) And (Working.Range(ReqRcvUp & CStr(leg)).Interior.ColorIndex <> 1) Then
        Req = Req + 1
    End If
Next leg
If Total = 0 Then
    TypeCounter = 0
Else
    TypeCounter = Round((Req / Total) * 100)
End If


End Function


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
Sheets("Summary").Range("B1").Value = "% Legal Rqstd"
Sheets("Summary").Range("C1").Value = "% Legal Rcvd"
Sheets("Summary").Range("D1").Value = "% Legal Uploaded"
Sheets("Summary").Range("E1").Value = "% State Lic Rqstd"
Sheets("Summary").Range("F1").Value = "% State Lic Rcvd"
Sheets("Summary").Range("G1").Value = "% State Lic Uploaded"
Sheets("Summary").Range("H1").Value = "% Cert Rqstd"
Sheets("Summary").Range("I1").Value = "% Cert Rcvd"
Sheets("Summary").Range("J1").Value = "% Cert Lic Uploaded"
Sheets("Summary").Range("K1").Value = "% Ver of Cert Rqstd"
Sheets("Summary").Range("L1").Value = "% Ver of Cert Rcvd"
Sheets("Summary").Range("M1").Value = "% Ver of Cert Uploaded"
For i = 1 To counter
    If Sheets(i).Name <> "Template" Then
        Sheets("Summary").Range("A" & CStr(i + 1)).Value = Sheets(i).Name
    Else
        template_row = i + 1
    End If
Next i
Sheets("Summary").Rows(template_row).Delete
'   Iterate over worksheets (change 2 back to counter when done testing)
For i = 2 To counter
    
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
    Sheets("Summary").Range("B" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "LegalHeaderRow", "StateHeaderRow", "B")
    Sheets("Summary").Range("C" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "LegalHeaderRow", "StateHeaderRow", "C")
    Sheets("Summary").Range("D" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "LegalHeaderRow", "StateHeaderRow", "D")
'   State Licenses: variable between rows "State Licenses" and "Certificates"
    Sheets("Summary").Range("E" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "StateHeaderRow", "CertHeaderRow", "B")
    Sheets("Summary").Range("F" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "StateHeaderRow", "CertHeaderRow", "C")
    Sheets("Summary").Range("G" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "StateHeaderRow", "CertHeaderRow", "D")
'   Certificates: variable between rows "Certificates" and "Verifications of Certificates"
    Sheets("Summary").Range("H" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "CertHeaderRow", "VerifCertHeaderRow", "B")
    Sheets("Summary").Range("I" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "CertHeaderRow", "VerifCertHeaderRow", "C")
    Sheets("Summary").Range("J" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "CertHeaderRow", "VerifCertHeaderRow", "D")
'   Verification of Certificates
    Sheets("Summary").Range("K" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", "B")
    Sheets("Summary").Range("L" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", "C")
    Sheets("Summary").Range("M" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", "D")
'   Addition Information/Documents and Additional Items - Point Person


    
Next i
Columns("A").AutoFit
Rows(1).RowHeight = 30
Columns("B:Z").ColumnWidth = 9
Range("A1:Z1").WrapText = True


End Sub
