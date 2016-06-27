Attribute VB_Name = "Module1"
Public Function TypeCounter(HeaderRows, Working, TopHeaderRow, BottomHeaderRow, ReqRcvUp)
Dim Total As Long
Dim Completed As Long
Dim Percent As Long

Completed = 0
Total = 0
'Total = HeaderRows(BottomHeaderRow) - (HeaderRows(TopHeaderRow) + 1)
For leg = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
    If Working.Range(ReqRcvUp & CStr(leg)).Interior.ColorIndex <> 1 Then
        Total = Total + 1
    End If
    If ((IsEmpty(Working.Range(ReqRcvUp & CStr(leg)).Value) = False) And (Working.Range(ReqRcvUp & CStr(leg)).Interior.ColorIndex <> 1)) _
    Or Working.Range(ReqRcvUp & CStr(leg)).Interior.ColorIndex = 15 Then
        Completed = Completed + 1
    End If
Next leg
If Total = 0 Then
    TypeCounter = 100
Else
    TypeCounter = Round((Completed / Total) * 100)
End If


End Function

Public Function MissingItems(HeaderRows, Working, TopHeaderRow, BottomHeaderRow, MissingRow, i)
For mrow = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
    If IsEmpty(Working.Range("C" & mrow).Value) And Working.Range("C" & mrow).Interior.ColorIndex <> 1 _
    And Working.Range("C" & mrow).Interior.ColorIndex <> 15 Then
        Sheets("Missing Items").Cells(MissingRow, i - 1).Value = Working.Range("A" & mrow).Value
        MissingRow = MissingRow + 1
    End If
Next mrow

MissingItems = MissingRow
End Function

Public Function MissingItemsST(HeaderRows, Working, TopHeaderRow, BottomHeaderRow, MissingRow, i)
For mrow = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
    If IsEmpty(Working.Range("C" & mrow).Value) And Working.Range("C" & mrow).Interior.ColorIndex <> 1 _
    And Working.Range("C" & mrow).Interior.ColorIndex <> 15 Then
        If Working.Range("A" & mrow) Like "*Wallet/Wall" Then
            premrow = CStr(CInt(mrow) - 1)
            Sheets("Missing Items").Cells(MissingRow, i - 1).Value = Working.Range("A" & premrow).Value & Working.Range("A" & mrow).Value
        End If
        If Working.Range("A" & mrow) Like "*Verification" Then
            premrow = CStr(CInt(mrow) - 2)
            Sheets("Missing Items").Cells(MissingRow, i - 1).Value = Working.Range("A" & premrow).Value & Working.Range("A" & mrow).Value
        End If
        MissingRow = MissingRow + 1
    End If
Next mrow

MissingItemsST = MissingRow
End Function


Sub Information()
Attribute Information.VB_ProcData.VB_Invoke_Func = "R\n14"
'   Set variables
Dim counter As Long
Dim template_row As Long
Dim Working As Worksheet
Dim WorkingName As String
Dim LastWorkRow As Long
Dim HeaderRows As Collection
Dim HeaderNames As Variant


'   check for summary worksheet, if exists, deletes it
counter = Worksheets.Count
Application.DisplayAlerts = False
For shnum = 1 To counter - 1
    If Sheets(shnum).Name = "Summary" Then
        Sheets(shnum).Delete
    End If
Next shnum
Application.DisplayAlerts = True
counter = Worksheets.Count
Application.DisplayAlerts = False
For shnum = 1 To counter
    If Sheets(shnum).Name = "Missing Items" Then
        Sheets(shnum).Delete
    End If
Next shnum
Application.DisplayAlerts = True

'   creating list of worksheet names
counter = Worksheets.Count
Sheets.Add After:=Sheets(counter)
Sheets(counter + 1).Name = "Summary"
Sheets.Add After:=Sheets("Summary")
Sheets(counter + 2).Name = "Missing Items"

'   Create and set table headers
HeaderNames = Array( _
"Physicians", _
"% Legal Rqstd", "% Legal Rcvd", "% Legal Upload", _
"% State Lic Rqstd", "% State Lic Rcvd", "% State Lic Upload", _
"% Cert Rqstd", "% Cert Rcvd", "% Cert Upload", _
"% Additional Rqst", "% Additional Rcvd", "% Additional Upload", _
"% Education Requested", "% Education Recieved", "% Education Upload", _
"% Work Requested", "% Work Recieved", "% Work Uploaded", _
"% Affiliation Requested", "% Affiliation Recieved", "% Affiliation Uploaded", _
"% Insurance Requested", "% Insurance Recieved", "% Insurance Uploaded", _
"% Reports Requested", "% Reports Recieved", "% Reports Uploaded", _
"% Military Requested", "% Military Recieved", "% Military Uploaded", _
"% Reference Requested", "% Reference Recieved", "% Reference Uploaded", _
"% Total Requested", "% Total Recieved", "%Total Uploaded", _
"% Pending")
Sheets("Summary").Range("A1:AL1").Value = HeaderNames


'   Add physicians name to summary page
For i = 1 To counter
    If Sheets(i).Name <> "Template" Then
        Sheets("Summary").Range("A" & CStr(i + 1)).Value = Sheets(i).Name
        Sheets("Summary").Range("A" & CStr(i + 1)).Interior.ColorIndex = Sheets(i).Tab.ColorIndex
        Sheets("Missing Items").Cells(1, i).Value = Sheets(i).Name
        Sheets("Missing Items").Cells(1, i).Interior.ColorIndex = Sheets(i).Tab.ColorIndex
        If Sheets(i).Tab.ColorIndex = 1 Then
            Sheets("Missing Items").Cells(1, i).Font.Color = RGB(255, 255, 255)
            Sheets("Summary").Range("A" & CStr(i + 1)).Font.Color = RGB(255, 255, 255)
        End If
        
    Else
        template_row = i + 1
    End If
Next i
Sheets("Summary").Rows(template_row).Delete
Sheets("Missing Items").Columns(template_row - 1).EntireColumn.Delete Shift:=xlToLeft
'   Resize columns
With Sheets("Summary")
    .Columns("A").AutoFit
    .Rows(1).RowHeight = 45
    .Columns("B:AL").ColumnWidth = 12
End With
With Sheets("Summary").Range("A1:AL1")
    .WrapText = True
    .VerticalAlignment = xlTop
    .HorizontalAlignment = xlCenter
End With

'   Iterate over worksheets (should be '2 to counter')
For i = 2 To counter
    
'   Clear dict to be blank for new worksheet
    Set HeaderRows = New Collection
    MissingRow = 2
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
            HeaderRows.Add j, "InsHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Reports/Malpractice*" Then
            HeaderRows.Add (j + 1), "ReportHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Military*" Then
            HeaderRows.Add j, "MilHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "References*" And _
        Working.Range("A" & CStr(j)).Font.Bold = True Then
            HeaderRows.Add j, "RefHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Additional Items*" Then
            HeaderRows.Add j, "PointAddPHeaderRow"
        End If
    Next j
    HeaderRows.Add (HeaderRows("PointAddPHeaderRow") + 4), "LastEmptyRow"
'   Fill out missing row spreadsheet

    MissingRow = MissingItems(HeaderRows, Working, "LegalHeaderRow", "StateHeaderRow", MissingRow, i)
    MissingRow = MissingItemsST(HeaderRows, Working, "StateHeaderRow", "CertHeaderRow", MissingRow, i)
    MissingRow = MissingItems(HeaderRows, Working, "CertHeaderRow", "VerifCertHeaderRow", MissingRow, i)
    MissingRow = MissingItems(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", MissingRow, i)
    MissingRow = MissingItems(HeaderRows, Working, "AddHeaderRow", "EduCertHeaderRow", MissingRow, i)
    MissingRow = MissingItems(HeaderRows, Working, "PointAddPHeaderRow", "LastEmptyRow", MissingRow, i)
    MissingRow = MissingItems(HeaderRows, Working, "EduCertHeaderRow", "PremedHeaderRow", MissingRow, i)
    MissingRow = MissingItems(HeaderRows, Working, "PremedHeaderRow", "MedHeaderRow", MissingRow, i)
    MissingRow = MissingItems(HeaderRows, Working, "MedHeaderRow", "PostGradHeaderRow", MissingRow, i)
    MissingRow = MissingItems(HeaderRows, Working, "PostGradHeaderRow", "ExamHeaderRow", MissingRow, i)
    MissingRow = MissingItems(HeaderRows, Working, "ExamHeaderRow", "WorkHeaderRow", MissingRow, i)
    
    With Sheets("Summary")
'   Legal: between "Legal Documents" and "State Licenses"
        .Range("B" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "LegalHeaderRow", "StateHeaderRow", "B")
        .Range("C" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "LegalHeaderRow", "StateHeaderRow", "C")
        .Range("D" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "LegalHeaderRow", "StateHeaderRow", "D")
'   State Licenses: variable between rows "State Licenses" and "Certificates"
        .Range("E" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "StateHeaderRow", "CertHeaderRow", "B")
        .Range("F" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "StateHeaderRow", "CertHeaderRow", "C")
        .Range("G" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "StateHeaderRow", "CertHeaderRow", "D")
'   Certificates: "Certificates" and "Verifications of Certificates"
        .Range("H" & CStr(i)).Value = Round((TypeCounter(HeaderRows, Working, "CertHeaderRow", "VerifCertHeaderRow", "B") _
            + TypeCounter(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", "B")) / 2)
        .Range("I" & CStr(i)).Value = Round((TypeCounter(HeaderRows, Working, "CertHeaderRow", "VerifCertHeaderRow", "C") _
            + TypeCounter(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", "C")) / 2)
        .Range("J" & CStr(i)).Value = Round((TypeCounter(HeaderRows, Working, "CertHeaderRow", "VerifCertHeaderRow", "D") _
            + TypeCounter(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", "D")) / 2)
'   Verification of Certificates
'        .Range("K" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", "B")
'        .Range("L" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", "C")
'        .Range("M" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", "D")
'   Addition Information/Documents and Additional Items - Point Person
        .Range("K" & CStr(i)).Value = Round((TypeCounter(HeaderRows, Working, "AddHeaderRow", "EduCertHeaderRow", "B") _
            + TypeCounter(HeaderRows, Working, "PointAddPHeaderRow", "LastEmptyRow", "B")) / 2)
        .Range("L" & CStr(i)).Value = Round((TypeCounter(HeaderRows, Working, "AddHeaderRow", "EduCertHeaderRow", "C") _
            + TypeCounter(HeaderRows, Working, "PointAddPHeaderRow", "LastEmptyRow", "C")) / 2)
        .Range("M" & CStr(i)).Value = Round((TypeCounter(HeaderRows, Working, "AddHeaderRow", "EduCertHeaderRow", "D") _
            + TypeCounter(HeaderRows, Working, "PointAddPHeaderRow", "LastEmptyRow", "D")) / 2)
'   Education Certificates / PSVs
        .Range("N" & CStr(i)).Value = Round((TypeCounter(HeaderRows, Working, "EduCertHeaderRow", "PremedHeaderRow", "B") _
            + TypeCounter(HeaderRows, Working, "PremedHeaderRow", "MedHeaderRow", "B") _
            + TypeCounter(HeaderRows, Working, "MedHeaderRow", "PostGradHeaderRow", "B") _
            + TypeCounter(HeaderRows, Working, "PostGradHeaderRow", "ExamHeaderRow", "B")) / 4)
        .Range("O" & CStr(i)).Value = Round((TypeCounter(HeaderRows, Working, "EduCertHeaderRow", "PremedHeaderRow", "C") _
            + TypeCounter(HeaderRows, Working, "PremedHeaderRow", "MedHeaderRow", "C") _
            + TypeCounter(HeaderRows, Working, "MedHeaderRow", "PostGradHeaderRow", "C") _
            + TypeCounter(HeaderRows, Working, "PostGradHeaderRow", "ExamHeaderRow", "C")) / 4)
        .Range("P" & CStr(i)).Value = Round((TypeCounter(HeaderRows, Working, "EduCertHeaderRow", "PremedHeaderRow", "D") _
            + TypeCounter(HeaderRows, Working, "PremedHeaderRow", "MedHeaderRow", "D") _
            + TypeCounter(HeaderRows, Working, "MedHeaderRow", "PostGradHeaderRow", "D") _
            + TypeCounter(HeaderRows, Working, "PostGradHeaderRow", "ExamHeaderRow", "D")) / 4)
'   Work History
        .Range("Q" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "WorkHeaderRow", "HospHeaderRow", "B")
        .Range("R" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "WorkHeaderRow", "HospHeaderRow", "C")
        .Range("S" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "WorkHeaderRow", "HospHeaderRow", "D")
'   Affiliations
        .Range("T" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "HospHeaderRow", "InsHeaderRow", "B")
        .Range("U" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "HospHeaderRow", "InsHeaderRow", "C")
        .Range("V" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "HospHeaderRow", "InsHeaderRow", "D")
'   Insurance
        .Range("W" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "InsHeaderRow", "ReportHeaderRow", "B")
        .Range("X" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "InsHeaderRow", "ReportHeaderRow", "C")
        .Range("Y" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "InsHeaderRow", "ReportHeaderRow", "D")
'   Reports
        .Range("Z" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "ReportHeaderRow", "MilHeaderRow", "B")
        .Range("AA" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "ReportHeaderRow", "MilHeaderRow", "C")
        .Range("AB" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "ReportHeaderRow", "MilHeaderRow", "D")
'   Military
        .Range("AC" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "MilHeaderRow", "RefHeaderRow", "B")
        .Range("AD" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "MilHeaderRow", "RefHeaderRow", "C")
        .Range("AE" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "MilHeaderRow", "RefHeaderRow", "D")
'   References
        .Range("AF" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "RefHeaderRow", "PointAddPHeaderRow", "B")
        .Range("AG" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "RefHeaderRow", "PointAddPHeaderRow", "C")
        .Range("AH" & CStr(i)).Value = TypeCounter(HeaderRows, Working, "RefHeaderRow", "PointAddPHeaderRow", "D")
'   Totals
         .Range("AI" & CStr(i)).Formula = "=ROUND(AVERAGE(B" & CStr(i) _
         & ",E" & CStr(i) _
         & ",H" & CStr(i) _
         & ",K" & CStr(i) _
         & ",N" & CStr(i) _
         & ",Q" & CStr(i) _
         & ",T" & CStr(i) _
         & ",W" & CStr(i) _
         & ",Z" & CStr(i) _
         & ",AC" & CStr(i) _
         & ",AF" & CStr(i) _
         & "),0)"
         .Range("AJ" & CStr(i)).Formula = "=ROUND(AVERAGE(C" & CStr(i) _
         & ",F" & CStr(i) _
         & ",I" & CStr(i) _
         & ",L" & CStr(i) _
         & ",O" & CStr(i) _
         & ",R" & CStr(i) _
         & ",U" & CStr(i) _
         & ",X" & CStr(i) _
         & ",AA" & CStr(i) _
         & ",AD" & CStr(i) _
         & ",AG" & CStr(i) _
         & "),0)"
         .Range("AK" & CStr(i)).Formula = "=ROUND(AVERAGE(D" & CStr(i) _
         & ",G" & CStr(i) _
         & ",J" & CStr(i) _
         & ",M" & CStr(i) _
         & ",P" & CStr(i) _
         & ",S" & CStr(i) _
         & ",V" & CStr(i) _
         & ",Y" & CStr(i) _
         & ",AB" & CStr(i) _
         & ",AE" & CStr(i) _
         & ",AH" & CStr(i) _
         & "),0)"
         .Range("AL" & CStr(i)).Formula = "=(AI" & CStr(i) & "-AJ" & CStr(i) & ")"
    End With


Sheets("Missing Items").Columns(1).Resize(, counter).AutoFit
Sheets("Missing Items").Rows(1).Resize(MissingRow, 1).AutoFit
    
Next i
Sheets("Summary").Activate
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 0
End With
ActiveWindow.FreezePanes = True

End Sub
