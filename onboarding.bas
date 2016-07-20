Attribute VB_Name = "Module1"
Public Function TypeCounter(HeaderRows, Working, TopHeaderRow, BottomHeaderRow, ReqRcvUp)
Dim total, Completed, Percent As Long

Completed = 0
total = 0

For leg = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
    If Working.Range(ReqRcvUp & CStr(leg)).Interior.ColorIndex <> 1 Then
        total = total + 1
    End If
    If ((IsEmpty(Working.Range(ReqRcvUp & CStr(leg)).Value) = False) And (Working.Range(ReqRcvUp & CStr(leg)).Interior.ColorIndex <> 1)) _
    Or Working.Range(ReqRcvUp & CStr(leg)).Interior.ColorIndex = 15 Then
        Completed = Completed + 1
    End If
Next leg
If total = 0 Then
    TypeCounter = 100
Else
    TypeCounter = Round((Completed / total) * 100)
End If

End Function

Public Function MissingItems(HeaderRows, Working, TopHeaderRow, BottomHeaderRow, MissingRow, i)
If TopHeaderRow = "StateHeaderRow" Then
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
ElseIf TopHeaderRow = "WorkHeaderRow" Then
    For mrow = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
        If IsEmpty(Working.Range("C" & mrow).Value) And Working.Range("C" & mrow).Interior.ColorIndex <> 1 _
        And Working.Range("C" & mrow).Interior.ColorIndex <> 15 Then
            If Working.Range("A" & mrow).Value Like "*Work*" Then
                Sheets("Missing Items").Cells(MissingRow, i - 1).Value = Working.Range("A" & mrow).Value
            Else
                Sheets("Missing Items").Cells(MissingRow, i - 1).Value = "Work: " & Working.Range("A" & mrow).Value
            End If
            MissingRow = MissingRow + 1
            
        End If
    Next mrow
ElseIf TopHeaderRow = "HospHeaderRow" Then
    For mrow = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
        If IsEmpty(Working.Range("C" & mrow).Value) And Working.Range("C" & mrow).Interior.ColorIndex <> 1 _
        And Working.Range("C" & mrow).Interior.ColorIndex <> 15 Then
            If Working.Range("A" & mrow).Value Like "Hospital Verif*" Then
                Sheets("Missing Items").Cells(MissingRow, i - 1).Value = Working.Range("A" & mrow).Value
            Else
                Sheets("Missing Items").Cells(MissingRow, i - 1).Value = "Hospital: " & Working.Range("A" & mrow).Value
            End If
            MissingRow = MissingRow + 1
            
        End If
    Next mrow
ElseIf TopHeaderRow = "InsHeaderRow" Then
    For mrow = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
        If IsEmpty(Working.Range("C" & mrow).Value) And Working.Range("C" & mrow).Interior.ColorIndex <> 1 _
        And Working.Range("C" & mrow).Interior.ColorIndex <> 15 Then
            If Working.Range("A" & mrow).Value Like "Insurance Verif*" Then
                Sheets("Missing Items").Cells(MissingRow, i - 1).Value = Working.Range("A" & mrow).Value
            Else
                Sheets("Missing Items").Cells(MissingRow, i - 1).Value = "Insurance: " & Working.Range("A" & mrow).Value
            End If
            MissingRow = MissingRow + 1
            
        End If
    Next mrow
ElseIf TopHeaderRow = "RefHeaderRow" Then
    For mrow = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
        If IsEmpty(Working.Range("C" & mrow).Value) And Working.Range("C" & mrow).Interior.ColorIndex <> 1 _
        And Working.Range("C" & mrow).Interior.ColorIndex <> 15 Then
            If Working.Range("A" & mrow).Value Like "Reference*" Then
                Sheets("Missing Items").Cells(MissingRow, i - 1).Value = Working.Range("A" & mrow).Value
            Else
                Sheets("Missing Items").Cells(MissingRow, i - 1).Value = "Reference: " & Working.Range("A" & mrow).Value
            End If
            MissingRow = MissingRow + 1
            
        End If
    Next mrow
Else
    For mrow = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
        If IsEmpty(Working.Range("C" & mrow).Value) And Working.Range("C" & mrow).Interior.ColorIndex <> 1 _
        And Working.Range("C" & mrow).Interior.ColorIndex <> 15 Then
            Sheets("Missing Items").Cells(MissingRow, i - 1).Value = Working.Range("A" & mrow).Value
            MissingRow = MissingRow + 1
        End If
    Next mrow
End If
MissingItems = MissingRow
End Function

Public Function Check_delete(sheet)
counter = Worksheets.Count
Application.DisplayAlerts = False
For shnum = counter To 1 Step -1
    If Sheets(shnum).Name = sheet Then
        Sheets(shnum).Delete
    End If
Next shnum

Application.DisplayAlerts = True


End Function

Public Function DateDiffCalc(HeaderRows, Working, TopHeaderRow, BottomHeaderRow, i)
' Working = working sheet
' HeaderRows = dict of row #s
' Top/Bottom = dict keys
' column of final info
Dim TotalDiff As Long
Dim numdiff As Integer
Dim DateSplit() As String
Dim FirstDate As String

Dim ReqDate_A() As String
Dim RcvDate_A() As String
error_row = 24

TotalDiff = 0
numdiff = 0
For dadiff = HeaderRows(TopHeaderRow) + 1 To HeaderRows(BottomHeaderRow) - 1
    On Error GoTo ErrorHandler
    ReqDate = CStr(Working.Range("B" + CStr(dadiff)).Value)
    rcvdate = CStr(Working.Range("C" + CStr(dadiff)).Value)
    
    If ReqDate Like "*/*/*" And rcvdate Like "*/*/*" Then
        ReqDate = Replace(ReqDate, Chr(10), " ")
        rcvdate = Replace(rcvdate, Chr(10), " ")
        ReqDate = Replace(ReqDate, ",", " ")
        rcvdate = Replace(rcvdate, ",", " ")
        
        
        RcvDate_A = Split(rcvdate, " ")
        ReqDate = Split(ReqDate, " ")(0)
        rcvdate = RcvDate_A(UBound(RcvDate_A))
        
        If Right(ReqDate, 1) = "/" Then
            ReqDate = Left(ReqDate, Len(ReqDate) - 1)
        End If
        If Right(rcvdate, 1) = "/" Then
            rcvdate = Left(rcvdate, Len(rcvdate) - 1)
        End If
        If Not (ReqDate Like "*/*/*") Then
            ReqDate = ReqDate + "/" + CStr(Year(Date))
        End If
        If Not (rcvdate Like "*/*/*") Then
            rcvdate = rcvdate + "/" + CStr(Year(Date))
        End If
        ReqDate = CDate(ReqDate)
        rcvdate = CDate(rcvdate)
        If Year(ReqDate) <= Year(rcvdate) And Year(rcvdate) <= Year(Now) + 1 Then
            TotalDiff = TotalDiff + DateDiff("d", ReqDate, rcvdate)
            numdiff = numdiff + 1
        End If
    ElseIf rcvdate <> "" Then
        TotalDiff = TotalDiff + 1
        numdiff = numdiff + 1
    End If
    
Next dadiff

    

 If numdiff = 0 Then
    DateDiffCalc = "N/A"
 Else
     DateDiffCalc = Round((TotalDiff / numdiff))
 End If

Exit Function
ErrorHandler:

Do While True
    If Sheets("Date Difference").Range("B" & CStr(error_row)).Value <> "" Then
      error_row = error_row + 1
    Else
       Sheets("Date Difference").Range("B" & CStr(error_row)).Value = Working.Name & ": Malformed Date -- Line:  " & _
       CStr(dadiff) & "  (Request Date:  " & CStr(ReqDate) & "  Received Date:  " & CStr(rcvdate) & ")"
       Exit Do
    End If
'     change this to a list at the end instead of pop-ups (maybe add the dates outside sum range above too) starts at B24
'     MsgBox Working.Name & " - Malformed Date" & vbNewLine & "Line:  " & CStr(dadiff) & vbNewLine & "Request Date:  " & CStr(ReqDate) & vbNewLine & "Received Date:  " & CStr(rcvdate)
     Err.Clear
Loop
End Function

Sub Information()
Attribute Information.VB_ProcData.VB_Invoke_Func = "R\n14"
'   Set variables
Dim counter, template_row, LastWorkRow As Long
Dim Working As Worksheet
Dim HeaderRows As Collection
Dim RowLabels As Variant
Dim HeaderNames As Variant

' speed tweaks
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False

'   check for summary worksheet, if exists, deletes it
Check_delete "Missing Items"
Check_delete "Summary"
Check_delete "Date Difference"


'   creating worksheet names
counter = Worksheets.Count
Sheets.Add After:=Sheets(counter)
Sheets(counter + 1).Name = "Summary"
Sheets.Add After:=Sheets("Summary")
Sheets(counter + 2).Name = "Missing Items"
Sheets.Add After:=Sheets("Missing Items")
Sheets(counter + 3).Name = "Date Difference"

'   Create and set Summary headers
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

' Create Date Difference Row Labels
RowLabels = Array( _
"Legal Documents", "State Licenses", "Certificates", _
"Verification of Certificates", "Additional Info / Docs", _
"Education Certificates", "Premed", "Medical School", _
"Post Graduate Training", "Exam Records", "Work History", _
"Hospital Affiliations", "Insurance", "Reports", "Military", _
"References", "Additional Items")
RowLabels = WorksheetFunction.Transpose(RowLabels)
Sheets("Date Difference").Range("A1").Interior.ColorIndex = 1
Sheets("Date Difference").Range("A24").Value = "Errors"
Sheets("Date Difference").Range("A24").Interior.ColorIndex = 3
With Sheets("Date Difference").Range("A2:A18")
    .Value = RowLabels
    .Interior.ColorIndex = 23
    .Font.Color = RGB(255, 255, 255)
End With
With Sheets("Date Difference").Range("A12:A14")
    .Interior.ColorIndex = 15
    .Font.Color = RGB(0, 0, 0)
End With
With Sheets("Date Difference").Range("A17")
    .Interior.ColorIndex = 15
    .Font.Color = RGB(0, 0, 0)
End With

Sheets("Date Difference").Columns("A").AutoFit
Sheets("Date Difference").Columns("B:AZ").ColumnWidth = 15

'   Add physicians name to summary page
For i = 1 To counter
    If Sheets(i).Name <> "Template" Then
        Sheets("Summary").Range("A" & CStr(i + 1)).Value = Sheets(i).Name
        Sheets("Summary").Range("A" & CStr(i + 1)).Interior.ColorIndex = Sheets(i).Tab.ColorIndex
        Sheets("Missing Items").Cells(1, i).Value = Sheets(i).Name
        Sheets("Missing Items").Cells(1, i).Interior.ColorIndex = Sheets(i).Tab.ColorIndex
        Sheets("Date Difference").Cells(1, i + 1).Value = Sheets(i).Name
        Sheets("Date Difference").Cells(1, i + 1).Interior.ColorIndex = Sheets(i).Tab.ColorIndex
        If Sheets(i).Tab.ColorIndex = 1 Then
            Sheets("Missing Items").Cells(1, i).Font.Color = RGB(255, 255, 255)
            Sheets("Summary").Range("A" & CStr(i + 1)).Font.Color = RGB(255, 255, 255)
            Sheets("Date Difference").Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
        End If
        
    Else
        template_row = i + 1
    End If
Next i
Sheets("Missing Items").Columns(template_row - 1).EntireColumn.Delete Shift:=xlToLeft
Sheets("Date Difference").Columns(template_row).EntireColumn.Delete Shift:=xlToLeft
'   Resize columns
With Sheets("Summary")
    .Rows(template_row).Delete
    .Columns("A").AutoFit
    .Rows(1).RowHeight = 45
    .Columns("B:AL").ColumnWidth = 12
End With
With Sheets("Summary").Range("A1:AL1")
    .WrapText = True
    .VerticalAlignment = xlTop
    .HorizontalAlignment = xlCenter
End With
'On Error Resume Next

'   Iterate over worksheets (should be '2 to counter')
For i = 2 To counter
    
'   Clear iteratives to be default for new worksheet
    Set HeaderRows = New Collection
    MissingRow = 2
'   Get worksheet from summary page, set to "Working"
    Set Working = Sheets(Sheets("Summary").Range("A" & CStr(i)).Value)
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
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Medical School*" _
        And Working.Range("A" & CStr(j)).Interior.ColorIndex = 23 Then
            HeaderRows.Add j, "MedHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Post Graduate Training*" Then
            HeaderRows.Add j, "PostGradHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "*Exam Records*" Then
            HeaderRows.Add j, "ExamHeaderRow"
        ElseIf Working.Range("A" & CStr(j)).Value Like "Work History*" _
        And Working.Range("A" & CStr(j)).Interior.ColorIndex <> 1 Then
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
    
'   The missing item spreadsheet is filled out using the MissingItems function while
'   the Date Difference spreadsheet uses the DateDiffCalc function (see above)
    
    MissingRow = MissingItems(HeaderRows, Working, "LegalHeaderRow", "StateHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(2, i).Value = DateDiffCalc(HeaderRows, Working, "LegalHeaderRow", "StateHeaderRow", i)
    Sheets("Date Difference").Cells(2, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "StateHeaderRow", "CertHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(3, i).Value = DateDiffCalc(HeaderRows, Working, "StateHeaderRow", "CertHeaderRow", i)
    Sheets("Date Difference").Cells(3, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "CertHeaderRow", "VerifCertHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(4, i).Value = DateDiffCalc(HeaderRows, Working, "CertHeaderRow", "VerifCertHeaderRow", i)
    Sheets("Date Difference").Cells(4, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(5, i).Value = DateDiffCalc(HeaderRows, Working, "VerifCertHeaderRow", "AddHeaderRow", i)
    Sheets("Date Difference").Cells(5, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "AddHeaderRow", "EduCertHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(6, i).Value = DateDiffCalc(HeaderRows, Working, "AddHeaderRow", "EduCertHeaderRow", i)
    Sheets("Date Difference").Cells(6, i).HorizontalAlignment = xlHAlignRight
      
    MissingRow = MissingItems(HeaderRows, Working, "EduCertHeaderRow", "PremedHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(7, i).Value = DateDiffCalc(HeaderRows, Working, "EduCertHeaderRow", "PremedHeaderRow", i)
    Sheets("Date Difference").Cells(7, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "PremedHeaderRow", "MedHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(8, i).Value = DateDiffCalc(HeaderRows, Working, "PremedHeaderRow", "MedHeaderRow", i)
    Sheets("Date Difference").Cells(8, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "MedHeaderRow", "PostGradHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(9, i).Value = DateDiffCalc(HeaderRows, Working, "MedHeaderRow", "PostGradHeaderRow", i)
    Sheets("Date Difference").Cells(9, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "PostGradHeaderRow", "ExamHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(10, i).Value = DateDiffCalc(HeaderRows, Working, "PostGradHeaderRow", "ExamHeaderRow", i)
    Sheets("Date Difference").Cells(10, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "ExamHeaderRow", "WorkHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(11, i).Value = DateDiffCalc(HeaderRows, Working, "ExamHeaderRow", "WorkHeaderRow", i)
    Sheets("Date Difference").Cells(11, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "WorkHeaderRow", "HospHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(12, i).Value = DateDiffCalc(HeaderRows, Working, "WorkHeaderRow", "HospHeaderRow", i)
    Sheets("Date Difference").Cells(12, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "HospHeaderRow", "InsHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(13, i).Value = DateDiffCalc(HeaderRows, Working, "HospHeaderRow", "InsHeaderRow", i)
    Sheets("Date Difference").Cells(13, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "InsHeaderRow", "ReportHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(14, i).Value = DateDiffCalc(HeaderRows, Working, "InsHeaderRow", "ReportHeaderRow", i)
    Sheets("Date Difference").Cells(14, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "ReportHeaderRow", "MilHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(15, i).Value = DateDiffCalc(HeaderRows, Working, "ReportHeaderRow", "MilHeaderRow", i)
    Sheets("Date Difference").Cells(15, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "MilHeaderRow", "RefHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(16, i).Value = DateDiffCalc(HeaderRows, Working, "MilHeaderRow", "RefHeaderRow", i)
    Sheets("Date Difference").Cells(16, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "RefHeaderRow", "PointAddPHeaderRow", MissingRow, i)
    Sheets("Date Difference").Cells(17, i).Value = DateDiffCalc(HeaderRows, Working, "RefHeaderRow", "PointAddPHeaderRow", i)
    Sheets("Date Difference").Cells(17, i).HorizontalAlignment = xlHAlignRight
    
    MissingRow = MissingItems(HeaderRows, Working, "PointAddPHeaderRow", "LastEmptyRow", MissingRow, i)
    Sheets("Date Difference").Cells(18, i).Value = DateDiffCalc(HeaderRows, Working, "PointAddPHeaderRow", "LastEmptyRow", i)
    Sheets("Date Difference").Cells(18, i).HorizontalAlignment = xlHAlignRight
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

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub
