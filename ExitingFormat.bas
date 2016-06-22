Attribute VB_Name = "Module1"
Sub Fromatting()

' Fromatting Macro
' Format Rad Priv Report Data for processing


' Delete extra columns
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("B:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:H").Select
    Selection.Delete Shift:=xlToLeft

' Filter for active applications
    ActiveSheet.ListObjects("Worklist").Range.AutoFilter Field:=3, Criteria1:= _
        Array("REAP-RC", "REAP-INP", "REAP-ELEC", "REAP-ORIG", "REAP-QA", _
        "RC", "INP", "ELEC-SIG", "ORIG-SIG", "INP-QA"), Operator:= _
        xlFilterValues
    ActiveSheet.ListObjects("Worklist").Range.AutoFilter Field:=4, Criteria1:="<>See Note (Rad Leaving)", _
    Operator:=xlAnd, Criteria2:="<>RESIGN"
        Columns("A:E").Select
 ' Move to data page
    Selection.Copy
    Sheets("Data").Select
    Range("D4").Select
    ActiveSheet.Paste
    

' Look for matching names and copy the dates next to them
   
    lastRowRads = Worksheets("Rads").Range("D3").End(xlDown).Row
    lastRowData = Worksheets("Data").Range("D4").End(xlDown).Row
    ReferenceRange = "D5:D" & lastRowData
    TestRange = "D3:D" & lastRowRads
    If lastRowData < 200 Then
        For Each rcell In Worksheets("Data").Range(ReferenceRange)
            For Each tcell In Worksheets("Rads").Range(TestRange)
                If rcell.Value = tcell.Value Then
                   rcell.Offset(0, -1).Value = tcell.Offset(0, -2).Value
                End If
            Next tcell
        Next rcell
    
' Fill privs last column
        Range("I5").Activate
        If lastRowData <> 5 Then
            Selection.AutoFill Destination:=Range("I5:I" & lastRowData), Type:=xlFillDefault
        End If
        Range("C4:I" & lastRowData).Select
        Selection.AutoFilter Field:=7, Criteria1:="Yes"
        
            ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add Key:=Range( _
            "E4:E" & lastRowData), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add Key:=Range( _
            "D4:D" & lastRowData), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End Sub


