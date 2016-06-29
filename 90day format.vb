Sub UnMergeSameCell()
Dim Rng As Range
Dim Length As String
Dim Filter As Range

With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
		.DisplayStatusBar = False
End With
Length = CStr(ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row)
For i = 2 To Length
    i = CStr(i)
    Range("A" & i & ":E" & i).Borders.LineStyle = xlLineStyleNone
    If Range("A" & i).MergeCells Then
        With Range("A" & i).MergeArea
            .UnMerge
            .Formula = Range("A" & i)
        End With
    ElseIf Range("B" & i).MergeCells Then
        With Range("B" & i).MergeArea
            .UnMerge
        End With
    ElseIf Range("C" & i).MergeCells Then
        With Range("C" & i).MergeArea
            .UnMerge
        End With
    ElseIf Range("D" & i).MergeCells Then
        With Range("D" & i).MergeArea
            .UnMerge
        End With
    ElseIf Range("E" & i).MergeCells Then
        With Range("E" & i).MergeArea
            .UnMerge
        End With
    End If
Next i

For i = Length To 2 Step -1
    j = CStr(i)
    If ((Range("B" & j).Offset(1, -1).Value = Range("B" & j).Offset(0, -1).Value) Or IsEmpty(Range("A" & j).Value)) _
    And IsEmpty(Range("B" & j).Value) And IsEmpty(Range("C" & j).Value) And _
    IsEmpty(Range("D" & j).Value) And IsEmpty(Range("E" & j).Value) Then
        Rows(i).EntireRow.Delete
    End If
Next i

Columns("A:E").AutoFilter
With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
		.DisplayStatusBar = True
End With
End Sub
