Sub FindDateRange()
'
' FindDateRange Macro
'
' Keyboard Shortcut: Ctrl+d
'
    Selection.AutoFilter
    
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:= _
        Range("B1:B15774"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B2").Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:= _
        Range("B1:B15774"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B2").Select
    Selection.Copy
    Range("E1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.AutoFilter
End Sub


Sub AggregateUsage()
'
' AggregateUsage Macro
'
' Keyboard Shortcut: Ctrl+e
'
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
End Sub

Sub Horizontalize()

    Dim NumberOfRows As Integer
    Dim StartRow As Integer, StartCol As Integer
    Dim i As Integer, j As Integer, counter As Integer
    
    
    
    
    NumberOfRows = InputBox("Number of Rows")
    StartRow = ActiveCell.Row
    StartCol = ActiveCell.Column
    i = StartRow
    j = StartCol
    counter = 0
    
    With ActiveSheet
    For i = StartRow To (StartRow + NumberOfRows)
        .Cells(StartRow, StartCol + counter) = .Cells(StartRow + counter, StartCol)
        counter = counter + 1
    Next i
    End With
    
    
    
    

End Sub

Sub Verticalize()

    Dim NumberOfCols As Integer
    Dim StartCols As Integer, StartCol As Integer
    Dim i As Integer, j As Integer, counter As Integer
    
    
    NumberOfCols = InputBox("Number of Columns")
    StartRow = ActiveCell.Row
    StartCol = ActiveCell.Column
    i = StartRow
    j = StartCol
    counter = 0
    
    With ActiveSheet
    For i = StartCol To (StartCol + NumberOfCols)
        .Cells(StartRow + counter, StartCol) = .Cells(StartRow, StartCol + counter)
        If counter <> 0 Then
            .Cells(StartRow, StartCol + counter) = ""
        End If
        counter = counter + 1
    Next i
    End With

End Sub