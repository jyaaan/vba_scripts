Function ParseCSVLine( _
     QueryString As String) As Variant
    
    Dim ArrValueStore() As Variant
    Dim i As Integer, counter As Integer
    Dim TempPos As Integer, TempString As String
    i = 1
    counter = 0
    Do While InStr(i, QueryString, ",") <> 0
        TempPos = InStr(i, QueryString, ",")
        ReDim Preserve ArrValueStore(counter)
        ArrValueStore(counter) = Mid(QueryString, i, TempPos - i)
        i = TempPos + 1
        counter = counter + 1
    Loop
        ReDim Preserve ArrValueStore(counter)
        ArrValueStore(counter) = Mid(QueryString, i)
        ParseCSVLine = ArrValueStore
End Function

Sub ParseCSV()

    Dim i As Long, LastRow As Long
    Dim ThisWorksheet As Worksheet, TempData As String
    Dim ArrParsedData As Variant, j As Integer
    
    
    Set ThisWorksheet = ActiveSheet
    LastRow = GetLastUsedRowByIndex(ThisWorksheet.index)
    For i = 1 To LastRow
        TempData = ThisWorksheet.Cells(i, 1)
        ArrParsedData = ParseCSVLine(TempData)
        For j = 0 To (UBound(ArrParsedData))
            ThisWorksheet.Cells(i, j + 1) = ArrParsedData(j)
        Next j
    Next i
    'Rows("1:1").Select
    'Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    'ActiveCell.FormulaR1C1 = "Email"
    'Range("B1").Select
    'ActiveCell.FormulaR1C1 = "Week"
    'Range("C1").Select
   ' ActiveCell.FormulaR1C1 = "ContactEvents"
    'Columns("A:A").EntireColumn.AutoFit
    'Columns("C:C").EntireColumn.AutoFit
    'Columns("B:B").EntireColumn.AutoFit
End Sub

Sub ParseCSVSheet(TargetSheet As Worksheet)

    Dim i As Long, LastRow As Long
    Dim TempData As String
    Dim ArrParsedData As Variant, j As Integer
    

    LastRow = GetLastUsedRowByIndex(TargetSheet.index)
    For i = 1 To LastRow
        TempData = TargetSheet.Cells(i, 1)
        ArrParsedData = ParseCSVLine(TempData)
        For j = 0 To (UBound(ArrParsedData))
            TargetSheet.Cells(i, j + 1) = ArrParsedData(j)
        Next j
    Next i
    'Rows("1:1").Select
    'Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    'ActiveCell.FormulaR1C1 = "Email"
    'Range("B1").Select
    'ActiveCell.FormulaR1C1 = "Week"
    'Range("C1").Select
   ' ActiveCell.FormulaR1C1 = "ContactEvents"
    'Columns("A:A").EntireColumn.AutoFit
    'Columns("C:C").EntireColumn.AutoFit
    'Columns("B:B").EntireColumn.AutoFit
End Sub

'will parse out domains after '@' from a user email
'requires columns titled domain and email
Sub ParseDomains()

    Dim LastRowInSheet As Long, i As Long
    Dim DomainCol As Integer, EmailCol As Integer, Position As Integer
    Dim Domain As String
    
    With ActiveSheet
        LastRowInSheet = .Cells(Rows.count, "A").End(xlUp).Row
    End With
    
    EmailCol = FindColumnIndexByTitle("Email", ActiveSheet.index)
    DomainCol = FindColumnIndexByTitle("Domain", ActiveSheet.index)
    
    With ActiveSheet
    For i = 2 To LastRowInSheet
        Position = Len(.Cells(i, EmailCol)) - InStrRev(.Cells(i, EmailCol), "@")
        .Cells(i, DomainCol) = Right(.Cells(i, EmailCol), Position)
        
    Next i
    End With
End Sub

Function SerializeRowsToCSV( _
     Optional HasHeader As Boolean = True) As String

    Dim thisSheet As Worksheet, LastRow As Integer
    Dim Output As String, i As Integer, HeaderPresent As Integer
    Dim TargetCol As Integer
    
    Set thisSheet = ActiveSheet
    TargetCol = InputBox("Enter target column")
    If HasHeader Then
        HeaderPresent = 1
    Else
        HeaderPresent = 0
    End If
    
    With thisSheet
        LastRow = GetLastUsedRowByIndex(thisSheet.index)
        For i = (1 + HeaderPresent) To LastRow
            Output = Output + ",'" + .Cells(i, TargetCol) + "'"
        Next i
    End With
    
    SerializeRowsToCSV = Output
    thisSheet.Cells(1, 5) = Output
End Function

Sub CopyDataToClipboard(SendToClipboard As String)

    'Dim ClipBoard As MSForms.DataObject
    
    'Set ClipBoard = New MSForms.DataObject
    
    'ClipBoard.SetText SendToClipboard
    'ClipBoard.PutInClipboard


End Sub

Sub SerializeAndCopy()

    If MsgBox("Is a header present?", vbYesNo) = vbYes Then
        Call CopyDataToClipboard(SerializeRowsToCSV(True))
    Else
        Call CopyDataToClipboard(SerializeRowsToCSV(False))
    End If

End Sub