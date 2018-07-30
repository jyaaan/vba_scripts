Sub CreateNewSheetOfTitle( _
     SheetName As String)

    Sheets.Add after:=Sheets(Sheets.count)
    Sheets(Sheets.count).Name = SheetName
    
End Sub

Sub CreateNewSheet( _
     Optional SheetName As String)

    Sheets.Add after:=Sheets(Sheets.count)
    If SheetName <> "" Then
        Sheets(Sheets.count).Name = SheetName
    End If

End Sub
Sub InsertColumnHeaderAfter( _
     ColumnNumber As Integer, _
     SheetNumber As Integer, _
     HeaderInsert As Variant)

    Sheets(SheetNumber).Cells(1, ColumnNumber).EntireColumn.Offset(0, 1).Insert
    Sheets(SheetNumber).Cells(1, ColumnNumber + 1).Value = HeaderInsert
    
End Sub
Sub InsertRowAfter( _
     RowNumber As Long, _
     SheetNumber As Integer)

    Sheets(SheetNumber).Rows(RowNumber).Offset (1)
    
End Sub
Sub InsertRowAt( _
     RowNumber As Long, _
     InsertSheet As Worksheet, _
     InsertString As String)

    InsertSheet.Cells(RowNumber, 1) = InsertString
    
End Sub

Sub InsertRowAtEnd( _
     InsertSheet As Worksheet, _
     InsertCol As Integer, _
     InsertString As String)
    Dim LastRow As Long
    
    With InsertSheet
        LastRow = GetLastUsedRowByIndex(.index)
        .Cells(LastRow + 1, InsertCol) = InsertString
    End With
    
End Sub

Sub InsertColumnAtEnd( _
     InsertSheet As Worksheet, _
     ColumnName As String)
    Dim LastCol As Integer
    
    LastCol = GetLastUsedColumnByIndex(InsertSheet.index)
    InsertSheet.Cells(1, LastCol + 1) = ColumnName
    
End Sub
Sub InsertCellData( _
     SheetNumber As Integer, _
     DestinationColumn As Integer, _
     DestinationRow As Long, _
     InsertData As Variant)

    Sheets(SheetNumber).Cells(DestinationRow, DestinationColumn).Value = InsertData

End Sub

Sub FilterSheet( _
     FilterColumn As Integer, _
     QueryString As String, _
     QuerySheet As Worksheet, _
     Optional IsHeaderPresent As Boolean = True)
    'Will filter a sheet for a match
    'does not check to see if query exists within sheet.
    
    
    
End Sub

Sub RemoveFilterSheet( _
     QuerySheet As Worksheet)
'removes applied filters from given sheet.
End Sub