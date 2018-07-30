'last update 7/27/15
Function GetSheetByTitle( _
     SheetTitle As String) As Worksheet

    Dim i As Integer

    For i = 1 To ActiveWorkbook.Sheets.count
        If Sheets(i).Name = SheetTitle Then
            Set GetSheetByTitle = Sheets(i)
            Exit For
            
        Else
            'GetSheetByTitle = Nothing
        End If
    Next i
End Function
Function FindColumnIndexByTitle( _
     SearchQuery As String, _
     QuerySheetIndex As Integer) As Integer
'returns index of column with matching header
    Dim QueryWorksheet As Worksheet
    Dim HeaderRange As Range, SearchResult As Range
    Dim LastUsedColumn As Integer
    

    '~~> Change this to the relevant sheet
    Set QueryWorksheet = ActiveWorkbook.Sheets(QuerySheetIndex)
    LastUsedColumn = GetLastUsedColumnByIndex(QuerySheetIndex)
    
    With QueryWorksheet
        Set HeaderRange = Range(.Cells(1, 1), .Cells(1, LastUsedColumn))
        Set SearchResult = HeaderRange.Find(what:=SearchQuery, LookIn:=xlValues, lookat:=xlWhole, _
                    MatchCase:=False, searchformat:=False)

        '~~> If Found
        If Not SearchResult Is Nothing Then
            FindColumnIndexByTitle = SearchResult.Column
        '~~> If not found
        Else
            FindColumnIndexByTitle = -1
        End If
        
    End With
End Function

Function FindSheetIndexByTitle( _
     SearchQuery As String) As Integer
'returns sheet number with matching sheet name
    Dim i As Integer, MatchingSheetIndex As Integer
    
    
    MatchingSheetIndex = 0
    
    For i = 1 To ActiveWorkbook.Sheets.count
        
        If Sheets(i).Name = SearchQuery Then
            MatchingSheetIndex = i
        End If
        
    Next i
    
    If MatchingSheetIndex = 0 Then
        MsgBox "Sheet Not Found"
    Else
        FindSheetIndexByTitle = MatchingSheetIndex
    End If
    
End Function

Function GetLastUsedColumnByIndex( _
     QuerySheetIndex As Integer) As Integer
    Dim LastCol As Integer
    LastCol = Sheets(QuerySheetIndex).Range("A1").End(xlToRight).Column
    GetLastUsedColumnByIndex = LastCol
End Function
Function GetLastUsedColumn( _
     QuerySheet As Worksheet) As Integer
    GetLastUsedColumn = QuerySheet.UsedRange.Column.count
End Function

Function GetLastUsedRowByIndex( _
     QuerySheetIndex As Integer) As Long

    Dim LastUsedRow As Long
    
    With Sheets(QuerySheetIndex)
        LastUsedRow = .UsedRange.Rows.count
    End With
    
    GetLastUsedRowByIndex = LastUsedRow

End Function

Function GetLastUsedRow( _
     QuerySheet As Worksheet) As Variant
    GetLastUsedRowByIndex = QuerySheet.UsedRange.Rows.count
End Function

Function GetSearchRange( _
     SearchSheet As Worksheet, _
     SearchCol As Integer, _
     Optional HasHeader As Boolean = True) As Range

    Dim SearchRange As Range
    Dim LastRow As Long
    Dim FirstRow As Long
    
    If HasHeader Then
        FirstRow = 2
    Else
        FirstRow = 1
    End If
    
    With SearchSheet
        LastRow = GetLastUsedRowByIndex(.index)
        Set SearchRange = Range(.Cells(FirstRow, SearchCol), .Cells(LastRow, SearchCol))
    End With
    
    Set GetSearchRange = SearchRange

End Function

Function GetHeaderRange( _
     SearchSheet As Worksheet, _
     Optional SearchRow As Integer = 1) As Range
    Dim HeaderRange As Range
    Dim LastColumn As Integer
    LastColumn = GetLastUsedColumnByIndex(SearchSheet.index)
    With SearchSheet
        Set SearchRange = Range(.Cells(SearchRow, 1), .Cells(SearchRow, LastColumn))
    End With
    Set GetHeaderRange = SearchRange
End Function