Sub PrintEmailComparison( _
      isPrintMissing As Boolean, _
      isShowReport As Boolean, _
      QuerySheet As Worksheet, _
      SearchSheet As Worksheet)
     'not fully dynamic. needs querycol

    Dim QueryEmail As String
    Dim i As Long, LastQueryRow As Long
    Dim SearchRange As Range, SearchResult As Range
    Dim ResultSheet As Worksheet
    Dim counter As Long
    
    
    Set SearchRange = GetSearchRange(SearchSheet, 1, False)
    'look at querysheet. get number of elements
    LastQueryRow = GetLastUsedRowByIndex(QuerySheet.index)
    
    'create result sheet
    If isPrintMissing Then
        Call CreateNewSheet("Missing Emails")
    Else
        Call CreateNewSheet("Matching Emails")
    End If
    
    Set ResultSheet = Sheets(Sheets.count)
    counter = 1
    
    'for each element in query sheet, find if there exists
    'Create result sheet. name dependent upon whether we are printing missing or existing
    'if it exists
    'if statement to print or not to print in new sheet.
    For i = 1 To LastQueryRow
        QueryEmail = QuerySheet.Cells(i, 1)
        Set SearchResult = SearchRange.Find(QueryEmail)
        If (Not SearchResult Is Nothing) And (Not isPrintMissing) Then
            Call InsertRowAt(counter, ResultSheet, QueryEmail)
            counter = counter + 1
        ElseIf (SearchResult Is Nothing) And isPrintMissing Then
            Call InsertRowAt(counter, ResultSheet, QueryEmail)
            counter = counter + 1
        End If
    Next i
    
    

End Sub

Sub CompareEmailsBetweenSheets()

    Dim QuerySheetName As String, SearchSheetName As String
    Dim QuerySheet As Worksheet, SearchSheet As Worksheet
    Dim isPrintMissing As Boolean
    
    QuerySheetName = InputBox("Query Sheet Name plox")
    SearchSheetName = InputBox("Search Sheet Name plox")
    If MsgBox("Print Missing?", vbYesNo) = vbYes Then
        isPrintMissing = True
    Else
        isPrintMissing = False
    End If
    
    
    Set QuerySheet = GetSheetByTitle(QuerySheetName)
    Set SearchSheet = GetSheetByTitle(SearchSheetName)
    
    Call PrintEmailComparison(isPrintMissing, True, QuerySheet, SearchSheet)

End Sub
'query: the main data search: the list of blacklisted items
Sub MarkMatches()

    Dim QuerySheetName As String, SearchSheetName As String
    Dim QuerySheet As Worksheet, SearchSheet As Worksheet
    Dim isMarkMissing As Boolean
    
    QuerySheetName = InputBox("Query Sheet Name plox(main data)")
    SearchSheetName = InputBox("Search Sheet Name plox(blacklist)")
    If MsgBox("Mark", vbYesNo) = vbYes Then
        isMarkMissing = True
    Else
        isMarkMissing = False
    End If
    
    
    Set QuerySheet = GetSheetByTitle(QuerySheetName)
    Set SearchSheet = GetSheetByTitle(SearchSheetName)
    
    Call MarkMatching(isMarkMissing, True, QuerySheet, SearchSheet)

End Sub

Sub MarkMatching( _
     isMarkMissing As Boolean, _
     isShowReport As Boolean, _
     QuerySheet As Worksheet, _
     SearchSheet As Worksheet)
     'not fully dynamic. needs querycol

    Dim QueryEmail As String, MarkCol As Integer, QueryDataCol As Integer
    Dim i As Long, LastQueryRow As Long
    Dim SearchRange As Range, SearchResult As Range
    Dim ResultSheet As Worksheet
    Dim counter As Long
    Dim QueryCol As Integer, SearchCol As Integer
    QueryCol = InputBox("Query main source column")
    MarkCol = InputBox("Mark destination Col")
    
    SearchCol = InputBox("Search blacklist column")
    'QueryDataCol = InputBox("Query data destination column")
    Set SearchRange = GetSearchRange(SearchSheet, SearchCol, True)
    'look at querysheet. get number of elements
    LastQueryRow = GetLastUsedRowByIndex(QuerySheet.index)
    
    
    Set ResultSheet = QuerySheet
    counter = 1
    
    'for each element in query sheet, find if there exists
    'Create result sheet. name dependent upon whether we are printing missing or existing
    'if it exists
    'if statement to print or not to print in new sheet.
    For i = 2 To LastQueryRow
        QueryEmail = QuerySheet.Cells(i, QueryCol)
        Set SearchResult = SearchRange.Find(QueryEmail)
        If (Not SearchResult Is Nothing) Then
            QuerySheet.Cells(i, MarkCol) = "x"
            'QuerySheet.Cells(i, MarkCol) = SearchSheet.Cells(SearchResult.Row, QueryDataCol)
        ElseIf (SearchResult Is Nothing) Then
            'what happens when there is no match
        End If
    Next i
    
    

End Sub
'This will look for ID matches between sheets. on match, the contents of source column will be sent to target column
'remember, matches go from source to target
'can only do 1 to 1
Sub TransferMatchingData( _
     SourceSheet As Worksheet, _
     TargetSheet As Worksheet, _
     SourceMatchCol As Integer, _
     TargetMatchCol As Integer, _
     SourceDataCol As Integer, _
     TargetDataCol As Integer, _
     Optional HasHeader As Boolean = True)
    
    Dim i As Integer, SourceLastRow As Integer, thisID As String
    Dim SearchRange As Range, SearchResult As Range, StartRow As Integer
    
    Set SearchRange = GetSearchRange(TargetSheet, TargetMatchCol)
    SourceLastRow = GetLastUsedRow(SourceSheet)
    
    If HasHeader Then
        StartRow = 2
    Else
        StartRow = 1
    End If
    
    For i = StartRow To LastRow
        thisID = SourceSheet.Cells(i, SourceMatchCol)
        Set SearchResult = SearchRange.Find(thisID)
        If Not SearchResult Is Nothing Then
            TargetSheet.Cells(SearchResult.Row, TargetDataCol) = _
             SourceSheet.Cells(i, SourceDataCol)
        End If
    Next i
    

End Sub