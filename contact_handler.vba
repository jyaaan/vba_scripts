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

Sub SortNames()

    Dim i As Long
    Dim LastName As String
    Dim FirstName As String, TempName As String
    Dim CommaPos As Integer, StringLength As Integer
    Dim SpacePos As Integer
    Dim LastRow As Long
    Dim FirstNameCol As Integer, LastNameCol As Integer
    Dim NameCol As Integer
    Dim ws As Worksheet
    
    LastRow = GetLastUsedRowByIndex(ActiveSheet.index)
    FirstNameCol = InputBox("First Name Col")
    LastNameCol = InputBox("Last Name Col")
    NameCol = 2
    Set ws = ActiveSheet
    
    For i = 2 To LastRow
        
        TempName = ws.Cells(i, NameCol)
        CommaPos = InStr(TempName, ",")
        StringLength = Len(TempName)
        If CommaPos <> 0 Then
            If Mid(TempName, CommaPos + 1, 1) = " " Then
                FirstName = Right(TempName, StringLength - CommaPos - 1)
                LastName = Left(TempName, CommaPos - 1)
            Else
                FirstName = Right(TempName, StringLength - CommaPos)
                LastName = Left(TempName, CommaPos - 1)
            End If
            
        Else
            SpacePos = InStr(TempName, " ")
            If SpacePos > 0 Then
                LastName = Right(TempName, StringLength - SpacePos)
                FirstName = Left(TempName, SpacePos - 1)
                
            Else
                FirstName = "FIX"
                LastName = "ME!!!!!!"
            End If
            
        End If
        
        ws.Cells(i, FirstNameCol) = FirstName
        ws.Cells(i, LastNameCol) = LastName
        
        
    Next i

End Sub

Sub DeDupeEmails()

    Dim SearchSheet As Worksheet, SearchCol As Integer
    Dim QuerySheet As Worksheet, QueryCol As Integer
    Dim QueryReturnCol As Integer
    
    'change to make dynamic
    Set SearchSheet = Sheets(1)
    SearchCol = 3
    
    Set QuerySheet = ActiveSheet
    QueryCol = InputBox("Query Column")
    QueryReturnCol = InputBox("Return Column")
    
    'Call GetIsColumnDuplicate(QuerySheet, SearchSheet, QueryCol, SearchCol, QueryReturnCol)
    

End Sub
'
Sub ParseDomains()

    Dim LastRowInSheet As Long, i As Long
    Dim col As Integer, ecol As Integer, Position As Integer
    Dim Domain As String
    
    With ActiveSheet
        LastRowInSheet = .Cells(Rows.count, "A").End(xlUp).Row
    End With
    
    ecol = FindColumnIndexByTitle("Email", ActiveSheet.index)
    col = FindColumnIndexByTitle("Domain", ActiveSheet.index)
    
    With ActiveSheet
    For i = 2 To LastRowInSheet
        Position = Len(.Cells(i, ecol)) - InStrRev(.Cells(i, ecol), "@")
        .Cells(i, col) = Right(.Cells(i, ecol), Position)
        
    Next i
    End With
End Sub

Sub FlagExistingContactDomains()

    Dim IgnoreSourceWorksheet As Worksheet, ContactWorksheet As Worksheet
    Dim qdomcol As Integer, sdomcol As Integer
    Dim i As Long, j As Long
    Dim QuerySheetIndex As Integer, ssheet As Integer, IgnoreSheetIndex As Integer
    Dim qlrow As Long, slrow As Long, igrow As Long
    Dim QueryDomain As String
    Dim sCell As Range
        
    
    IgnoreSheetIndex = FindSheetIndexByTitle("Ignore")
    QuerySheetIndex = ActiveSheet.index
    ssheet = FindSheetIndexByTitle("Contacts")
    
    qlrow = GetLastUsedRowByIndex(QuerySheetIndex)
    igrow = GetLastUsedRowByIndex(IgnoreSheetIndex)
    slrow = GetLastUsedRowByIndex(ssheet)
    
    qdomcol = FindColumnIndexByTitle("Domain", QuerySheetIndex)
    sdomcol = FindColumnIndexByTitle("Domain", ssheet)
    
    Set IgnoreSourceWorksheet = ActiveWorkbook.Sheets(igsheet)
    Set ContactWorksheet = ActiveWorkbook.Sheets(ssheet)
    
    For i = 2 To qlrow
        QueryDomain = Sheets(QuerySheetIndex).Cells(i, qdomcol)
       
            With IgnoreSourceWorksheet
                Set sCell = .Range(.Cells(1, 1), .Cells(igrow, 1)).Find(what:=QueryDomain, _
                 LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False, searchformat:=False)
            End With
            If Not sCell Is Nothing Then
            
                'Call FlagCellAsTrue(QuerySheetIndex, i, FindColumnIndexByTitle("Exempt", QuerySheetIndex))
                
                
            '~~> If not found
    
            Else
            
                With sws
                    Set sCell = .Range(.Cells(2, sdomcol), .Cells(slrow, sdomcol)).Find(what:=QueryDomain, _
                     LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False, searchformat:=False)
                 
                    If Not sCell Is Nothing Then
                        'Call FlagCellAsTrue(QuerySheetIndex, i, FindColumnIndexByTitle("Exists", QuerySheetIndex))
                     End If
                 
                End With
            End If
            
    Next i

End Sub
Sub GetAllEmails()

    Dim ThisWorksheet As Worksheet, i As Integer
    Dim StartRow As Long, LastRow As Long
    Dim CurrentEmail As String, SearchRange As Range
    Dim SearchResult As Range
    Dim j As Long
    
    Set ThisWorksheet = ActiveSheet
    StartRow = 2
    For i = 1 To (ThisWorksheet.index - 1)
        LastRow = GetLastUsedRowByIndex(i)
        For j = 2 To LastRow
            CurrentEmail = Sheets(i).Cells(j, 1)
            Set SearchRange = GetSearchRange(ThisWorksheet, 1)
            Set SearchResult = SearchRange.Find(CurrentEmail)
            If SearchResult Is Nothing Then
                ThisWorksheet.Cells(StartRow, 1) = CurrentEmail
                StartRow = StartRow + 1
            End If
        Next j
    Next i

End Sub