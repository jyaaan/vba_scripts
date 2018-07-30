Function GetSearchSheetIndexes( _
     IndexLimit As Integer, _
     QueryWeek As String) As Integer()
    Dim SheetIndexArr() As Integer
    Dim i As Integer, StartDate As String, EndDate As String
    Dim ArrayIndex As Integer
    
    ArrayIndex = 0
    
    For i = 1 To IndexLimit
        StartDate = Sheets(i).Cells(1, 4)
        EndDate = Sheets(i).Cells(1, 5)
        If Left(QueryWeek, 4) >= Left(StartDate, 4) Then
            
        End If
        
    Next i
    
    GetSearchSheets = SheetIndexArr
End Function
'generates new sheets with matching YYYY-WW. but make sure you sort by week first or it'll try to create duplicate sheets
Sub GenerateWeeklySheets()

    Dim LastRow As Long, SourceSheet As Worksheet
    Dim i As Long, CurrentWeek As String
    Dim TargetSheet As Worksheet, j As Integer
    Dim TargetRowCounter As Long
    
    Set SourceSheet = ActiveSheet
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    CurrentWeek = SourceSheet.Cells(2, 2)
    CreateNewSheet (CurrentWeek)
    Set TargetSheet = Sheets(Sheets.count)
    TargetSheet.Cells(1, 1) = "Email"
    TargetSheet.Cells(1, 2) = "Week"
    TargetSheet.Cells(1, 3) = "ContactEvents"
    TargetRowCounter = 2
    For i = 2 To LastRow
        If SourceSheet.Cells(i, 2) <> CurrentWeek Then
            CurrentWeek = SourceSheet.Cells(i, 2)
            CreateNewSheet (CurrentWeek)
            Set TargetSheet = Sheets(Sheets.count)
            TargetSheet.Cells(1, 1) = "Email"
            TargetSheet.Cells(1, 2) = "Week"
            TargetSheet.Cells(1, 3) = "ContactEvents"
            TargetRowCounter = 2
        End If
        For j = 1 To 3
            TargetSheet.Cells(TargetRowCounter, j) = SourceSheet.Cells(i, j)
        Next j
        TargetRowCounter = TargetRowCounter + 1
    Next i

End Sub

'will take empty table with emails and YYYY-WW axis to generate usage history
'FIX
Sub PopulateExistingUsageTable()

    Dim EmailCol As Integer, LastCol As Integer, i As Integer, j As Long
    Dim SearchRange As Range, SearchResult As Range, k As Long
    Dim thisSheet As Worksheet, CurrentTimespan As String, SourceSheet As Worksheet
    Dim SourceLastRow As Long, SourceEmailCol As Integer, SourceUsageCol As Integer
    Dim LastRow As Long
    
    'connects to current sheet
    Set thisSheet = ActiveSheet
    LastCol = GetLastUsedColumnByIndex(thisSheet.index)
    EmailCol = FindColumnIndexByTitle("Email", ActiveSheet.index)
    LastRow = GetLastUsedRowByIndex(thisSheet.index)
    
    'goes through each YYYY-WW column and finds a corresponding sheet and finds col index
    For i = (EmailCol + 1) To LastCol
        CurrentTimespan = thisSheet.Cells(1, i)
        'Set SourceSheet = ReturnSheetOfTitle(CurrentTimespan)
        SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
        SourceEmailCol = FindColumnIndexByTitle("Email", SourceSheet.index)
        SourceUsageCol = FindColumnIndexByTitle("ContactEvents", SourceSheet.index)
        Set SearchRange = GetSearchRange(SourceSheet, SourceEmailCol)
        'checks to through every email on thissheet
        For j = 2 To LastRow
            'searches for matching email
            Set SearchResult = SearchRange.Find(thisSheet.Cells(j, EmailCol))
            'copies over usage number if there is a match
            If Not SearchResult Is Nothing Then
                thisSheet.Cells(j, i) = SourceSheet.Cells(SearchResult.Row, SourceUsageCol)
            End If
        Next j
    Next i
    
    

End Sub
'empty table with domains. needs the user table filled out
Sub PopulateExistingDomainUsageTable()

    Dim DomainCol As Integer, LastCol As Integer, i As Integer, j As Long
    Dim SearchRange As Range, SearchResult As Range, k As Long
    Dim thisSheet As Worksheet, CurrentTimespan As String, SourceSheet As Worksheet
    Dim SourceLastRow As Long, SourceDomainCol As Integer, SourceLastCol As Integer
    Dim LastRow As Long, TotalUsage As Integer, FirstAddress As String
    
    'connects to current sheet
    Set thisSheet = ActiveSheet
    LastCol = GetLastUsedColumnByIndex(thisSheet.index)
    EmailCol = FindColumnIndexByTitle("Domain", ActiveSheet.index)
    LastRow = GetLastUsedRowByIndex(thisSheet.index)
    Set SourceSheet = Sheets("Summary")
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    SourceDomainCol = FindColumnIndexByTitle("Domain", SourceSheet.index)
    Set SearchRange = GetSearchRange(SourceSheet, SourceDomainCol)
    TotalUsage = 0
    
    For i = 2 To LastCol
        For j = 2 To LastRow
            Set SearchResult = SearchRange.Find(thisSheet.Cells(j, 1))
            If Not SearchResult Is Nothing Then
                FirstAddress = SearchResult.Address
                Do
                    TotalUsage = TotalUsage + SourceSheet.Cells(SearchResult.Row, i + 1)
                    Set SearchResult = SearchRange.FindNext(SearchResult)
                Loop While (Not SearchResult Is Nothing) And (SearchResult.Address <> FirstAddress)
            End If
            thisSheet.Cells(j, i) = TotalUsage
            TotalUsage = 0
        Next j
    Next i
    
End Sub

Sub SoloDomainsFromSheet()

    Dim SummarySheet As Worksheet, SoloDomainSheet As Worksheet
    Dim i As Long, SummaryLastRow As Long, j As Long, SoloLastRow As Long
    Dim SearchRange As Range, SearchResult As Range
    Dim SheetName As String, CurrentDomain As String
    
    Set SummarySheet = ActiveSheet
    SheetName = InputBox("Sheet name containing domains to solo")
    Set SoloDomainSheet = Sheets(SheetName)
    Set SearchRange = GetSearchRange(SoloDomainSheet, 1)
    SummaryLastRow = GetLastUsedRowByIndex(SummarySheet.index)
    
    
    With SummarySheet
        For i = SummaryLastRow To 2 Step -1
            CurrentDomain = .Cells(i, 2)
            Set SearchResult = SearchRange.Find(CurrentDomain)
            If SearchResult Is Nothing Then
                Rows(i).EntireRow.Delete
                'Rows(SearchResult.Row + ":" + SearchResult.Row).Select
                'Application.CutCopyMode = False
                'Selection.Delete Shift:=xlUp
            End If
        Next i
    End With

    
End Sub

Sub AnalyzeVacations()
    Dim SummarySheet As Worksheet, TargetSheet As Worksheet, VacationCounter As Integer
    Dim UsageArr() As Integer, SummaryLastRow As Long, i As Long, FirstUsageIndex As Integer
    Dim j As Integer, IsVacation As Boolean
    
    Set TargetSheet = ActiveSheet
    Set SummarySheet = GetSheetByTitle("Summary")
    SummaryLastRow = GetLastUsedRowByIndex(SummarySheet.index)
    For i = 2 To SummaryLastRow
    'For i = 2 To 6
        VacationCounter = 0
        IsVacation = False
        UsageArr = LoadUsageArray(SummarySheet, i, 3)
        FirstUsageIndex = GetFirstUsageIndex(UsageArr)
        'If FirstUsageIndex <> UBound(UsageArr) Then
            For j = FirstUsageIndex To UBound(UsageArr)
                If UsageArr(j) = 0 And Not IsVacation Then
                    VacationCounter = 1
                    IsVacation = True
                ElseIf UsageArr(j) = 0 And IsVacation Then
                    VacationCounter = VacationCounter + 1
                ElseIf UsageArr(j) <> 0 And IsVacation Then
                    IsVacation = False
                    TargetSheet.Cells(2, VacationCounter) = TargetSheet.Cells(2, VacationCounter) + 1
                    VacationCounter = 0
                End If
            Next j
        'End If
        If IsVacation Then
            TargetSheet.Cells(7, 1) = TargetSheet.Cells(7, 1) + 1
            TargetSheet.Cells(7, 2) = TargetSheet.Cells(7, 2) + VacationCounter
            TargetSheet.Cells(3, VacationCounter) = TargetSheet.Cells(3, VacationCounter) + 1
            'If VacationCounter = 41 Then
                'MsgBox ("The problem is " + i)
            'End If
        End If
        
    Next i

End Sub

Function LoadUsageArray( _
     SummarySheet As Worksheet, _
     GatherRow As Long, _
     StartCol As Integer) As Integer()

    Dim LastCol As Integer, i As Integer, UsageArr() As Integer
    
    LastCol = GetLastUsedColumnByIndex(SummarySheet.index)
    ReDim UsageArr(0 To (LastCol - StartCol))
    For i = StartCol To LastCol
        UsageArr(i - StartCol) = SummarySheet.Cells(GatherRow, i)
    Next i
    
    LoadUsageArray = UsageArr
    
End Function

Function GetFirstUsageIndex( _
     UsageArr() As Integer) As Integer
    Dim i As Integer
    
    For i = LBound(UsageArr) To UBound(UsageArr)
        If UsageArr(i) > 0 Then
            GetFirstUsageIndex = i
            Exit For
        End If
    Next i

End Function