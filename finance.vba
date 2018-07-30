Option Explicit
Public Const DaysInMonth As Double = 30.43

'finds the mrr of the opportunity by finding drr and multiplyig by const daysinmonth
Function GetMRROfOppty( _
     SourceSheet As Worksheet, _
     RowNumber As Long, AmountCol As Integer, _
     StartDateCol As Integer, _
     EndDateCol As Integer) As Double
    Dim DRR As Double
    
    With SourceSheet
        DRR = .Cells(RowNumber, AmountCol) / (.Cells(RowNumber, EndDateCol) _
         - .Cells(RowNumber, StartDateCol))
        GetMRROfOppty = DRR * DaysInMonth
         
    End With
End Function
'requires opportunity export from salesforce. creates a column titled real mrr that will be filled with the true mrr of that
Sub CreateMRRColumnnnnnnnn()
    Dim LastRow As Long, AmountCol As Integer, StartDateCol As Integer
    Dim EndDateCol As Integer, i As Long, SourceSheet As Worksheet
    Dim LastCol As Integer
    
    
    Set SourceSheet = ActiveSheet
    LastCol = GetLastUsedColumnByIndex(SourceSheet.index)
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    AmountCol = FindColumnIndexByTitle("Amount", SourceSheet.index)
    StartDateCol = FindColumnIndexByTitle("Contract Effective Date", SourceSheet.index)
    EndDateCol = FindColumnIndexByTitle("Contract End Date", SourceSheet.index)
    
    With SourceSheet
        .Cells(1, LastCol + 1) = "Real MRR"
        For i = 2 To LastRow
            .Cells(i, LastCol + 1) = GetMRROfOppty(SourceSheet, i, _
             AmountCol, StartDateCol, EndDateCol)
        Next i
    End With
    
End Sub
'generate a column titled "Real MRR" in the
Sub CreateMRRColumn( _
     SourceSheet As Worksheet, _
     AmountCol As Integer, _
     StartDateCol As Integer, _
     EndDateCol As Integer)
    Dim LastRow As Long
    Dim i As Long
    Dim LastCol As Integer

    LastCol = GetLastUsedColumnByIndex(SourceSheet.index)
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    With SourceSheet
        .Cells(1, LastCol + 1) = "Real MRR"
        For i = 2 To LastRow
            .Cells(i, LastCol + 1) = GetMRROfOppty(SourceSheet, i, _
             AmountCol, StartDateCol, EndDateCol)
        Next i
    End With
End Sub
Sub JustCreateMRRColMang()

    Dim SourceSheet As Worksheet
    Set SourceSheet = GetSheetByTitle("Closed")
    Call CreateMRRColumn(SourceSheet, 18, 15, 16)
    

End Sub

'assumes you are starting with a table with owner names on first column and monthly on first row.
'also assumes that you already have the real mrr column in source from running createmrrcolumn
'run from table sheet
Sub PopulateTeamMRRTable()
    Dim TableLastCol As Integer, SourceSheet As Worksheet, TableSheet As Worksheet
    Dim SourceSheetName As String, i As Integer, TableLastRow As Long, QueryDate As Date
    Dim SourceLastRow As Long, SourceMRRCol As Integer, SourceNameCol As Integer
    Dim SourceStartDateCol As Integer, SourceEndDateCol As Integer, SearchRange As Range
    Dim SearchResult As Range, SourceOwnerCol As Integer, j As Integer, FirstAddress As String
    Dim MRRTotal As Double, OwnerNameToSearch As String
    
    Set TableSheet = ActiveSheet
    SourceSheetName = "Closed"
    Set SourceSheet = GetSheetByTitle(SourceSheetName)
    
    TableLastCol = GetLastUsedColumnByIndex(TableSheet.index)
    TableLastRow = GetLastUsedRowByIndex(TableSheet.index)
    SourceStartDateCol = FindColumnIndexByTitle("Contract Effective Date", SourceSheet.index)
    SourceEndDateCol = FindColumnIndexByTitle("Contract End Date", SourceSheet.index)
    SourceMRRCol = FindColumnIndexByTitle("Real MRR", SourceSheet.index)
    SourceOwnerCol = FindColumnIndexByTitle("Opportunity Owner", SourceSheet.index)
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    Set SearchRange = GetSearchRange(SourceSheet, SourceOwnerCol)
    
    For i = 2 To TableLastCol
        QueryDate = TableSheet.Cells(1, i)
        For j = 2 To TableLastRow
            
            OwnerNameToSearch = TableSheet.Cells(j, 1)
            Set SearchResult = SearchRange.Find(OwnerNameToSearch)
            If Not SearchResult Is Nothing Then
                FirstAddress = SearchResult.Address
                Do
                    With SourceSheet
                        If (IsActiveDuring(.Cells(SearchResult.Row, SourceStartDateCol), _
                         .Cells(SearchResult.Row, SourceEndDateCol), QueryDate)) Then
                            MRRTotal = MRRTotal + .Cells(SearchResult.Row, SourceMRRCol)
                        End If
                    End With
                    Set SearchResult = SearchRange.FindNext(SearchResult)
                    
                Loop While (Not SearchResult Is Nothing) And (SearchResult.Address <> FirstAddress)
                TableSheet.Cells(j, i) = MRRTotal
                MRRTotal = 0
            End If
        Next j
        
    Next i
    
End Sub

Sub PopulateTeamNewMRRTable()
    Dim TableLastCol As Integer, SourceSheet As Worksheet, TableSheet As Worksheet
    Dim SourceSheetName As String, i As Integer, TableLastRow As Long, QueryDate As Date
    Dim SourceLastRow As Long, SourceMRRCol As Integer, CloseDate As Date
    Dim SearchRange As Range, SourceCloseDateCol As Integer
    Dim SearchResult As Range, SourceOwnerCol As Integer, j As Integer, FirstAddress As String
    Dim MRRTotal As Double, OwnerNameToSearch As String
    
    Set TableSheet = ActiveSheet
    SourceSheetName = "Closed"
    Set SourceSheet = GetSheetByTitle(SourceSheetName)
    
    'column and row assignments
    TableLastCol = GetLastUsedColumnByIndex(TableSheet.index)
    TableLastRow = GetLastUsedRowByIndex(TableSheet.index)
    SourceCloseDateCol = FindColumnIndexByTitle("Close Date", SourceSheet.index)
    SourceMRRCol = FindColumnIndexByTitle("Real MRR", SourceSheet.index)
    SourceOwnerCol = FindColumnIndexByTitle("Opportunity Owner", SourceSheet.index)
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    'set search range
    Set SearchRange = GetSearchRange(SourceSheet, SourceOwnerCol)
    
    For i = 2 To TableLastCol
        QueryDate = TableSheet.Cells(1, i)
        For j = 2 To TableLastRow
            
            OwnerNameToSearch = TableSheet.Cells(j, 1)
            Set SearchResult = SearchRange.Find(OwnerNameToSearch)
            If Not SearchResult Is Nothing Then
                FirstAddress = SearchResult.Address
                Do
                    With SourceSheet
                        CloseDate = .Cells(SearchResult.Row, SourceCloseDateCol)
                        If (IsActiveDuring(dhFirstDayInMonth(CloseDate), _
                         dhLastDayInMonth(CloseDate), QueryDate)) Then
                            MRRTotal = MRRTotal + .Cells(SearchResult.Row, SourceMRRCol)
                        End If
                    End With
                    Set SearchResult = SearchRange.FindNext(SearchResult)
                    
                Loop While (Not SearchResult Is Nothing) And (SearchResult.Address <> FirstAddress)
                TableSheet.Cells(j, i) = MRRTotal
                MRRTotal = 0
            End If
        Next j
        
    Next i
    
End Sub

Sub PopulateTeamNewVMRRTable()
    Dim TableLastCol As Integer, SourceSheet As Worksheet, TableSheet As Worksheet
    Dim SourceSheetName As String, i As Integer, TableLastRow As Long, QueryDate As Date
    Dim SourceLastRow As Long, SourceMRRCol As Integer, CloseDate As Date
    Dim SearchRange As Range, SourceCloseDateCol As Integer
    Dim SearchResult As Range, SourceOwnerCol As Integer, j As Integer, FirstAddress As String
    Dim MRRTotal As Double, OwnerNameToSearch As String
    
    Set TableSheet = ActiveSheet
    SourceSheetName = "Closed"
    Set SourceSheet = GetSheetByTitle(SourceSheetName)
    
    'column and row assignments
    TableLastCol = GetLastUsedColumnByIndex(TableSheet.index)
    TableLastRow = GetLastUsedRowByIndex(TableSheet.index)
    SourceCloseDateCol = FindColumnIndexByTitle("Close Date", SourceSheet.index)
    SourceMRRCol = FindColumnIndexByTitle("Real MRR", SourceSheet.index)
    SourceOwnerCol = FindColumnIndexByTitle("Opportunity Owner", SourceSheet.index)
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    'set search range
    Set SearchRange = GetSearchRange(SourceSheet, SourceOwnerCol)
    
    For i = 2 To TableLastRow
        QueryDate = TableSheet.Cells(i, 1)
        For j = 2 To TableLastCol
            
            OwnerNameToSearch = TableSheet.Cells(1, j)
            Set SearchResult = SearchRange.Find(OwnerNameToSearch)
            If Not SearchResult Is Nothing Then
                FirstAddress = SearchResult.Address
                Do
                    With SourceSheet
                        CloseDate = .Cells(SearchResult.Row, SourceCloseDateCol)
                        If (IsActiveDuring(dhFirstDayInMonth(CloseDate), _
                         dhLastDayInMonth(CloseDate), QueryDate)) Then
                            MRRTotal = MRRTotal + .Cells(SearchResult.Row, SourceMRRCol)
                        End If
                    End With
                    Set SearchResult = SearchRange.FindNext(SearchResult)
                    
                Loop While (Not SearchResult Is Nothing) And (SearchResult.Address <> FirstAddress)
                TableSheet.Cells(i, j) = MRRTotal
                MRRTotal = 0
            End If
        Next j
        
    Next i
    
End Sub

'requires real mrr to have been calculated and placed into a column in the source data. Or not, maybe just DRR
'opportunity data must come from SF export
'must be launched from a prepped account sheet which includes all accounts on left side with "Account" header. Rest will be filled in
Sub GetAccountExpectedRenewal()
    Dim AccountSheet As Worksheet, i As Long, AccountLastRow As Long
    Dim SearchRange As Range, SearchResult As Range, SourceSheet As Worksheet
    Dim DaysOfContract As Integer, TotalMRR As Double, TotalDRR As Double, ContractDuration As Integer
    Dim SourceSheetName As String, SourceAccountIDCol As Integer, SourceMRRCol As Integer
    Dim SourceStartDateCol As Integer, SourceEndDateCol As Integer, SourceTypeCol As Integer
    Dim SourceActiveEndDateCol As Integer, AccountEndDateCol As Integer, AccountExpectedRenAmtCol As Integer
    Dim AccountExpectedRenMRRCol As Integer, AccountContractLengthCol As Integer, AccountIDCol As Integer
    Dim FirstAddress As String, ThisActiveEndDate As Date, SourceContractDurationCol As Integer
    Dim AccountExpansionCtCol As Integer, ExpansionCount As Integer
    
    'setting up worksheets and sheet dimensions
    Set AccountSheet = ActiveSheet
    SourceSheetName = InputBox("Source sheet name")
    Set SourceSheet = GetSheetByTitle(SourceSheetName)
    AccountLastRow = GetLastUsedRowByIndex(AccountSheet.index)
    
    'Creating new columns in account sheet
    Call InsertColumnAtEnd(AccountSheet, "End Date")
    Call InsertColumnAtEnd(AccountSheet, "Expected Renewal Amount")
    Call InsertColumnAtEnd(AccountSheet, "Expected Renewal MRR")
    Call InsertColumnAtEnd(AccountSheet, "Contract Length")
    Call InsertColumnAtEnd(AccountSheet, "Expansion Count")
    
    'setting up column data
    SourceActiveEndDateCol = FindColumnIndexByTitle("Active End Date", SourceSheet.index)
    SourceStartDateCol = FindColumnIndexByTitle("Contract Effective Date", SourceSheet.index)
    SourceEndDateCol = FindColumnIndexByTitle("Contract End Date", SourceSheet.index)
    SourceTypeCol = FindColumnIndexByTitle("Type", SourceSheet.index)
    SourceAccountIDCol = FindColumnIndexByTitle("Account ID", SourceSheet.index)
    SourceMRRCol = FindColumnIndexByTitle("Real MRR", SourceSheet.index)
    SourceContractDurationCol = FindColumnIndexByTitle("Contract Duration", SourceSheet.index)
    AccountEndDateCol = FindColumnIndexByTitle("End Date", AccountSheet.index)
    AccountExpectedRenAmtCol = FindColumnIndexByTitle("Expected Renewal Amount", AccountSheet.index)
    AccountExpectedRenMRRCol = FindColumnIndexByTitle("Expected Renewal MRR", AccountSheet.index)
    AccountContractLengthCol = FindColumnIndexByTitle("Contract Length", AccountSheet.index)
    AccountIDCol = FindColumnIndexByTitle("Account ID", AccountSheet.index)
    AccountExpansionCtCol = FindColumnIndexByTitle("Expansion Count", AccountSheet.index)
    
    'set up search parameters
    Set SearchRange = GetSearchRange(SourceSheet, SourceAccountIDCol)
    
    ExpansionCount = 0
    TotalMRR = 0
    For i = 2 To AccountLastRow
        Set SearchResult = SearchRange.Find(AccountSheet.Cells(i, 1))
        If Not SearchResult Is Nothing Then
            FirstAddress = SearchResult.Address
            ThisActiveEndDate = SourceSheet.Cells(SearchResult.Row, SourceActiveEndDateCol)
            Do
                With SourceSheet
                    If (.Cells(SearchResult.Row, SourceEndDateCol) = ThisActiveEndDate) Then
                        ExpansionCount = ExpansionCount + 1
                        TotalMRR = TotalMRR + .Cells(SearchResult.Row, SourceMRRCol)
                        If (.Cells(SearchResult.Row, SourceTypeCol) = "New Business" Or .Cells(SearchResult.Row, SourceTypeCol) = "Renewal") Then
                            ContractDuration = .Cells(SearchResult.Row, SourceContractDurationCol)
                            DaysOfContract = .Cells(SearchResult.Row, SourceEndDateCol) - .Cells(SearchResult.Row, SourceStartDateCol)
                        End If
                    End If
                End With
                Set SearchResult = SearchRange.FindNext(SearchResult)
                    
            Loop While (Not SearchResult Is Nothing) And (SearchResult.Address <> FirstAddress)
            AccountSheet.Cells(i, AccountEndDateCol) = ThisActiveEndDate
            AccountSheet.Cells(i, AccountContractLengthCol) = ContractDuration
            AccountSheet.Cells(i, AccountExpectedRenMRRCol) = TotalMRR
            AccountSheet.Cells(i, AccountExpectedRenAmtCol) = TotalMRR / DaysInMonth * DaysOfContract
            AccountSheet.Cells(i, AccountExpansionCtCol) = ExpansionCount
            ExpansionCount = 0
            TotalMRR = 0
        End If
    Next i
    

End Sub

Sub PopulateAcctQtyChurn()

    Dim TableSheet As Worksheet, SourceSheet As Worksheet
    Dim TableLastCol As Integer
    Dim i As Integer, TableLastRow As Long, QueryDate As Date
    Dim SourceLastRow As Long, CloseDate As Date
    Dim StartSearchRange As Range, SourceCloseDateCol As Integer, EndSearchRange As Range
    Dim SearchResult As Range, j As Integer, FirstAddress As String
    Dim ActiveEndDateCol As Integer, EffectiveDateCol As Integer, EndDateCol As Integer
    Dim OwnerNameToSearch As String, MRRTotal As Double

    Set TableSheet = ActiveSheet
    Set SourceSheet = GetSheetByTitle("Closed")
    
    'column and row assignments
    TableLastCol = GetLastUsedColumnByIndex(TableSheet.index)
    TableLastRow = GetLastUsedRowByIndex(TableSheet.index)
    ActiveEndDateCol = FindColumnIndexByTitle("Active End Date", SourceSheet.index)
    EffectiveDateCol = FindColumnIndexByTitle("Contract Effective Date", SourceSheet.index)
    EndDateCol = FindColumnIndexByTitle("Contract End Date", SourceSheet.index)
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    'set search range
    Set StartSearchRange = GetSearchRange(SourceSheet, EffectiveDateCol)
    
    
    For i = 2 To TableLastCol
        QueryDate = TableSheet.Cells(1, i)
        For j = 2 To TableLastRow
            
            OwnerNameToSearch = TableSheet.Cells(j, 1)
            Set SearchResult = StartSearchRange.Find(OwnerNameToSearch)
            If Not SearchResult Is Nothing Then
                FirstAddress = SearchResult.Address
                Do
                    With SourceSheet
                        CloseDate = .Cells(SearchResult.Row, SourceCloseDateCol)
                        If (IsActiveDuring(dhFirstDayInMonth(CloseDate), _
                         dhLastDayInMonth(CloseDate), QueryDate)) Then
                            MRRTotal = MRRTotal + 1
                        End If
                    End With
                    Set SearchResult = StartSearchRange.FindNext(SearchResult)
                    
                Loop While (Not SearchResult Is Nothing) And (SearchResult.Address <> FirstAddress)
                TableSheet.Cells(j, i) = MRRTotal
                MRRTotal = 0
            End If
        Next j
        
    Next i
    

End Sub

Sub PopulateCustQtyChurn()

End Sub
'requires a "Closed" sheet
Sub PopulateTableMaster()
    Dim cSFLink As New cSFOpptyLink, i As Integer
    Dim SourceSheet As Worksheet, TableSheet As Worksheet
    Dim QueryDate As Date, StopDate As Date
    
    'currently stopdate will be set to TODAY, but modify later to make it so you can view this from any date
    StopDate = Date
    
    Set TableSheet = GetSheetByTitle("testtable")
    Set SourceSheet = GetSheetByTitle("Closed")
    
    'should insert column indexes into column title array, -1 if not found
    Call EstablishColLink(SourceSheet, cSFLink)
    'will check to see if real mrr has been calculated. calculate and establish link if not present.
    If (cSFLink.FindColumnIndex("Real MRR") = -1) Then
        Call CreateMRRColumn(SourceSheet, cSFLink.FindColumnIndex("Amount"), cSFLink.FindColumnIndex("Contract Effective Date"), cSFLink.FindColumnIndex("Contract End Date"))
        Call EstablishColLink(SourceSheet, cSFLink)
    End If
    
    If (cSFLink.FindColumnIndex("Expansion Link") = -1) Then
        Call LinkExpansionsToParent(SourceSheet, cSFLink)
        Call EstablishColLink(SourceSheet, cSFLink)
    End If
    
    If (cSFLink.FindColumnIndex("Previous ID") = -1) Then
        Call LinkRenewalsToPrevious(SourceSheet, cSFLink)
        Call EstablishColLink(SourceSheet, cSFLink)
    End If
    
    If (cSFLink.FindColumnIndex("Contract MRR") = -1) Then
        Call PopulateContractMRR(SourceSheet, cSFLink)
        Call EstablishColLink(SourceSheet, cSFLink)
    End If
    
    If (cSFLink.FindColumnIndex("Expected Renewal MRR") = -1) Then
        Call PopulateExpectedRenewalMRR(SourceSheet, cSFLink)
        Call EstablishColLink(SourceSheet, cSFLink)
        Call PopulateExpectedRenewal(SourceSheet, cSFLink)
    End If
    
    If (cSFLink.FindColumnIndex("Contract Amount") = -1) Then
        Call PopulateContractAmount(SourceSheet, cSFLink)
    End If
    
    'connection confirmation
    'For i = 0 To 10
        'TableSheet.Cells(i + 4, 2) = cSFLink.ColumnTitles(i)
        'TableSheet.Cells(i + 4, 3) = cSFLink.ColumnIndex(i)
    'Next i
    
    'has to be replaced with something more dynamic
    For i = 2 To GetLastUsedColumnByIndex(TableSheet.index)
        QueryDate = TableSheet.Cells(1, i)
        If (IsSameMonthYear(QueryDate, StopDate)) Then
            TableSheet.Cells(2, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "TOT")
            TableSheet.Cells(3, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "RENTOT")
            TableSheet.Cells(5, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "NEW")
            TableSheet.Cells(6, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "RNEW")
            TableSheet.Cells(8, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "LOST")
            TableSheet.Cells(9, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "REC")
            TableSheet.Cells(13, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "REX")
            TableSheet.Cells(17, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "EXNB")
            TableSheet.Cells(18, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "NEXNB")
            TableSheet.Cells(19, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "EXR")
            TableSheet.Cells(20, i) = GetMRRForThisMonthByType(SourceSheet, cSFLink, StopDate, "NEXR")
            
        Else
            TableSheet.Cells(2, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "TOT")
            TableSheet.Cells(3, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "RENTOT")
            TableSheet.Cells(5, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "NEW")
            TableSheet.Cells(6, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "RNEW")
            TableSheet.Cells(8, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "LOST")
            TableSheet.Cells(9, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "REC")
            TableSheet.Cells(13, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "REX")
            TableSheet.Cells(14, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "RCON")
            TableSheet.Cells(13, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "REX")
            TableSheet.Cells(17, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "EXNB")
            TableSheet.Cells(18, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "NEXNB")
            TableSheet.Cells(19, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "EXR")
            TableSheet.Cells(20, i) = GetMRRForMonthByType(SourceSheet, cSFLink, QueryDate, "NEXR")
        End If
    Next i
    
    

End Sub
'TOT: active total, NEW: new mrr, LOST: mrr lost, RNEW: mrr renewed, REX: mrr of renewal expansion, RCON: Mrr of renewal contraction
'RENTOT: total mrr of renewals, REC: mrr lost that is later recovered, EXNB: expansion on new biz, EXR: expansion on renewal
'NEXNB: new expansion on new biz, NEXR: new expansion renewal
Function GetMRRForMonthByType( _
     SourceSheet As Worksheet, _
     cSFLink As cSFOpptyLink, _
     QueryDate As Date, _
     MRRType As String) As Double
    Dim SourceLastRow As Integer, i As Integer, TempMRR As Double
    Dim TempDelta As Double
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    TempMRR = 0
    
    With SourceSheet
        For i = 2 To SourceLastRow
        
            Select Case MRRType
                Case "TOT"
                    If (IsBetweenDates(dhFirstDayInMonth(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date"))), _
                     .Cells(i, cSFLink.FindColumnIndex("Contract End Date")), QueryDate)) And _
                     (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract End Date")), QueryDate) = False) Then
                        TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                    End If
                Case "NEW"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), QueryDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) <> "Renewal") Then
                        TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                    End If
                Case "LOST"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract End Date")), QueryDate)) And _
                     (QueryDate <= Date) Then
                    'If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Active End Date")), QueryDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Active End Date")) = .Cells(i, cSFLink.FindColumnIndex("Contract End Date"))) Then
                        TempMRR = TempMRR - .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                    End If
                Case "RNEW"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), QueryDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Renewal") Then
                        TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                    End If
                Case "REX"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), QueryDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Renewal") Then
                        TempDelta = .Cells(i, cSFLink.FindColumnIndex("Real MRR")) - .Cells(i, cSFLink.FindColumnIndex("Expected Renewal MRR"))
                        If TempDelta > 0 Then
                            TempMRR = TempMRR + TempDelta
                        End If
                    End If
                Case "RCON"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), QueryDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Renewal") Then
                        TempDelta = .Cells(i, cSFLink.FindColumnIndex("Real MRR")) - .Cells(i, cSFLink.FindColumnIndex("Expected Renewal MRR"))
                        If TempDelta < 0 Then
                            TempMRR = TempMRR + TempDelta
                        End If
                    End If
                Case "RENTOT"
                    If (IsBetweenDates(dhFirstDayInMonth(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date"))), _
                     .Cells(i, cSFLink.FindColumnIndex("Contract End Date")), QueryDate)) And (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Renewal") Then
                        TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                    End If
                Case "REC"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract End Date")), QueryDate)) Then
                        If (.Cells(i, cSFLink.FindColumnIndex("Contract End Date")) <> .Cells(i, cSFLink.FindColumnIndex("Active End Date"))) Then
                            TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                        End If
                    End If
                Case "EXNB"
                    If (IsBetweenDates(dhFirstDayInMonth(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date"))), _
                     .Cells(i, cSFLink.FindColumnIndex("Contract End Date")), QueryDate)) And _
                     (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract End Date")), QueryDate) = False) Then
                        If (GetExpansionParentType(SourceSheet, cSFLink, i) = "New Business") And _
                         .Cells(i, cSFLink.FindColumnIndex("Type")) = "Expansion" Then
                            TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                        End If
                    End If
                Case "EXR"
                    If (IsBetweenDates(dhFirstDayInMonth(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date"))), _
                     .Cells(i, cSFLink.FindColumnIndex("Contract End Date")), QueryDate)) And _
                     (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract End Date")), QueryDate) = False) Then
                        If (GetExpansionParentType(SourceSheet, cSFLink, i) = "Renewal") And _
                         .Cells(i, cSFLink.FindColumnIndex("Type")) = "Expansion" Then
                            TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                        End If
                    End If
                Case "NEXNB"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), QueryDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Expansion") Then
                        If (GetExpansionParentType(SourceSheet, cSFLink, i) = "New Business") Then
                            TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                        End If
                    End If
                Case "NEXR"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), QueryDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Expansion") Then
                        If (GetExpansionParentType(SourceSheet, cSFLink, i) = "Renewal") Then
                            TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                        End If
                    End If
            End Select
        Next i
    End With
    GetMRRForMonthByType = TempMRR
End Function
'TOT: cumulative, NEW: new mrr, LOST: mrr lost
Function GetMRRForThisMonthByType( _
     SourceSheet As Worksheet, _
     cSFLink As cSFOpptyLink, _
     StopDate As Date, _
     MRRType As String _
    ) As Double
    Dim CloseDateSearchRange As Range, SearchResult As Range
    Dim SourceLastRow As Integer, i As Integer, TempMRR As Double
    Dim TempDelta As Double
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    TempMRR = 0
    
    With SourceSheet
        For i = 2 To SourceLastRow
        
            Select Case MRRType
                Case "TOT"
                    If (IsBetweenDates(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")) - 1, _
                     .Cells(i, cSFLink.FindColumnIndex("Contract End Date")), StopDate)) Then
                     'And _
                     'Not (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract End Date")), StopDate)) Then
                        TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                    End If
                Case "NEW"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), StopDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")) < StopDate) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) <> "Renewal") Then
                        TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                    End If
                Case "LOST"
                    'If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Active End Date")), StopDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Active End Date")) < StopDate) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Active End Date")) = .Cells(i, cSFLink.FindColumnIndex("Contract End Date"))) Then
                     
                     If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract End Date")), StopDate)) And _
                      (.Cells(i, cSFLink.FindColumnIndex("Contract End Date")) < StopDate) Then
                        TempMRR = TempMRR - .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                    End If
                Case "RNEW"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), StopDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Renewal") And _
                     (.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")) < StopDate) Then
                        TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                    End If
                Case "REX"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), StopDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Renewal") And _
                     (.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")) < StopDate) Then
                        TempDelta = .Cells(i, cSFLink.FindColumnIndex("Real MRR")) - .Cells(i, cSFLink.FindColumnIndex("Expected Renewal MRR"))
                        If TempDelta > 0 Then
                            TempMRR = TempMRR + TempDelta
                        End If
                    End If
                Case "RCON"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), StopDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Renewal") And _
                     (.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")) < StopDate) Then
                        TempDelta = .Cells(i, cSFLink.FindColumnIndex("Real MRR")) - .Cells(i, cSFLink.FindColumnIndex("Expected Renewal MRR"))
                        If TempDelta < 0 Then
                            TempMRR = TempMRR + TempDelta
                        End If
                    End If
                Case "RENTOT"
                    If (IsBetweenDates(dhFirstDayInMonth(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date"))), _
                     .Cells(i, cSFLink.FindColumnIndex("Contract End Date")), StopDate)) And (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Renewal") And _
                     (.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")) < StopDate) Then
                        TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                    End If
                Case "REC"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract End Date")), StopDate)) And _
                      (.Cells(i, cSFLink.FindColumnIndex("Contract End Date")) < StopDate) Then
                        If (.Cells(i, cSFLink.FindColumnIndex("Contract End Date")) <> .Cells(i, cSFLink.FindColumnIndex("Active End Date"))) Then
                            TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                        End If
                    End If
                Case "EXNB"
                    If (IsBetweenDates(dhFirstDayInMonth(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date"))), _
                     .Cells(i, cSFLink.FindColumnIndex("Contract End Date")), StopDate)) And _
                     (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract End Date")), StopDate) = False) Then
                        If (GetExpansionParentType(SourceSheet, cSFLink, i) = "New Business") And _
                         .Cells(i, cSFLink.FindColumnIndex("Type")) = "Expansion" Then
                            TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                        End If
                    End If
                Case "EXR"
                    If (IsBetweenDates(dhFirstDayInMonth(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date"))), _
                     .Cells(i, cSFLink.FindColumnIndex("Contract End Date")), StopDate)) And _
                     (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract End Date")), StopDate) = False) Then
                        If (GetExpansionParentType(SourceSheet, cSFLink, i) = "Renewal") And _
                         .Cells(i, cSFLink.FindColumnIndex("Type")) = "Expansion" Then
                            TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                        End If
                    End If
                Case "NEXNB"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), StopDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Expansion") Then
                        If (GetExpansionParentType(SourceSheet, cSFLink, i) = "New Business") Then
                            TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                        End If
                    End If
                Case "NEXR"
                    If (IsSameMonthYear(.Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")), StopDate)) And _
                     (.Cells(i, cSFLink.FindColumnIndex("Type")) = "Expansion") Then
                        If (GetExpansionParentType(SourceSheet, cSFLink, i) = "Renewal") Then
                            TempMRR = TempMRR + .Cells(i, cSFLink.FindColumnIndex("Real MRR"))
                        End If
                    End If
            End Select
        Next i
    End With
    GetMRRForThisMonthByType = TempMRR
End Function

Sub LinkExpansionsToParent( _
     SourceSheet As Worksheet, _
     cSFLink As cSFOpptyLink)
    Dim IsMatchFound As Boolean
    Dim LastRow As Integer, i As Integer, AccountIDSearchRange As Range
    Dim SearchResult As Range, FirstAddress As String, TempEndDate As Date
    Dim TypeCol As Integer, ContractEndCol As Integer, ExpansionLinkCol As Integer
    Dim OpportunityIDCol As Integer, LastCol As Integer
    
    'wauugh i am here to make this not work
    LastCol = GetLastUsedColumnByIndex(SourceSheet.index)
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    
    TypeCol = cSFLink.FindColumnIndex("Type")
    ContractEndCol = cSFLink.FindColumnIndex("Contract End Date")
    ExpansionLinkCol = LastCol + 1
    OpportunityIDCol = cSFLink.FindColumnIndex("Opportunity ID")
    
    Set AccountIDSearchRange = GetSearchRange(SourceSheet, 1)
    IsMatchFound = False
    With SourceSheet
    .Cells(1, ExpansionLinkCol) = "Expansion Link"
    For i = 2 To LastRow
        IsMatchFound = False
        If .Cells(i, TypeCol) = "Expansion" Then
            Set SearchResult = AccountIDSearchRange.Find(.Cells(i, cSFLink.FindColumnIndex("Account ID")))
            TempEndDate = .Cells(i, ContractEndCol)
            If Not SearchResult Is Nothing Then
                FirstAddress = SearchResult.Address
                Do
                    If (TempEndDate = .Cells(SearchResult.Row, ContractEndCol)) And _
                     (.Cells(SearchResult.Row, TypeCol) = "New Business" Or .Cells(SearchResult.Row, TypeCol) = "Renewal") Then
                        .Cells(i, ExpansionLinkCol) = .Cells(SearchResult.Row, OpportunityIDCol)
                        IsMatchFound = True
                    End If
                    Set SearchResult = AccountIDSearchRange.FindNext(SearchResult)
                    
                Loop While (Not SearchResult Is Nothing) And (SearchResult.Address <> FirstAddress)
            End If
            
            If Not IsMatchFound Then
                .Cells(i, ExpansionLinkCol) = "ERROR"
            End If
        Else
            .Cells(i, ExpansionLinkCol) = "Parent"
            
        End If
        
    Next i
    End With
    

End Sub

Sub LinkRenewalsToPrevious( _
     SourceSheet As Worksheet, _
     cSFLink As cSFOpptyLink)
    Dim IsMatchFound As Boolean
    Dim LastRow As Integer, i As Integer, AccountIDSearchRange As Range
    Dim SearchResult As Range, FirstAddress As String, TempEndDate As Date
    Dim AccountIDArr As Variant, ThisAccountID As String, ThisDateArray() As Date
    Dim SortedDateArray() As Date, j As Integer, LastCol As Integer
    Dim ThisOpportunityID As String, PreviousIDCol As Integer

    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    LastCol = GetLastUsedColumnByIndex(SourceSheet.index)
    PreviousIDCol = LastCol + 1
    
    'collect array of unique account ids
    AccountIDArr = AggregateColumnToUniqueArr(cSFLink.FindColumnIndex("Account ID"), SourceSheet)
    
    
    Set AccountIDSearchRange = GetSearchRange(SourceSheet, cSFLink.FindColumnIndex("Account ID"))
    
    SourceSheet.Cells(1, PreviousIDCol) = "Previous ID"
    For i = LBound(AccountIDArr) To UBound(AccountIDArr)
        ThisAccountID = AccountIDArr(i)
        ThisDateArray = AggregateContractEndDates(cSFLink, SourceSheet, ThisAccountID)
        
        If isArrayEmpty(ThisDateArray) Then
            SourceSheet.Cells(SearchResult.Row, PreviousIDCol) = "Initial"
            GoTo NextIteration
        Else
            SortedDateArray = SortDateArray(ThisDateArray, "Asc")
        End If
        
        
        For j = LBound(SortedDateArray) To UBound(SortedDateArray)
            Set SearchResult = AccountIDSearchRange.Find(ThisAccountID)
            If Not SearchResult Is Nothing Then
                FirstAddress = SearchResult.Address
                Do
                    If SourceSheet.Cells(SearchResult.Row, cSFLink.FindColumnIndex("Type")) = "New Business" Or _
                     SourceSheet.Cells(SearchResult.Row, cSFLink.FindColumnIndex("Type")) = "Renewal" Then
                     
                        If SourceSheet.Cells(SearchResult.Row, cSFLink.FindColumnIndex("Contract End Date")) = SortedDateArray(j) Then
                            If j = LBound(SortedDateArray) Then
                                ThisOpportunityID = SourceSheet.Cells(SearchResult.Row, cSFLink.FindColumnIndex("Opportunity ID"))
                                SourceSheet.Cells(SearchResult.Row, PreviousIDCol) = "Initial"
                                Exit Do
                            Else
                                SourceSheet.Cells(SearchResult.Row, PreviousIDCol) = ThisOpportunityID
                                ThisOpportunityID = SourceSheet.Cells(SearchResult.Row, cSFLink.FindColumnIndex("Opportunity ID"))
                                Exit Do
                            End If
                        End If
                    End If
                    Set SearchResult = AccountIDSearchRange.FindNext(SearchResult)
                Loop While (FirstAddress <> SearchResult.Address) And (Not SearchResult Is Nothing)
            End If
        Next j
NextIteration:
    Next i



End Sub

Function AggregateColumnToUniqueArr( _
     TargetColumn As Integer, _
     SourceSheet As Worksheet) As Variant
    Dim AggregateArr As Variant, SourceLastRow As Integer
    Dim counter As Integer, ArrayInitialized As Boolean
    Dim i As Integer
    
    ArrayInitialized = False
    counter = 0
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    'initialize
    
    
    For i = 2 To SourceLastRow
        If Not ArrayInitialized Then
            AggregateArr = InitializeArray(AggregateArr, SourceSheet.Cells(i, TargetColumn))
            SourceSheet.Cells(1, 2) = AggregateArr
            ArrayInitialized = True
        End If
        AggregateArr = AddValueToArrIfUnique(AggregateArr, SourceSheet.Cells(i, TargetColumn))
    Next i
    AggregateColumnToUniqueArr = AggregateArr
    

End Function

Function AggregateDateColumnToUniqueArr( _
     TargetColumn As Integer, _
     SourceSheet As Worksheet) As Date()
    Dim AggregateArr() As Date, SourceLastRow As Integer
    Dim counter As Integer, ArrayInitialized As Boolean
    Dim i As Integer
    
    ArrayInitialized = False
    counter = 1
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    'initialize
    
    
    For i = 2 To SourceLastRow
        If Not ArrayInitialized Then
            ReDim AggregateArr(0)
            AggregateArr(0) = SourceSheet.Cells(i, 1)
            SourceSheet.Cells(1, 2) = AggregateArr
            ArrayInitialized = True
        End If
        If IsArrayValueUnique(AggregateArr, SourceSheet.Cells(i, TargetColumn)) Then
            ReDim Preserve AggregateArr(UBound(AggregateArr) + 1)
            AggregateArr = AddDateValueToArrIfUnique(AggregateArr, SourceSheet.Cells(i, TargetColumn))
        End If
    Next i
    AggregateDateColumnToUniqueArr = AggregateArr
    

End Function

Function AggregateColumnOfMatchingField( _
     cSFLink As cSFOpptyLink, _
     SourceSheet As Worksheet, _
     MatchValue As Variant, _
     MatchValueCol As Integer, _
     DataCol As Integer) As Variant
         
End Function

Function AggregateContractEndDates( _
     cSFLink As cSFOpptyLink, _
     SourceSheet As Worksheet, _
     AccountID As String) As Date()
    Dim DateArray() As Date, AccountIDSearchRange As Range, SearchResult As Range
    Dim FirstAddress As String, ThisType As String, counter As Integer
    
    counter = 0
    Set AccountIDSearchRange = GetSearchRange(SourceSheet, cSFLink.FindColumnIndex("Account ID"))
    
    Set SearchResult = AccountIDSearchRange.Find(AccountID)
    If Not SearchResult Is Nothing Then
        FirstAddress = SearchResult.Address
        Do
            ThisType = SourceSheet.Cells(SearchResult.Row, cSFLink.FindColumnIndex("Type"))
            If ThisType = "New Business" Or ThisType = "Renewal" Then
                ReDim Preserve DateArray(counter)
                DateArray(counter) = SourceSheet.Cells(SearchResult.Row, cSFLink.FindColumnIndex("Contract End Date"))
                counter = counter + 1
            End If
            Set SearchResult = AccountIDSearchRange.FindNext(SearchResult)
        Loop While (FirstAddress <> SearchResult.Address) And (Not SearchResult Is Nothing)
    End If
    
    AggregateContractEndDates = DateArray
    
    
End Function

Sub PopulateContractMRR( _
     SourceSheet As Worksheet, _
     cSFLink As cSFOpptyLink)
    Dim TotalMRR As Double, LastRow As Integer, i As Integer, RealMRRCol As Integer
    Dim ThisOpptyID As String, ThisType As String, ExpansionLinkSearchRange As Range
    Dim SearchResult As Range, ContractMRRCol As Integer, FirstAddress As String
    
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    ContractMRRCol = GetLastUsedColumnByIndex(SourceSheet.index) + 1
    RealMRRCol = cSFLink.FindColumnIndex("Real MRR")
    Set ExpansionLinkSearchRange = GetSearchRange(SourceSheet, cSFLink.FindColumnIndex("Expansion Link"))
    
    
    With SourceSheet
        .Cells(1, ContractMRRCol) = "Contract MRR"
        For i = 2 To LastRow
            TotalMRR = 0
            ThisType = .Cells(i, cSFLink.FindColumnIndex("Type"))
            If ThisType = "New Business" Or ThisType = "Renewal" Then
                ThisOpptyID = .Cells(i, cSFLink.FindColumnIndex("Opportunity ID"))
                TotalMRR = .Cells(i, RealMRRCol)
                Set SearchResult = ExpansionLinkSearchRange.Find(ThisOpptyID)
                If Not SearchResult Is Nothing Then
                    FirstAddress = SearchResult.Address
                    Do
                        TotalMRR = TotalMRR + .Cells(SearchResult.Row, RealMRRCol)
                        Set SearchResult = ExpansionLinkSearchRange.FindNext(SearchResult)
                    Loop While (FirstAddress <> SearchResult.Address) And (Not SearchResult Is Nothing)
                    
                End If
                .Cells(i, ContractMRRCol) = TotalMRR
            End If
        Next i
    End With
End Sub

Sub PopulateExpectedRenewalMRR( _
     SourceSheet As Worksheet, _
     cSFLink As cSFOpptyLink)
    Dim TotalRenewal As Double, LastRow As Integer, i As Integer, SearchResult As Range
    Dim ThisType As String, OpptyIDSearchRange As Range, LastOpptyID As String
    Dim FirstAddress As String
    
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    Set OpptyIDSearchRange = GetSearchRange(SourceSheet, cSFLink.FindColumnIndex("Opportunity ID"))
    SourceSheet.Cells(1, GetLastUsedColumnByIndex(SourceSheet.index) + 1) = "Expected Renewal MRR"
    Call EstablishColLink(SourceSheet, cSFLink)
    With SourceSheet
        For i = 2 To LastRow
            TotalRenewal = 0
            ThisType = .Cells(i, cSFLink.FindColumnIndex("Type"))
            If ThisType = "Renewal" Then
                LastOpptyID = .Cells(i, cSFLink.FindColumnIndex("Previous ID"))
                Set SearchResult = OpptyIDSearchRange.Find(LastOpptyID)
                If Not SearchResult Is Nothing Then
                    FirstAddress = SearchResult.Address
                    Do
                        TotalRenewal = TotalRenewal + .Cells(SearchResult.Row, cSFLink.FindColumnIndex("Contract MRR"))
                        Set SearchResult = OpptyIDSearchRange.FindNext(SearchResult)
                    Loop While (FirstAddress <> SearchResult.Address) And (Not SearchResult Is Nothing)
                End If
                .Cells(i, cSFLink.FindColumnIndex("Expected Renewal MRR")) = TotalRenewal
            End If
        Next i
    End With
    
End Sub


Sub PopulateExpectedRenewal( _
     SourceSheet As Worksheet, _
     cSFLink As cSFOpptyLink)
    Dim ThisRenewal As Double, LastRow As Integer, i As Integer
    Dim ThisType As String
    
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    With SourceSheet
        For i = 2 To LastRow
            ThisType = .Cells(i, cSFLink.FindColumnIndex("Type"))
            If ThisType = "Renewal" Then
                ThisRenewal = (.Cells(i, cSFLink.FindColumnIndex("Expected Renewal MRR")) / DaysInMonth) * _
                 (.Cells(i, cSFLink.FindColumnIndex("Contract End Date")) - .Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")))
                .Cells(i, cSFLink.FindColumnIndex("Expected Renewal")) = ThisRenewal
            End If
            
        Next i
    End With
     
End Sub

Sub PopulateContractAmount( _
     SourceSheet As Worksheet, _
     cSFLink As cSFOpptyLink)
    Dim LastRow As Integer, i As Integer
    
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    SourceSheet.Cells(1, GetLastUsedColumnByIndex(SourceSheet.index) + 1) = "Contract Amount"
    Call EstablishColLink(SourceSheet, cSFLink)
    
    With SourceSheet
        For i = 2 To LastRow
            If .Cells(i, cSFLink.FindColumnIndex("Contract MRR")) > 0 Then
                .Cells(i, cSFLink.FindColumnIndex("Contract Amount")) = (.Cells(i, cSFLink.FindColumnIndex("Contract MRR")) / DaysInMonth) * _
                 (.Cells(i, cSFLink.FindColumnIndex("Contract End Date")) - .Cells(i, cSFLink.FindColumnIndex("Contract Effective Date")))
            End If
        Next i
    End With
End Sub

Function GetExpansionParentType( _
     SourceSheet As Worksheet, _
     cSFLink As cSFOpptyLink, _
     RowNumber As Integer) As String
    Dim LastRow As Integer, ExpansionParentID As String, OpptyIDSearchRange As Range
    Dim SearchResult As Range
    
    Set OpptyIDSearchRange = GetSearchRange(SourceSheet, cSFLink.FindColumnIndex("Opportunity ID"))
    ExpansionParentID = SourceSheet.Cells(RowNumber, cSFLink.FindColumnIndex("Expansion Link"))
    If ExpansionParentID <> "" Then
        Set SearchResult = OpptyIDSearchRange.Find(ExpansionParentID)
        If Not SearchResult Is Nothing Then
            GetExpansionParentType = SourceSheet.Cells(SearchResult.Row, cSFLink.FindColumnIndex("Type"))
        End If
    Else
        GetExpansionParentType = "ERROR"
    End If
    
End Function

Sub CreateProductUpdateSheet()
    Dim ProductNumberArr(3, 1) As Variant, cSFLink As New cSFOpptyLink, i As Integer, j As Integer
    Dim SourceSheet As Worksheet, TableSheet As Worksheet, ThisOpportunityID As String, SingleSeatCount As Integer
    Dim LastRow As Integer, ThisAmount As Double, DualSeatCount As Integer, InvoiceAmount As Double
    Dim TableCounter As Integer, SinglePrice As Double, SingleQuantity As Integer, DualPrice As Double
    Dim DualQuantity As Integer, AdjustPrice As Double, AdjustQuantity As Integer, ThisOpportunityName As String
    
    ProductNumberArr(0, 0) = "DUAL"
    ProductNumberArr(0, 1) = "01ud0000007nm9gAAA"
    ProductNumberArr(1, 0) = "SINGLE"
    ProductNumberArr(1, 1) = "01ud0000007nm9hAAA"
    ProductNumberArr(2, 0) = "ADJUST"
    ProductNumberArr(2, 1) = "01ud0000007nmA0AAI"
    
    Set SourceSheet = GetSheetByTitle("Closed")
    Call CreateNewSheet("Product Updater")
    Set TableSheet = GetSheetByTitle("Product Updater")
    TableCounter = 2
    Call EstablishColLink(SourceSheet, cSFLink)
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    'PricebookentryID, Quantity, Amount, opportunity name
    TableSheet.Cells(1, 1) = "Opportunity ID"
    TableSheet.Cells(1, 2) = "PricebookentryID"
    TableSheet.Cells(1, 3) = "Quantity"
    TableSheet.Cells(1, 4) = "Sales Price"
    TableSheet.Cells(1, 5) = "Opportunity Name"
    
    For i = 2 To LastRow
        ThisOpportunityID = SourceSheet.Cells(i, cSFLink.FindColumnIndex("Opportunity ID"))
        ThisOpportunityName = SourceSheet.Cells(i, cSFLink.FindColumnIndex("Opportunity Name"))
        ThisAmount = SourceSheet.Cells(i, cSFLink.FindColumnIndex("Amount"))
        SingleSeatCount = SourceSheet.Cells(i, cSFLink.FindColumnIndex("Single Product Licenses"))
        DualSeatCount = SourceSheet.Cells(i, cSFLink.FindColumnIndex("Dual Product Licenses"))
        DualQuantity = 0
        SingleQuantity = 0
        AdjustQuantity = 0
        InvoiceAmount = ThisAmount / (Round(1 / SourceSheet.Cells(i, cSFLink.FindColumnIndex("Billing Term Modifier")), 0))
        
        AdjustPrice = InvoiceAmount
        AdjustQuantity = (Round(1 / SourceSheet.Cells(i, cSFLink.FindColumnIndex("Billing Term Modifier")), 0)) - 1
        DualQuantity = SourceSheet.Cells(i, cSFLink.FindColumnIndex("Dual Product Licenses"))
        SingleQuantity = SourceSheet.Cells(i, cSFLink.FindColumnIndex("Single Product Licenses"))
        SinglePrice = InvoiceAmount / (DualQuantity + SingleQuantity)
        DualPrice = InvoiceAmount / (DualQuantity + SingleQuantity)
        
        If (DualQuantity > 0) Then
            TableSheet.Cells(TableCounter, 1) = ThisOpportunityID
            TableSheet.Cells(TableCounter, 2) = ReturnSecondDimension(ProductNumberArr, "DUAL")
            TableSheet.Cells(TableCounter, 3) = DualQuantity
            TableSheet.Cells(TableCounter, 4) = DualPrice
            TableSheet.Cells(TableCounter, 5) = ThisOpportunityName
            TableCounter = TableCounter + 1
        End If
        If (SingleQuantity > 0) Then
            TableSheet.Cells(TableCounter, 1) = ThisOpportunityID
            TableSheet.Cells(TableCounter, 2) = ReturnSecondDimension(ProductNumberArr, "SINGLE")
            TableSheet.Cells(TableCounter, 3) = SingleQuantity
            TableSheet.Cells(TableCounter, 4) = SinglePrice
            TableSheet.Cells(TableCounter, 5) = ThisOpportunityName
            TableCounter = TableCounter + 1
        End If
        If (AdjustQuantity > 0) Then
            TableSheet.Cells(TableCounter, 1) = ThisOpportunityID
            TableSheet.Cells(TableCounter, 2) = ReturnSecondDimension(ProductNumberArr, "ADJUST")
            TableSheet.Cells(TableCounter, 3) = AdjustQuantity
            TableSheet.Cells(TableCounter, 4) = AdjustPrice
            TableSheet.Cells(TableCounter, 5) = ThisOpportunityName
            TableCounter = TableCounter + 1
        End If
        
    Next i
    
    
End Sub

Function ReturnSecondDimension(TwoDimArray As Variant, OneDimQuery As String) As String
    Dim i As Integer
    
    For i = LBound(TwoDimArray) To UBound(TwoDimArray)
        If TwoDimArray(i, 0) = OneDimQuery Then
            ReturnSecondDimension = TwoDimArray(i, 1)
        End If
    Next i
End Function