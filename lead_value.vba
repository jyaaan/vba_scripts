Option Explicit
'requires source generated from mongodb in "Source" sheet, CSV parsed
'also requires table sheet with copied dates titled "Table"
'also requires date sheet with csv parsed dates, seat count, and price "Dates"

Sub LEADGenerateUsageTable()
    Dim TableLastCol As Integer, SourceSheet As Worksheet, TableSheet As Worksheet
    Dim TableLastRow As Long, QueryWeekYear As String
    Dim SourceLastRow As Long, SourceEmailCol As Integer, SourceYearWeekCol As Integer
    Dim SourceContactsCol As Integer, SearchRange As Range
    Dim SearchResult As Range, FirstAddress As String
    Dim UserEmailToSearch As String, DateSheet As Worksheet, DateLastRow As Integer
    
    Application.DisplayAlerts = False
    Sheets("CreateList").Delete
    Set TableSheet = GetSheetByTitle("Table")
    Set SourceSheet = GetSheetByTitle("Source")

    Set DateSheet = GetSheetByTitle("Dates")
    
    Call CleanRawData(SourceSheet, DateSheet)
    
    Dim x As Integer
    With DateSheet
        x = 1
        
        While .Cells(x, 1) <> ""
            x = x + 1
        Wend
        DateLastRow = x - 1
    End With
    
    With DateSheet
    ' parse csv for both
        Call ParseCSVSheet(DateSheet)
        Call ParseCSVSheet(SourceSheet)
    End With
    
    'copy unique addresses and group
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    With SourceSheet
        Range(.Cells(1, 1), .Cells(SourceLastRow, 1)).AdvancedFilter _
         Action:=xlFilterCopy, CopyToRange:=TableSheet.Range("A2"), Unique:=True
    End With
    
    TableSheet.Rows(2).Delete
    SourceEmailCol = FindColumnIndexByTitle("userEmail", SourceSheet.index)
    SourceYearWeekCol = FindColumnIndexByTitle("yearWeek", SourceSheet.index)
    SourceContactsCol = FindColumnIndexByTitle("contacts", SourceSheet.index)
    TableLastCol = GetLastUsedColumnByIndex(TableSheet.index)
    TableLastRow = GetLastUsedRowByIndex(TableSheet.index)

    'group users
    With TableSheet
        .Range(.Cells(2, 1), .Cells(TableLastRow, 1)).Rows.Group
    End With
    
    Dim i As Integer, j As Integer
    Set SearchRange = GetSearchRange(SourceSheet, SourceEmailCol)
    
    With TableSheet
        For i = 2 To TableLastCol
            QueryWeekYear = .Cells(1, i)
            For j = 2 To TableLastRow
                UserEmailToSearch = .Cells(j, 1)
                Set SearchResult = SearchRange.Find(UserEmailToSearch)
                If Not SearchResult Is Nothing Then
                    FirstAddress = SearchResult.Address
                    Do
                        If (SourceSheet.Cells(SearchResult.Row, SourceYearWeekCol) = QueryWeekYear) Then
                            .Cells(j, i) = SourceSheet.Cells(SearchResult.Row, SourceContactsCol)
                        End If
                        Set SearchResult = SearchRange.FindNext(SearchResult)
                    Loop While (Not SearchResult Is Nothing) And (SearchResult.Address <> FirstAddress)
                    
                End If
            Next j
        Next i
    End With
    
    'dates stuff
    Dim DateLastCol As Integer
    Dim DateStartCol As Integer, DateEndCol As Integer
    Dim DateSeatsCol As Integer, DatePriceCol As Integer, DateISOStartCol As Integer, DateISOEndCol As Integer
    
    
    
    DateStartCol = FindColumnIndexByTitle("contractEffectiveDate", DateSheet.index)
    DateEndCol = FindColumnIndexByTitle("contractEndDate", DateSheet.index)
    DateSeatsCol = FindColumnIndexByTitle("numPaidSeats", DateSheet.index)
    DatePriceCol = FindColumnIndexByTitle("pricePerSeatMonthly", DateSheet.index)
    DateLastCol = GetLastUsedColumnByIndex(DateSheet.index)
    DateISOStartCol = DateLastCol + 1
    DateISOEndCol = DateISOStartCol + 1
    
    
    With DateSheet
        .Cells(1, DateISOStartCol) = "isoStartDate"
        .Cells(1, DateISOEndCol) = "isoEndDate"
        Call DateFormulaSetter(DateISOStartCol, DateStartCol, 2, DateLastRow, DateSheet)
        Call DateFormulaSetter(DateISOEndCol, DateEndCol, 2, DateLastRow, DateSheet)
    End With
    
    'Create new row on table sheet for contract updates
    With TableSheet
        .Cells(TableLastRow + 1, 1) = "Subscription Status"
        Call PopulateSubscriptionStatus(TableLastRow + 1, TableLastCol, TableSheet, DateSheet)
        .Cells(TableLastRow + 2, 1) = "User Count"
        Call PopulateActiveUsers(TableLastRow + 2, TableLastCol, TableSheet, DateSheet)
        .Cells(TableLastRow + 3, 1) = "Aggregate Contacts"
        Call AggContactsFormula(TableLastRow + 3, TableLastCol, TableSheet)
        .Cells(TableLastRow + 4, 1) = "Active Users"
        Call ActiveUsersFormula(TableLastRow + 4, TableLastCol, TableSheet)
        .Cells(TableLastRow + 5, 1) = "Average Contacts Per User"
        Call AvUsageFormula(TableLastRow + 5, TableLastCol, TableSheet)
    End With
    Application.DisplayAlerts = True
End Sub

Sub DateFormulaSetter( _
    DestinationCol As Integer, _
    SourceCol As Integer, _
    StartRow As Integer, _
    EndRow As Integer, _
    TargetSheet As Worksheet)
    
    Dim ColModifier As Integer
    
    
    ColModifier = SourceCol - DestinationCol
    With TargetSheet
        .Range(.Cells(StartRow, DestinationCol), _
         .Cells(EndRow, DestinationCol)).FormulaR1C1 = "=IF(AND(MONTH(RC[-4])=12,(INT((RC[-4]-DATE(YEAR(RC[-4]-WEEKDAY(RC[-4]-1)+4),1,3)+WEEKDAY(DATE(YEAR(RC[-4]-WEEKDAY(RC[-4]-1)+4),1,3))+5)/7-1))<5),(YEAR(RC[-4])+1)&""-""&TEXT(INT((RC[-4]-DATE(YEAR(RC[-4]-WEEKDAY(RC[-4]-1)+4),1,3)+WEEKDAY(DATE(YEAR(RC[-4]-WEEKDAY(RC[-4]-1)+4),1,3))+5)/7-1),""00""),YEAR(RC[-4])&""-""&(TEXT(INT((RC[-4]-DATE(YEAR(RC[-4]-WEEKDAY(RC[-4]-1)+4),1,3)+WEEKDAY(DATE(YEAR(RC[-4]-WEEKDAY(RC[-4]-1)+4),1,3))+5)/7-1),""00"")))"
    End With
End Sub

Sub PopulateSubscriptionStatus( _
     targetRow As Integer, _
     EndCol As Integer, _
     TableSheet As Worksheet, _
     DateSheet As Worksheet)
         
     Dim ISOStartCol As Integer, ISOEndCol As Integer, SeatCol As Integer
     Dim DateLastRow As Integer
     
     DateLastRow = GetLastUsedRowByIndex(DateSheet.index)
     
     ISOStartCol = FindColumnIndexByTitle("isoStartDate", DateSheet.index)
     ISOEndCol = FindColumnIndexByTitle("isoEndDate", DateSheet.index)
     SeatCol = FindColumnIndexByTitle("numPaidSeats", DateSheet.index)
     
     'initialize row
     With TableSheet
        .Range(.Cells(targetRow, 2), .Cells(targetRow, EndCol)) = ""
     End With
     
     Dim i As Integer, j As Integer, isContractActive As Boolean
     isContractActive = False
     
     
     For i = 2 To DateLastRow
        For j = 2 To EndCol
            If DateSheet.Cells(i, ISOStartCol) = TableSheet.Cells(1, j) Then
                If TableSheet.Cells(targetRow, j) = "End" Then
                    TableSheet.Cells(targetRow, j) = "Renew"
                    TableSheet.Cells(targetRow, j).Interior.Color = RGB(0, 0, 150)
                    isContractActive = True
                Else
                    TableSheet.Cells(targetRow, j) = "Start"
                    TableSheet.Cells(targetRow, j).Interior.Color = RGB(0, 0, 150)
                    isContractActive = True
                End If
            End If
            
            If DateSheet.Cells(i, ISOEndCol) = TableSheet.Cells(1, j) Then
                If (TableSheet.Cells(targetRow, j) <> "Start") And (TableSheet.Cells(targetRow, j) <> "Renew") Then
                    TableSheet.Cells(targetRow, j) = "End"
                    TableSheet.Cells(targetRow, j).Interior.Color = RGB(150, 0, 0)
                End If
                
                isContractActive = False
            End If
            
            If isContractActive Then
                If (TableSheet.Cells(targetRow, j) <> "Start") And (TableSheet.Cells(targetRow, j) <> "Renew") Then
                    TableSheet.Cells(targetRow, j).Interior.Color = RGB(0, 150, 0)
                    TableSheet.Cells(targetRow, j) = "Active"
                End If
            End If
        Next j
        isContractActive = False
     Next i
     
     
     
End Sub

Sub PopulateActiveUsers( _
     targetRow As Integer, _
     EndCol As Integer, _
     TableSheet As Worksheet, _
     DateSheet As Worksheet)
    
    Dim ISOStartCol As Integer, ISOEndCol As Integer, SeatCol As Integer
    Dim DateLastRow As Integer
     
    DateLastRow = GetLastUsedRowByIndex(DateSheet.index)
     
     ISOStartCol = FindColumnIndexByTitle("isoStartDate", DateSheet.index)
     ISOEndCol = FindColumnIndexByTitle("isoEndDate", DateSheet.index)
     SeatCol = FindColumnIndexByTitle("numPaidSeats", DateSheet.index)
    
    'initialize row
     With TableSheet
        .Range(.Cells(targetRow, 2), .Cells(targetRow, EndCol)) = ""
     End With
     
     Dim i As Integer, j As Integer, isContractActive As Boolean, thisPaidSeats As Integer
     isContractActive = False
     thisPaidSeats = 0
     
     For i = 2 To DateLastRow
        For j = 2 To EndCol
            If DateSheet.Cells(i, ISOStartCol) = TableSheet.Cells(1, j) Then
                thisPaidSeats = DateSheet.Cells(i, SeatCol)
                isContractActive = True
                If (TableSheet.Cells(targetRow, j) > 0) Then
                    TableSheet.Cells(targetRow, j) = TableSheet.Cells(targetRow, j) + thisPaidSeats
                Else
                    TableSheet.Cells(targetRow, j) = thisPaidSeats
                End If
            ElseIf DateSheet.Cells(i, ISOEndCol) = TableSheet.Cells(1, j) Then
                If (TableSheet.Cells(targetRow, j) > 0) Then
                    TableSheet.Cells(targetRow, j) = TableSheet.Cells(targetRow, j) + thisPaidSeats
                Else
                    TableSheet.Cells(targetRow, j) = thisPaidSeats
                End If
                thisPaidSeats = 0
                isContractActive = False
            ElseIf isContractActive Then
                If (TableSheet.Cells(targetRow, j) > 0) Then
                    TableSheet.Cells(targetRow, j) = TableSheet.Cells(targetRow, j) + thisPaidSeats
                Else
                    TableSheet.Cells(targetRow, j) = thisPaidSeats
                End If
            End If
            
        Next j
        
        isContractActive = False
     Next i
    

End Sub

Sub CleanRawData(SourceSheet As Worksheet, DateSheet As Worksheet)
    Dim thisSheet As Worksheet, counter As Integer, LastUsageRow As Integer
    Dim LastRow As Integer, DateStartRow As Integer
    
    Set thisSheet = GetSheetByTitle("RAWDATA")
    
    counter = 1
    With thisSheet
        While (Left(.Cells(counter, 1), 4) <> "user")
            counter = counter + 1
        Wend
        .Range(.Cells(1, 1), .Cells((counter - 1), 1)).EntireRow.Delete
    End With
    
    counter = 1
    
    With thisSheet
        While (Left(.Cells(counter, 1), 11) <> "contractEff")
            counter = counter + 1
        Wend
        LastUsageRow = counter - 2
        DateStartRow = LastUsageRow + 2
        LastRow = GetLastUsedRowByIndex(thisSheet.index)
        .Range(.Cells(1, 1), .Cells(LastUsageRow, 1)).Copy _
         Destination:=SourceSheet.Range("A1")
        .Range(.Cells(DateStartRow, 1), .Cells(LastRow, 1)).Copy _
         Destination:=DateSheet.Range("A1")
        
    End With
    

End Sub

Sub AggContactsFormula( _
     targetRow As Integer, _
     EndCol As Integer, _
     TableSheet As Worksheet)
    With TableSheet
        .Range(.Cells(targetRow, 2), .Cells(targetRow, EndCol)).FormulaR1C1 = _
         "=SUM(R2C:R[-2]C)"
    End With
    
    
End Sub

Sub ActiveUsersFormula( _
     targetRow As Integer, _
     EndCol As Integer, _
     TableSheet As Worksheet)
    
    With TableSheet
        .Range(.Cells(targetRow, 2), .Cells(targetRow, EndCol)).FormulaR1C1 = _
         "=COUNTIF(R2C:R[-3]C,"">0"")"
    End With
    
End Sub

Sub AvUsageFormula( _
     targetRow As Integer, _
     EndCol As Integer, _
     TableSheet As Worksheet)
    
    With TableSheet
        .Range(.Cells(targetRow, 2), .Cells(targetRow, EndCol)).FormulaR1C1 = _
         "=IFERROR(R[-2]C/R[-1]C,"""")"
    End With
    
End Sub