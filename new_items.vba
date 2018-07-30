Option Explicit

Sub GetAccountList()

    Dim thisSheet As Worksheet, SourceSheet As Worksheet, i As Integer
    Dim ThisCounter As Integer, DomainCol As Integer, StatusCol As Integer
    Dim LastRow As Integer
    
    Set thisSheet = ActiveSheet
    Set SourceSheet = Sheets(FindSheetIndexByTitle("main2"))
    ThisCounter = 1
    DomainCol = FindColumnIndexByTitle("Domain", SourceSheet.index)
    StatusCol = FindColumnIndexByTitle("Status", SourceSheet.index)
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    With SourceSheet
        For i = 2 To LastRow
            If .Cells(i, StatusCol) = "ACCOUNT_ACTIVE" Then
                thisSheet.Cells(ThisCounter, 1) = .Cells(i, DomainCol)
                ThisCounter = ThisCounter + 1
            End If
        Next i
    End With
End Sub


Sub ScrubActiveAccounts()

    Dim MainSheet As Worksheet, AccountSheet As Worksheet, i As Integer
    Dim MainDomainCol As Integer, MainEmailCol As Integer, MainStatusCol As Integer
    Dim AccountDomainCol As Integer, AccountEmailCol As Integer, AccountStatusCol As Integer
    Dim LastRow As Integer, ThisCounter As Integer, AccountRange As Range, SearchResult As Range
    
    ThisCounter = 2
    Set MainSheet = Sheets(FindSheetIndexByTitle("main2"))
    Set AccountSheet = Sheets(FindSheetIndexByTitle("scrubbed"))
    LastRow = GetLastUsedRowByIndex(MainSheet.index)
    Set AccountRange = GetSearchRange(Sheets(FindSheetIndexByTitle("activeaccounts")), 1, False)
    
    
    MainDomainCol = FindColumnIndexByTitle("Domain", MainSheet.index)
    MainEmailCol = FindColumnIndexByTitle("Email", MainSheet.index)
    MainStatusCol = FindColumnIndexByTitle("Status", MainSheet.index)
    AccountDomainCol = FindColumnIndexByTitle("Domain", AccountSheet.index)
    AccountEmailCol = FindColumnIndexByTitle("Email", AccountSheet.index)
    AccountStatusCol = FindColumnIndexByTitle("Status", AccountSheet.index)
    
    For i = 2 To LastRow
        Set SearchResult = AccountRange.Find(MainSheet.Cells(i, MainDomainCol))
        If Not SearchResult Is Nothing Then
            AccountSheet.Cells(ThisCounter, AccountEmailCol) = MainSheet.Cells(i, MainEmailCol)
            AccountSheet.Cells(ThisCounter, AccountDomainCol) = MainSheet.Cells(i, MainDomainCol)
            AccountSheet.Cells(ThisCounter, AccountStatusCol) = MainSheet.Cells(i, MainStatusCol)
        End If
    Next i
    
End Sub

'extract emails that are in Admin format
Sub ExtractEmail()
    Dim thisSheet As Worksheet, SourceCol As Integer, EmailCol As Integer
    Dim IsBetweenPren As Boolean, EmailStorage As String, LastRow As Integer
    Dim i As Integer, PrenPosition() As Variant, AtPosition As Integer
    Dim ThisEntry As String
    
    Set thisSheet = ActiveSheet
    SourceCol = InputBox("Source Column")
    EmailCol = InputBox("Email Column (destination)")
    LastRow = GetLastUsedRowByIndex(thisSheet.index)
    
    For i = 2 To LastRow
        ThisEntry = thisSheet.Cells(i, SourceCol)
        AtPosition = InStr(ThisEntry, "@")
        If HasPren(ThisEntry) Then
            PrenPosition = LoadPrenPosition(ThisEntry)
            If IsAtBetweenPren(AtPosition, PrenPosition) Then
                EmailStorage = Mid(ThisEntry, PrenPosition(0) + 1, PrenPosition(1) - PrenPosition(0) - 1)
            Else
                EmailStorage = Left(ThisEntry, PrenPosition(0) - 1)
            End If
        Else
            EmailStorage = ThisEntry
        End If
        thisSheet.Cells(i, EmailCol) = EmailStorage
    Next i
    
End Sub


Function HasPren(QueryString As String) As Boolean

    If InStr(QueryString, "(") > 0 And InStr(QueryString, ")") > 0 Then
        HasPren = True
    Else
        HasPren = False
    End If

End Function

Function LoadPrenPosition(QueryString As String) As Variant()

    Dim PrePosition(0 To 1) As Variant
    
    PrePosition(0) = InStr(QueryString, "(")
    PrePosition(1) = InStr(QueryString, ")")
    LoadPrenPosition = PrePosition

End Function
'returns true if string has '@' between ()
'doesn't account for multiple or mismatched ()
'so requires: verify only one set of () - if there are multiple, check contents, make sure () match
Function IsAtBetweenPren( _
     AtPosition As Integer, _
     PrenPosition() As Variant) As Boolean
    If (AtPosition > PrenPosition(0)) And (AtPosition < PrenPosition(1)) Then
        IsAtBetweenPren = True
    Else
        IsAtBetweenPren = False
    End If
End Function

Sub PopulateBookings()
    Dim SourceSheet As Worksheet, TableSheet As Worksheet
    Dim cSFLink As New cSFOpptyLink, SourceLastRow As Integer
    Dim LastRow As Integer, i As Integer, ThisDate As Date
    Dim j As Integer, SearchRange As Range, SearchResult As Range
    Dim FirstAddress As String, TotalAmount As Double
    
    
    Set SourceSheet = GetSheetByTitle("Closed")
    Set TableSheet = GetSheetByTitle("Bookings")
    Call EstablishColLink(SourceSheet, cSFLink)
    
    LastRow = GetLastUsedRowByIndex(TableSheet.index)
    SourceLastRow = GetLastUsedRowByIndex(SourceSheet.index)
    
    Set SearchRange = GetSearchRange(SourceSheet, cSFLink.FindColumnIndex("Type"))
    
    For i = 2 To 4
        For j = 2 To LastRow
            TotalAmount = 0
            ThisDate = TableSheet.Cells(j, 1)
            Set SearchResult = SearchRange.Find(TableSheet.Cells(1, i))
            If Not SearchResult Is Nothing Then
                FirstAddress = SearchResult.Address
                Do
                    If IsSameMonthYear(ThisDate, SourceSheet.Cells(SearchResult.Row, cSFLink.FindColumnIndex("Close Date"))) Then
                        TotalAmount = TotalAmount + SourceSheet.Cells(SearchResult.Row, cSFLink.FindColumnIndex("Amount"))
                    End If
                    Set SearchResult = SearchRange.FindNext(SearchResult)
                Loop While (FirstAddress <> SearchResult.Address) And (Not SearchResult Is Nothing)
                TableSheet.Cells(j, i) = TotalAmount
                Set SearchResult = Nothing
            End If
        Next j
    Next i
    Set SearchRange = Nothing
End Sub

Sub GenerateCommissionStatements()
'takes in raw export from SF report and establishes connnection
'pulls in salesrep data
'future: ability to pick specific month/date
'future: move logic to SF custom objects
    Dim SourceSheet As Worksheet, cSFLink As New cSFOpptyLink
    Dim cSalesRep As cCFSalesRep
    
End Sub

Function GetBookingForRep( _
     cSFLink As cSFOpptyLink, _
     cSalesRep As cCFSalesRep, _
     TargetMonth As Date, _
     SourceSheet As Worksheet, _
     Optional BookingType As String = "Total") As Double
    Dim TotalBookings As Double, LastRow As Integer, i As Integer
    Dim SearchRange As Range, SearchResult As Range
    
    TotalBookings = 0
    LastRow = GetLastUsedRowByIndex(SourceSheet.index)
    Set SearchRange = GetSearchRange(SourceSheet, cSFLink.FindColumnIndex("Opportunity Owner"))
    
    With SourceSheet
        For i = 2 To LastRow
            If .Cells(i, cSFLink.FindColumnIndex("Opportunity Owner")) = cSalesRep.Name Then
                TotalBookings = TotalBookings + .Cells(i, cSFLink.FindColumnIndex("Amount"))
            End If
        Next i
    End With
    
    GetBookingForRep = TotalBookings

End Function


Sub MarkMatches2()
    Dim DomainSheet As Worksheet, TargetSheet As Worksheet
    Dim SearchRange As Range, SearchResult As Range
    Dim i As Long, LastRow As Long, thisDomain As String
    
    Set DomainSheet = GetSheetByTitle("Remaining2")
    Set TargetSheet = ActiveSheet
    
    LastRow = GetLastUsedRowByIndex(TargetSheet.index)
    Set SearchRange = GetSearchRange(DomainSheet, 1)
    
    
    For i = 2 To LastRow
        thisDomain = TargetSheet.Cells(i, 1)
        If thisDomain = "" Then
            thisDomain = "Yamashiro"
        End If
        Set SearchResult = SearchRange.Find(thisDomain, lookat:=xlWhole)
        If Not SearchResult Is Nothing Then
            TargetSheet.Cells(i, 4) = "X"
        End If
    Next i
    

End Sub