
'Read from row and store. Need to have linked column index to each attribute
'create date stamped sheet for account data, returns index
Function CreateAccountDatabaseSheet( _
     Optional DatabaseName As String = "Account DB") As Integer
    Sheets.Add after:=Sheets(Sheets.count)
    Sheets(Sheets.count).Name = DatabaseName + " " + CStr(Sheets.count)
    CreateAccountDatabaseSheet = Sheets.count

End Function

Sub CreateHeadersInSheet( _
     TargetSheet As Worksheet, _
     ClassType As String)
    Dim HeaderArray() As String
    Dim i As Integer
    
    If ClassType = "Account" Then
        Dim TargetClass As cCFAccount
        'HeaderArray = TargetClass.GetHeaders()
        HeaderArray = Split("Account ID,Account Name,Terms,Opportunities,Total Seats,Active End Date", ",")
        
        For i = 1 To (UBound(HeaderArray(), 1) + 1)
            TargetSheet.Cells(1, i) = HeaderArray(i - 1)
        Next i
    End If
    
End Sub

Sub CreateDatabaseFromExport( _
     ClassType As String)
    Dim SourceSheet As Worksheet, DatabaseSheet As Worksheet
    Dim HeaderArray() As String
    Dim LastRowSource As Long, LastColSource As Integer
    Dim i As Integer, j As Integer
    Dim HeaderIndexArr(6) As Integer
    Dim HeaderSearchRange As Range, SearchResult As Range
    Dim SearchColumn As Integer, DatabaseRowCounter As Long
        
    HeaderArray = Split("Account ID,Account Name,Terms,Opportunities,Total Seats,Active End Date", ",")
    'set to be dynamic
    Set SourceSheet = Sheets(1)
    Set DatabaseSheet = Sheets(Sheets.count)
    LastRowSource = GetLastUsedRowByIndex(SourceSheet.index)
    LastColSource = GetLastUsedColumnByIndex(SourceSheet.index)
    Set HeaderSearchRange = GetHeaderRange(SourceSheet)
    
    'this collects indices of header locations. should be moved to its own function
    For j = 0 To UBound(HeaderArray, 1)
        Set SearchResult = HeaderSearchRange.Find(HeaderArray(j))
        If Not SearchResult Is Nothing Then
            HeaderIndexArr(j) = SearchResult.Column
        Else
            HeaderIndexArr(j) = -1
        End If
        
    Next j
    DatabaseRowCounter = 2
    For i = 2 To LastRowSource
        'if query is an account
        For j = 0 To UBound(HeaderArray, 1)
            If HeaderIndexArr(j) > 0 Then
                DatabaseSheet.Cells(DatabaseRowCounter, j + 1) = _
                 SourceSheet.Cells(i, HeaderIndexArr(j))
            End If
        Next j
        DatabaseRowCounter = DatabaseRowCounter + 1
    Next i
    
End Sub

Sub TestAccountCreation()
    
    Dim DatabaseIndex As Integer
    
    DatabaseIndex = CreateAccountDatabaseSheet
    Call CreateHeadersInSheet(Sheets(DatabaseIndex), "Account")
    Call CreateDatabaseFromExport("Account")

End Sub

Sub DisplayAccountData()
    Dim TestAccount As cCFAccount
    Dim QueryRow As Long
    
    QueryRow = InputBox("Which row?")
    
End Sub
