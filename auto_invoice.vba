
Option Explicit

Function GetXeroSpecificHeaders() As String()
    GetXeroSpecificHeaders = Split("*ContactName,EmailAddress,*InvoiceNumber" + _
    ",*InvoiceDate,*DueDate,*Description,*Quantity,*UnitAmount,*AccountCode" + _
    ",*TaxType", ",")
End Function

Function GetSFSpecificHeaders() As String()
    GetSFSpecificHeaders = Split("Account ID,Opportunity ID,Amount,Single Product Licenses" + _
    ",Dual Product Licenses,Contract Effective Date,Contract End Date,New AP Email,New AP Name" + _
    ",Billing,Contract Duration,Invoice Sent,Net Payment", ",")
End Function
Private Function GetSavePath() As String
    GetSavePath = "C:\Users\John\OneDrive\Documents\Connectifier\Invoices\Automated\Export\"
End Function



'generate all invoices necessary
Sub GenerateXeroInvoiceSheet()
    
    Dim CreateCount As Long
    Dim i As Long
    Dim HeaderArr() As String
    
    HeaderArr = GetXeroSpecificHeaders()
    'Call CreateNewSheetOfTitleWithHeaders("Xero Invoice Uploader", HeaderArr)
    For i = 0 To CreateCount
        
    Next i
    
End Sub

'fill in proper account names
Sub XeroPreparations()
    
End Sub

'delete SF account names when ready to create export csv
Sub CleanSFReferences(TargetSheet As Worksheet)

    Dim SFDeleteHeadersArr() As String
    Dim i As Integer
    
    'collects headers that will be deleted. modify in function
    SFDeleteHeadersArr = GetSFSpecificHeaders()
    
    'confirm ubound will not fail if array numbering is different. idea: create a function to do that
    For i = 0 To UBound(SFDeleteHeadersArr)
        'Call DeleteColumnByColHeader(TargetSheet, SFDeleteHeadersArr(i))
    Next i
    
End Sub

'get latest invoice number from inputbox. will be negative if it will be auto determined

Function GetLatestInvoiceNumber()
    Dim LatestInvoiceNumber As Integer
    GetLatestInvoiceNumber = LatestInvoiceNumber
End Function
Sub ExportAllSheets()
    
End Sub
'will return the number of invoices that will be associated with the opportunity
Function GetInvoiceQuantity(BillingTerm As String, IsTwoYear As Boolean)
    
    Dim YearMultiplier As Integer
    
    'Modify this section if multiyear becomes a thing. Instead of passing a boolean, pass a number of years and annual billing structure.
    If IsTwoYear Then
        YearMultiplier = 2
    Else
        YearMultiplier = 1
    End If
    
    Select Case BillingTerm
        Case "Upfront"
            GetInvoiceQuantity = 1 * YearMultiplier
        Case "Semi-Annual"
            GetInvoiceQuantity = 2 * YearMultiplier
        Case "Quarterly"
            GetInvoiceQuantity = 4 * YearMultiplier
        Case "Monthly"
            GetInvoiceQuantity = 12 * YearMultiplier
        Case "Annual"
            GetInvoiceQuantity = YearMultiplier
    End Select
End Function
'Create csv of all relevant sheets
Sub CreateCSV(ExportSheet As Worksheet, Optional CSVName As String = "")
    Dim SavePath As String
    
    SavePath = GetSavePath

    If CSVName = "" Then
        'if name is not specified, csv will be named for sheet
        
        ExportSheet.Copy
        With ActiveSheet
        .SaveAs SavePath + .Name + ".csv", xlCSV
        End With
    Else
        ExportSheet.Copy
        With ActiveSheet
        .SaveAs SavePath + CSVName + ".csv", xlCSV
        End With
    End If

End Sub
Sub GenerateInvoicesForOppty(InvoiceSheet As Worksheet, StartLine As Long, _
    HeaderArr() As String, ThisOppty As cInvoiceOppty, LatestInvoice As Integer)
    Dim InvoiceQuantity As Integer
    Dim i As Integer, LineItemsPerInvoice As Integer
    'find number of invoices
    
    'InvoiceQuantity = GetInvoiceQuantity(ThisOppty.BillingType)
    
    If ThisOppty.SingleQuantity > 0 And ThisOppty.DualQuantity > 0 Then
        LineItemsPerInvoice = 2
    Else
        LineItemsPerInvoice = 1
    End If
    'for every invoice, generate a header with unique inv number and then add in additional line items
    For i = 1 To InvoiceQuantity
        Call GenerateUniqueInvoiceFields(InvoiceSheet, StartLine + 1 + ((i - 1) * LineItemsPerInvoice), HeaderArr, ThisOppty, LatestInvoice + i)
    Next i
    
    
End Sub
Sub GenerateUniqueInvoiceFields(InvoiceSheet As Worksheet, StartLine As Long, _
    HeaderArr() As String, ThisOppty As cInvoiceOppty, InvoiceNumber As Integer)
    
    
    
End Sub
Sub PopulateLineItems(InvoiceSheet As Worksheet, InsertLine As Long, HeaderArr() As String, _
    ThisOppty As cInvoiceOppty)
    
End Sub
Function FindDueDates( _
     ThisOppty As cInvoiceOppty, _
     BillDates() As Date) As Date()
    Dim i As Integer
    Dim DueDates() As Date
    
    For i = 0 To UBound(BillDates)
        DueDates(i) = BillDates(i) + ThisOppty.NetPayment
    Next i
    
    FindDueDates = DueDates()
    
End Function
Function FindBillDates(ThisOppty As cInvoiceOppty) As Date()
    Dim BillDates() As Date, ContractDays As Integer
    Dim InvoiceDays As Integer, InvoiceQuantity As Integer
    Dim InvoiceMonths As Integer, i As Integer
    
    'InvoiceQuantity = GetInvoiceQuantity(ThisOppty.BillingType)
    ContractDays = ThisOppty.EndDate - ThisOppty.EffectiveDate
    InvoiceDays = ContractDays / InvoiceQuantity
    InvoiceMonths = 12 / InvoiceQuantity
    For i = 1 To InvoiceQuantity
        If i = 1 Then
            BillDates(0) = ThisOppty.EffectiveDate
        Else
            BillDates(i - 1) = ThisOppty.EffectiveDate + ((i - 1) * InvoiceDays)
        End If
    Next i
    FindBillDates = BillDates
End Function



Sub GenerateSFUploader()

End Sub

Function IsTwoYear( _
     StartDate As Date, _
     EndDate As Date) As Boolean

    If Year(EndDate) = (Year(StartDate) + 2) Then
        IsTwoYear = True
    ElseIf Abs((EndDate - StartDate) - 730) < 30 Then
        If MsgBox("Should this be considered a two year?", vbYesNo) = vbYes Then
            IsTwoYear = True
        Else
            IsTwoYear = False
        End If
    Else
        IsTwoYear = False
    End If

End Function

Sub TestInvoiceFunction()
    Dim HeaderArr() As String, SpecificCol As String
    
    SpecificCol = InputBox("Name of specific column")
    HeaderArr = GetXeroSpecificHeaders()
    'Call CreateNewSheetOfTitleWithHeaders("Whatever", HeaderArr)
    With Sheets("Whatever")
        '.Cells(2, FindStringIndexInArray(HeaderArr, SpecificCol) + 1) = "Here"
    End With

End Sub
