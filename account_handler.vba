Function FindMRRLostInMonth( _
     QueryDate As Date, _
     Optional buffer As Integer = 30) As Double

    Dim i As Long
    Dim drrsum As Double
    Dim col As Integer, drrcol As Integer
    Dim actendcol As Integer, EndCol As Integer
    
    'declarations
    actendcol = FindColumnIndexByTitle("Active End Date", 1)
    EndCol = FindColumnIndexByTitle("Contract End Date", 1)
    drrcol = FindColumnIndexByTitle("DRR", 1)
    drrsum = 0
    
    With Sheets(1)
    For i = 2 To 1642
    'NumberOfRows (1)
    
    If IsActiveDuring(.Cells(i, EndCol) + buffer, .Cells(i, EndCol) + buffer, QueryDate) Then
        If (.Cells(i, actendcol) + 5 + buffer > .Cells(i, EndCol) + buffer) And _
        (.Cells(i, actendcol) - 5 + buffer < .Cells(i, EndCol) + buffer) Then
            drrsum = drrsum + .Cells(i, drrcol) _
             * 30
        End If
    End If
    Next i
    End With
    
    FindMRRLostInMonth = drrsum
    
End Function

Function FindMRRGainFromPipelineByDate( _
     QueryDate As Date) As Double

    'Returns total amount of pipline new business for given month
    
    Dim TypeColumnNum As Integer, CloseColumnNum As Integer, StageColumnNum As Integer
    Dim AmountColumnNum As Integer
    Dim MRRGainAmount As Double
    
    
    TypeColumnNum = FindColumnIndexByTitle("Type", 1)
    CloseColumnNum = FindColumnIndexByTitle("Close Date", 1)
    StageColumnNum = FindColumnIndexByTitle("Stage", 1)
    AmountColumnNum = FindColumnIndexByTitle("Amount", 1)
    
    
    
    Dim i As Long
    Amount = 0
    For i = 2 To 1642
    
        If IsSameYear(QueryDate, Sheets(1).Cells(i, CloseColumnNum)) And _
         IsSameMonth(QueryDate, Sheets(1).Cells(i, CloseColumnNum)) Then
            If IsStageValid(Sheets(1).Cells(i, StageColumnNum)) And (Sheets(1).Cells(i, TypeColumnNum) = "New Business") Then
                MRRGainAmount = MRRGainAmount + Sheets(1).Cells(i, AmountColumnNum) / 12
            End If
            
        End If
    
    Next i
    
    FindMRRGainFromPipelineByDate = Amount

End Function

Function FindRenMRRDueInMonth( _
     QueryDate As Date) As Double

    Dim i As Long
    Dim drrsum As Double
    Dim col As Integer, drrcol As Integer
    Dim actendcol As Integer, EndCol As Integer
    
    'declarations
    actendcol = FindColumnIndexByTitle("Active End Date", 1)
    EndCol = FindColumnIndexByTitle("Contract End Date", 1)
    drrcol = FindColumnIndexByTitle("DRR", 1)
    drrsum = 0
    
    With Sheets(1)
    For i = 2 To 1642
    'NumberOfRows (1)
    
        If IsActiveDuring(.Cells(i, EndCol) + buffer, .Cells(i, EndCol) + buffer, QueryDate) Then
            drrsum = drrsum + .Cells(i, drrcol) _
             * 30.44

        End If
    Next i
    End With
    
    FindRenMRRDueInMonth = drrsum

End Function

Function IsStageValid( _
     Stage As String) As Boolean
    
    If (Stage = "Other resolution" Or Stage = "Beat by competitor") Or _
     (Stage = "Limbo" Or Stage = "Trial Negative") Then
        IsStageValid = False
    Else
        IsStageValid = True
    
    End If
    
    
End Function