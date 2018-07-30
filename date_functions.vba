'Return the first day in the specified month as date
'if run  empty, return date of first day of month
Function dhFirstDayInMonth( _
     Optional dtmDate As Date = 0) As Date
    If dtmDate = 0 Then
        dtmDate = Date
    End If
    dhFirstDayInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate), 1)

End Function
'Return the last day in the specified month as date
'if run  empty, return date of last day of month
Function dhLastDayInMonth( _
     Optional dtmDate As Date = 0) As Date
    If dtmDate = 0 Then
        dtmDate = Date
    End If
    dhLastDayInMonth = DateSerial(Year(dtmDate), Month(dtmDate) + 1, 0)
    
End Function

Function NumberActiveDays( _
     EffDate As Date, _
     EndDate As Date, _
     QueryDate As Date) As Integer
    Dim Duration As Long
    Duration = DateDiff("d", EffDate, EndDate)
    Dim precut As Long, postcut As Long
    Dim QueryBegin As Date, QueryEnd As Date
    
    QueryBegin = dhFirstDayInMonth(QueryDate)
    QueryEnd = dhLastDayInMonth(QueryDate)
    
    'find days elapsed until beginning of query month
    '0 if contract initiates on or after beginning of query month
    If EffDate >= QueryBegin Then
        'if contract started on or after query date
        precut = 0
    Else
        precut = DateDiff("d", EffDate, QueryBegin)
    End If
        
    'find days elapsed until after end of query month
    '0 if contract terminates on or before beginning of query month
    If QueryEnd >= EndDate Then
        postcut = 1
    Else
        postcut = DateDiff("d", QueryEnd, EndDate)
    End If
    
    NumberActiveDays = Duration - precut - postcut + 1
    

End Function
'checks to see if the querydate is between the first day of startdate's month and last of end date's
Function IsActiveDuring( _
     StartDate As Date, _
     EndDate As Date, _
     QueryDate As Date) As Boolean
    Dim StartMonth As Date, EndMonth As Date
    
    StartMonth = dhFirstDayInMonth(StartDate)
    EndMonth = dhLastDayInMonth(EndDate - 1)
    
    If IsBetweenDates(StartMonth, EndMonth, QueryDate) Then
        IsActiveDuring = True
    Else
        IsActiveDuring = False
    End If
    
End Function
'true if the querydate is in the first or last month in the "spandate"s
Function IsMonthOverlap( _
     BeginSpanDate As Date, _
     EndSpanDate As Date, _
     QueryDate As Date) As Boolean
    'What the hell is this for?
    
    If (IsSameMonth(BeginSpanDate, QueryDate) And IsSameYear(BeginSpanDate, QueryDate)) Or _
    (IsSameMonth(EndSpanDate, QueryDate) And IsSameYear(EndSpanDate, QueryDate)) Then
    
        IsMonthOverlap = True
    Else
        IsMonthOverlap = False
    End If
        

End Function
Function IsBetweenDates( _
     StartDate As Date, _
     EndDate As Date, _
     QueryDate As Date) As Boolean
    
    If QueryDate >= StartDate And QueryDate <= EndDate Then
        IsBetweenDates = True
    Else
        IsBetweenDates = False
    End If

End Function
Function IsSameMonth( _
     date1 As Date, _
     date2 As Date) As Boolean

    If Month(date1) = Month(date2) Then
        IsSameMonth = True
    Else
        IsSameMonth = False
    End If
    
    IsSameMonth = (Month(date1) = Month(date2))
        
End Function

Function IsSameYear( _
     date1 As Date, _
     date2 As Date) As Boolean

    If Year(date1) = Year(date2) Then
        IsSameYear = True
    Else
        IsSameYear = False
    End If
    
End Function
Function ClosedInMonth( _
     CloseDate As Date, _
     QueryDate As Date)
    If (IsSameYear(CloseDate, QueryDate)) And (IsSameMonth(CloseDate, QueryDate)) Then
        ClosedInMonth = True
    Else
        ClosedInMonth = False
    End If
    
End Function

Function IsSameMonthYear( _
     date1 As Date, _
     date2 As Date) As Boolean
    If (IsSameYear(date1, date2)) And (IsSameMonth(date1, date2)) Then
        IsSameMonthYear = True
    Else
        IsSameMonthYear = False
    End If
End Function