Sub ArrayBuilder( _
     ExistingArray As Variant, _
     NewValue As Variant)
    Dim PrevHiPosition As Integer
    Dim TempArray As Variant
    TempArray = ExistingArray
    PrevHiPosition = UBound(TempArray)
    ReDim Preserve TempArray(PrevHiPosition + 1)
    ExistingArray(PrevHiPosition + 1) = NewValue

End Sub
Function ArrayBuilderFunction( _
     ByVal ExistingArray As Variant, _
     NewValue As Variant) As Variant
    Dim PrevHiPosition As Integer
    Dim TempArray As Variant
    
    TempArray = ExistingArray
    PrevHiPosition = UBound(TempArray)
    ReDim Preserve TempArray(PrevHiPosition + 1)
    TempArray(PrevHiPosition + 1) = NewValue
    ArrayBuilderFunction = TempArray
    
End Function
Sub Array2DBuilder( _
     ExistingArray As Variant, _
     NewValue As Variant)
    Dim PrevHiPosition As Integer
    
    PrevHiPosition = UBound(ExistingArray)
    ReDim Preserve ExistingArray((PrevHiPosition), 1)
    ExistingArray(PrevHiPosition, 0) = NewValue(0, 0)
    ExistingArray(PrevHiPosition, 1) = NewValue(0, 1)

End Sub

Sub Fill2DArray( _
     LargerArray As Variant, _
     PreviousArray As Variant)
    Dim i As Integer, j As Integer
    Dim FirstBound As Integer, SecondBound As Integer
    
    FirstBound = UBound(PreviousArray, 1)
    SecondBound = UBound(PreviousArray, 2)
    
End Sub

Function IsArrayValueUnique( _
     ExistingArray As Variant, _
     QueryValue As Variant) As Boolean
    Dim i As Integer
    
    For i = LBound(ExistingArray) To UBound(ExistingArray)
        If ExistingArray(i) = QueryValue Then
            IsArrayValueUnique = False
            Exit Function
        End If
    Next i
    IsArrayValueUnique = True
End Function

Function AddValueToArrIfUnique( _
     ExistingArray As Variant, _
     QueryValue As Variant) As Variant()
    Dim ArrUBound As Integer
    ArrUBound = UBound(ExistingArray)
    If IsArrayValueUnique(ExistingArray, QueryValue) Then
        ReDim Preserve ExistingArray(ArrUBound)
        ExistingArray = ArrayBuilderFunction(ExistingArray, QueryValue)
    End If
    AddValueToArrIfUnique = ExistingArray
End Function

Function AddDateValueToArrIfUnique( _
     ExistingArray() As Date, _
     QueryValue As Date) As Date()
    Dim ArrUBound As Integer, TempArray() As Date
    
    ArrUBound = UBound(ExistingArray)
    If IsArrayValueUnique(ExistingArray, QueryValue) Then
        ReDim Preserve ExistingArray(ArrUBound)
        ExistingArray = ArrayBuilderFunction(ExistingArray, QueryValue)
    End If
    AddDateValueToArrIfUnique = ExistingArray
End Function
'will clear array
Function InitializeArray( _
     ExistingArray As Variant, _
     FirstValue As Variant) As Variant()
    'Clear ExistingArray
    ReDim ExistingArray(0)
    ExistingArray(0) = FirstValue
    InitializeArray = ExistingArray
End Function
Sub Initialize2DArray( _
     ExistingArray As Variant, _
     FirstValue As Variant)

    ReDim ExistingArray(0, 1)
    ExistingArray(0, 0) = FirstValue(0, 0)
    ExistingArray(0, 1) = FirstValue(0, 1)

End Sub
Sub Sort2DByFirstIndex( _
     ArrayToSort As Variant)
    Dim TempArray As Variant, TempValue As Variant
    Dim HighScore As Variant, i As Integer, HasSorted As Boolean
    
    HasSorted = False
    
    Call Initialize2DArray(TempArray, ArrayToSort)
    HighScore = TempArray(LBound(TempArray), 0)


    For i = (LBound(ArrayToSort) + 1) To UBound(ArrayToSort)
        If (ArrayToSort(i, 0) >= HighScore) Then
            ReDim TempValue(0, 1)
            TempValue(0, 0) = ArrayToSort(i, 0)
            TempValue(0, 1) = ArrayToSort(i, 1)
            Call Array2DBuilder(TempArray, TempValue)
            HighScore = ArrayToSort(i, 0)
        Else
            ReDim TempValue(0, 1)
            TempValue(0, 0) = TempArray(i - 1, 0)
            TempValue(0, 1) = TempArray(i - 1, 1)
            TempArray(i - 1, 0) = ArrayToSort(i, 0)
            TempArray(i - 1, 1) = ArrayToSort(i, 1)
            Call Array2DBuilder(TempArray, TempValue)
            HasSorted = True
        End If
    Next i
    If HasSorted Then
        Call Sort2DByFirstIndex(TempArray)
    End If
    'Return
    ArrayToSort = TempArray
End Sub

Sub PrintArray( _
     ArrayToPrint As Variant, _
     TargetSheet As Worksheet, _
     StartRow As Integer, _
     StartCol As Integer, _
     Optional DirectionVorH As String = "V")
    Dim StartPos As Integer
    Dim i As Integer, j As Integer
    
    For i = LBound(ArrayToPrint) To UBound(ArrayToPrint)
        If DirectionVorH = "V" Then
            TargetSheet.Cells(StartRow + (i - LBound(ArrayToPrint)), StartCol) = ArrayToPrint(i)
        Else
            TargetSheet.Cells(StartRow, StartCol + (i - LBound(ArrayToPrint))) = ArrayToPrint(i)
        End If
    Next
    
End Sub

Function SortDateArray( _
     DateArray() As Date, _
     Optional DirectionAscOrDec As String = "Asc") As Date()
    Dim TempStorage() As Date, TempStorage2() As Date, TempDate As Date
    Dim i As Integer, counter As Integer, IsChanged As Boolean
    IsChanged = False
    counter = 1
    ReDim TempStorage2(UBound(DateArray))
    TempStorage2 = DateArray
    ReDim TempStorage(0)
    TempStorage(0) = TempStorage2(LBound(TempStorage2))
    
    Do
        IsChanged = False
        counter = 1
        ReDim TempStorage(0)
        TempStorage(0) = TempStorage2(LBound(TempStorage2))
        For i = (LBound(TempStorage2) + 1) To UBound(TempStorage2)
            If TempStorage(counter - 1) > TempStorage2(i) Then
                TempDate = TempStorage(counter - 1)
                ReDim Preserve TempStorage(counter)
                TempStorage(counter - 1) = TempStorage2(i)
                TempStorage(counter) = TempDate
                IsChanged = True
                counter = counter + 1
            Else
                ReDim Preserve TempStorage(counter)
                TempStorage(counter) = TempStorage2(i)
                counter = counter + 1
            End If
        Next i
        TempStorage2 = TempStorage
    Loop While (IsChanged)
    SortDateArray = TempStorage
    
End Function
     
Public Function isArrayEmpty(parArray As Variant) As Boolean
'Returns false if not an array or dynamic array that has not been initialised (ReDim) or has been erased (Erase)

    If IsArray(parArray) = False Then isArrayEmpty = True
    On Error Resume Next
    If UBound(parArray) < LBound(parArray) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False

End Function