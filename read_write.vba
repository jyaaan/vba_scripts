Sub CreateCSV( _
     ExportSheet As Worksheet, _
     Optional CSVName As String = "")
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