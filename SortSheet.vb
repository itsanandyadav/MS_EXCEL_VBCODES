Sub SortSheets()
    Dim i As Integer
    Dim j As Integer
    'Dim k As Integer
    
    'Count the number of worksheets and store the number in variable "n"
    n = Application.Sheets.Count
    
    'Do the following loop for each worksheet again
    For i = 1 To n
        For j = i + 1 To n
            If LCase(Sheets(i).Name) > LCase(Sheets(j).Name) Then
                Sheets(j).Move Before:=Sheets(i)
            End If
        Next j
    Next i

End Sub
