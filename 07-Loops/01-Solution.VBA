Dim i As Integer

For i = 1 To 10
    If Cells(i, 1).Value > 100 Then
        Cells(i, 2).Value = "Yes"
    Else
        Cells(i, 2).Value = "No"
    End If
Next i

