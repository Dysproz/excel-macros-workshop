Dim i As Integer

For i = 1 To 10
    If i < 3 Then
        Cells(i, 4).Value = 1
    Else
        Cells(i, 4).Value = Cells(i-1, 4).Value + Cells(i-2, 4).Value
    End If
Next i
