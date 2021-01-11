Dim Films(1 To 3, 1 To 4) As String
Dim i As Integer, j As Integer

For i = 2 To 4
   For j = 1 To 4
        Films(i-1, j) = Cells(i, j).Value
   Next j
Next i

For i = 1 To 3
    Cells(i+1, 7).Value = Films(i, 3)
    Cells(i+1, 8).Value = Films(i, 2)
    Cells(i+1, 9).Value = Films(i, 4)
    Cells(i+1, 10).Value = Films(i, 1)
Next i
