Dim i As Integer, datevar as Date

For i = 2 To 4
    datevar = DateValue(Cells(i, 1).Value)
    Cells(i, 5).Value = DateAdd("m", -9, datevar)
Next i
