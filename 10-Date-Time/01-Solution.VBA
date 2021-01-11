Dim i As Integer, datevar as Date

For i = 2 To 4
    datevar= DateValue(Cells(i, 1).Value)
    Cells(i, 2).Value = Year(datevar)
    Cells(i, 3).Value = Month(datevar)
    Cells(i, 4).Value = Day(datevar)
Next i

