Dim i As Integer
i = 1

Do While Cells(i, 8).Value <= 10
    i = i + 1
Loop
Cells(i, 8).Value = "<- This is it"

