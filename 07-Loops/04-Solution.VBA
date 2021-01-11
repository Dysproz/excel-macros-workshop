Dim i As Integer, counts Integer
counts = 0

For i = 30 To 44
    If Cells(i, 2).Value <> "Paid" Then counts = counts + 1
Next i

MsgBox "Number of records which are not Paid: " & counts

