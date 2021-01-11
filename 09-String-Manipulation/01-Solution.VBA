Dim i As Integer, name As String, spaceidx As Integer

For i = 2 To 5
    name = Cells(i, 2).Value
    Cells(i, 2).Value = Left(name, 3) & "." & Right(name, 4) & "@evilcorp.com"
    spaceidx = Instr(name, " ")
    Cells(i, 3).Value = Left(name, spaceidx)
    Cells(i, 4).Value = Right(name, Len(name)-spaceidx)
Next i
