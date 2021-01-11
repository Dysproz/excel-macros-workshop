Dim mytable As Range
Set mytable = Range("B4:D11")

Worksheet("Objects").Range("D18:D25").Value = Worksheet("Objects").mytable.Columns(3).Value
Worksheet("Objects").Range("B29:D29").Value = Worksheet("Objects").mytable.Rows(8).Value

