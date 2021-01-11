Dim mytable As Range
Set mytable = Range("B4:D11")

Dim targettable As Range
Set targettable = Range("H4:I11")

Worksheet("Objects").Activate
Worksheet("Objects").mytable.Select
Selection.Copy
Worksheet("Objects").targettable.Select
ActiveSheet.Paste

'Or simply Worksheet("Objects").targettable.Value = mytable.Value

