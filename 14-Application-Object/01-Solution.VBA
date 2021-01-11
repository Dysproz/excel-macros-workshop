Function MySum(numbers As Range) As Integer
MySum = Application.WorksheetFunction.Sum(numbers)
End Function

Function MyAvg(numbers As Range) As Double
MyAvg = Application.WorksheetFunction.Average(numbers)
End Function

Function MyMax() As Integer
MyMax = "=MAX(A1:A10)"
End Function

Function MyMin() As Integer
MyMin = "=MIN(A1:A10)"
End Function

Cells(2, 4).Value = MySum(Range("A1:A10"))
Cells(3, 4).Value = MyAvg(Range("A1:A10:))
Range("D4").Value = MyMax()
Range("D5").Value = MyMin()

