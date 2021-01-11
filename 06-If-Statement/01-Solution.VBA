'We assume that age is in B1 cell

Dim age As Integer
age = Range("A1").Value

If age >= 18 Then
    Sheets.Add.Name = "Alcohol"
Else
    MsgBox "Go back home kiddo!"
End If

