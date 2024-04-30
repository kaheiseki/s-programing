Function CalculateCircle() As Double
  Randomize
  Dim x As Double
  Dim y As Double
  Dim s as Double
  Dim count As Integer
  For i = 1 To 10000 
    x = Rnd
    y = Rnd
    s = x * x + y * y
    If s < = 1 Then
      count = count + 1
    End If
  Next i
  CulculateCircle = count / 10000 * 4
End function