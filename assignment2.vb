
Sub NormDistBoxMuller()
  Dim count As Integer
  Dim x1 As Double
  Dim x2 As Double
  count = 0
  For i = 1 To 100000
    Dim rand As Variant
    rand = Box_Muller()
    x1 = rand(0)
    x2 = rand(1)

    If x1 > 1 Then
      count = count + 1
    End If

    If x2 > 1 Then
      count = count + 1
    End If
  Next i

  Cells(25, 25).Value = count / 200000
End Sub

Sub BoxMullerMarsaglia()
  Dim count1 As Integer
  Dim x1 As Double
  Dim x2 As Double
  count1 = 0
  For i = 1 To 100000
    Dim rand As Variant
    rand = Box_Muller_Marsaglia()
    x1 = rand(0)
    x2 = rand(1)

    If x1 > 1 Then
      count1 = count1 + 1
    End If

    If x2 > 1 Then
      count1 = count1 + 1
    End If
  Next i

  Cells(26, 25).Value = count1 / 200000
End Sub

Sub Assignment2()
  Dim func_return As Variant

  ' シートから行列を取得
  For i = 2 To 20 
    Cells(i, 1).Value = spline(i / 2)
  Next i
End Sub