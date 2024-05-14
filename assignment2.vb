FUnction NormDistBoxMuller()
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

  NormDistBoxMuller = count / 200000
End Function

Function BoxMullerMarsaglia()
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

  BoxMullerMarsaglia = count1 / 200000
End Function

Sub Assignment2()
  Dim swap As Variant
  ReDim swap(1 To 20)
  Dim p As Variant
  ReDim p(1 To 20)
  Dim rate As Variant
  ReDim rate(1 To 20)
  swap(1) = 1 / 100
  For i = 2 To 20 
    swap(i) = (spline(i / 2) / 100)
  Next i

  p(1) = 1 / (1 + swap(1) * 0.5)  
  For i = 2 To 20
    sum = 0
    For j = 1 To i - 1
      sum = sum + p(j) * 0.5
    Next j
    p(i) = (1 - swap(i) * sum) / (1 + swap(i) * 0.5)    
  Next i

  For i = 1 To 20
    rate(i) = -Log(p(i)) / (i * 0.5)
  Next i

  For i = 1 To 20
    Cells(i, 1).Value = rate(i) * 100
  Next i
End Sub