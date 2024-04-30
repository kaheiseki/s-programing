' n元連立一次方程式を解く関数
Function Equation(a_matrix As Variant, b_matrix As Variant,  a_row As Integer) As Variant
  ' 変数の宣言
  Dim tmp_a_matrix As Variant
  Dim tmp_b_matrix As Variant
  Dim return_matrix() As Variant
  ReDim return_matrix(1 To a_row, 1)

  ' 一時保管用の行列に引数を代入
  tmp_a_matrix = a_matrix
  tmp_b_matrix = b_matrix
  ' 行列処理
  For i = 1 To a_row
    For j = i + 1 To a_row 
      For k = 1 To a_row
        tmp_a_matrix(j, k) = a_matrix(j, k) - a_matrix(j, i) * a_matrix(i, k) / a_matrix(i, i)
      Next k
      tmp_b_matrix(j, 1) = b_matrix(j, 1) - ((a_matrix(j, i) * b_matrix(i, 1)) / a_matrix(i, i))
    Next j
    a_matrix = tmp_a_matrix
    b_matrix = tmp_b_matrix
  Next i

  ' 方程式の各変数を計算する処理
  For i = a_row To 1 Step -1
    sum = 0
    If i < a_row Then
      For j = i + 1 To a_row
        sum = sum + a_matrix(i, j) * return_matrix(j, 1)
      Next j
    End If
    return_matrix(i, 1) = (b_matrix(i, 1) - sum) / a_matrix(i, i)
  Next i
  Equation = return_matrix
End Function

  ' エクセルのシートから行列を取得する関数
  ' 行の始まり、行の終わり、列の始まり、列の終わりを引数に取る
Function GetMatrix(row_start As Integer, row_end As Integer, column_start As Integer, column_end As Integer) As Variant
  Dim returnArray() As Variant
  ReDim returnArray(1 To row_end - row_start + 1, 1 To column_end - column_start + 1)
  For i = row_start To row_end
    For j = column_start To column_end
      returnArray(i - row_start + 1, j - column_start + 1) = Cells(i, j).Value
    Next j
  Next i
  GetMatrix = returnArray
End Function

Function Spline(time As Double) As Variant
  Dim x As Variant
  Redim x(0 To 6)
  Dim y As Variant
  Redim y(0 To 6)
  Dim delta As Variant
  Redim delta(0 To 5)
  Dim b As Variant
  Redim b(0 To 4)
  Dim a As Variant
  Redim a(0 To 4, 0 To 4)
  Dim result As Variant
  Redim result(1 To 5)
  Dim adash As Variant
  Redim adash(0 To 5)
  Dim bdash As Variant
  Redim bdash(0 To 5)
  Dim cdash As Variant
  Redim cdash(0 To 5)
  Dim ddash As Variant
  Redim ddash(0 To 5)
  x(0) = 0.5
  x(1) = 1
  x(2) = 2
  x(3) = 3
  x(4) = 5
  x(5) = 7
  x(6) = 10

  y(0) = 1
  y(1) = 1.1
  y(2) = 1.2
  y(3) = 1.3
  y(4) = 1.4
  y(5) = 1.5
  y(6) = 1.6
  For i = 0 To 5
    delta(i) = x(i + 1) - x(i)
  Next i
  For i = 0 To 4
    b(i) = 6 * ((y(i + 2) - y(i + 1)) / delta(i + 1) - (y(i + 1) - y(i)) / delta(i))
  Next i
  For i = 0 To 4
    For j = 0 To 4
      If i = j Then
        a(i, j) = 2 * (delta(i) + delta(i + 1))
      ElseIf j = i - 1 Then
        a(i, j) = delta(i)
      ElseIf j = i + 1 Then
        a(i, j) = delta(i + 1)
      Else
        a(i, j) = 0
      End If
    Next j
  Next i

  For i = 1 To 4
    d = a(i, i - 1) / a(i - 1, i - 1)
    a(i, i - 1) = a(i, i - 1) - d * a(i - 1, i - 1)
    a(i, i) = a(i, i) - d * a(i - 1, i)
    b(i) = b(i) - d * b(i - 1)
  Next i

  result(5) = b(4) / a(4, 4)
  For i = 4 To 1 Step -1
    result(i) = (b(i - 1) - a(i - 1, i) * result(i + 1)) / a(i - 1, i - 1)
  Next i
  
  For i = 0 To 5
    ddash(i) = y(i)
  Next i

  For i = 1 To 5
    bdash(i) = result(i) / 2
  Next i

  For i = 1 To 4
    adash(i) = (result(i + 1) - 2 * bdash(i)) / (6 * delta(i))
  Next i

  For i = 1 To 4
    cdash(i) = (y(i + 1) - y(i)) / delta(i) - (2 * result(i) + result(i + 1)) * delta(i) / 6
  Next i

  cdash(5) = 3 * adash(4) * delta(4) ^ 2 + 2 * bdash(4) * delta(4) + cdash(4)
  adash(5) = (y(6) - bdash(5) * delta(5) ^ 2 - cdash(5) * delta(5) - ddash(5)) / (delta(5) ^ 3) 
  If x(1) <= time And time < x(2) Then
    Spline = adash(1) * (time - x(1)) ^ 3 + bdash(1) * (time - x(1)) ^ 2 + cdash(1) * (time - x(1)) + ddash(1)
  ElseIf x(2) <= time And time < x(3) Then
    Spline = adash(2) * (time - x(2)) ^ 3 + bdash(2) * (time - x(2)) ^ 2 + cdash(2) * (time - x(2)) + ddash(2)
  ElseIf x(3) <= time And time < x(4) Then
    Spline = adash(3) * (time - x(3)) ^ 3 + bdash(3) * (time - x(3)) ^ 2 + cdash(3) * (time - x(3)) + ddash(3)
  ElseIf x(4) <= time And time < x(5) Then
    Spline = adash(4) * (time - x(4)) ^ 3 + bdash(4) * (time - x(4)) ^ 2 + cdash(4) * (time - x(4)) + ddash(4)
  ElseIf x(5) <= time And time <= x(6) Then
    Spline = adash(5) * (time - x(5)) ^ 3 + bdash(5) * (time - x(5)) ^ 2 + cdash(5) * (time - x(5)) + ddash(5)
  End If
End Function

Function Box_Muller() As Variant
  Dim x1 As Double
  Dim x2 As Double
  Dim y1 As Double
  Dim y2 As Double

  x1 = Rnd()
  x2 = Rnd()
  y1 = Sqr(-2 * Log(x1)) * Cos(2 * Atn(1) * 4 * x2)
  y2 = Sqr(-2 * Log(x2)) * Sin(2 * Atn(1) * 4 * x1)
  Box_Muller = Array(y1, y2)
End Function

Function Box_Muller_Marsaglia() As Variant
  Dim x1 As Double
  Dim x2 As Double
  Dim y1 As Double
  Dim y2 As Double
  Dim r1 As Double
  Dim r2 As Double
  Dim s As Double

  x1 = Rnd()
  x2 = Rnd()
  r1 = 2 * x1 - 1
  r2 = 2 * x2 - 1
  s = r1 * r1 + r2 * r2
  While s >= 1
    x1 = Rnd()
    x2 = Rnd()
    r1 = 2 * x1 - 1
    r2 = 2 * x2 - 1
    s = r1 * r1 + r2 * r2
  Wend
  y1 = r1 * Sqr(-2 * Log(s) / s)
  y2 = r2 * Sqr(-2 * Log(s) / s)
  Box_Muller_Marsaglia = Array(y1, y2)
End Function