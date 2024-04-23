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