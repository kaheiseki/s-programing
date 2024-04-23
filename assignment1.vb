Sub Assignment1()
  ' 変数宣言
  Dim aMatrix As Variant
  Dim bMatrix As Variant
  Dim sum As Double
  Dim exp_return As Double
  Dim risk As Double
  risk = 0

  ' シートから行列を取得
  aMatrix = GetMatrix(1, 5, 1, 5)
  bMatrix = GetMatrix(1, 5, 6, 6)
  covMatrix = GetMatrix(8, 11, 1, 4)
  stdMatrix = GetMatrix(15, 15, 1, 4)

  ' 方程式を計算
  result = Equation(aMatrix, bMatrix, 5)

  ' wの総和を求める
  sum = result(1, 1) + result(2, 1) + result(3, 1) + result(4, 1)

  ' wの総和で割ることでxを求める
  For i = 1 To 4 
    result(i, 1) = result(i, 1) / sum
  Next i
  ' 結果をシートに出力
  For i = 1 To 5
    Cells(i, 8).Value = result(i, 1)
  Next i

  ' リターンの出力
  exp_return = (result(1, 1) * 0.05) + (result(2, 1) * 0.06) + (result(3, 1) * 0.03) + (result(4, 1) * 0.04)
  Cells(1, 10).Value = exp_return 

  ' リスクの出力
  For i = 1 To 4
    For j = 1 To 4
      risk = risk + result(i,1) * result(j,1) * covMatrix(i, j) * stdMatrix(1, i) * stdMatrix(1, j)
    Next j
  Next i
  Cells(2, 10).Value = risk
End Sub