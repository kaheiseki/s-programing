

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
  ' 結果をシートに出力
  For i = 1 To 5
    ' 各wから総和を割り、ウェイトを出力する
    Cells(i, 8).Value = result(i, 1)  / sum
  Next i

  ' リターンの出力
  exp_return = (result(1,1) * 0.05 / sum) + (result(2,1) * 0.06 / sum) + (result(3,1) * 0.03 / sum) + (result(4,1) * 0.04 / sum)
  Cells(1, 10).Value = exp_return 

  ' リスクの出力
  risk = (result(1,1)/sum) * (result(1,1)/sum) * 0.04 + (result(2,1)/sum) * (result(2,1)/sum) * 0.09 + (result(3,1)/sum) * (result(3,1)/sum) * 0.01 + (result(4,1)/sum) * (result(4,1)/sum) * 0.04 + (result(1,1)/sum) * (result(2,1)/sum) * 0.072 + (result(1,1)/sum) * (result(3,1)/sum) * 0.012 + (result(1,1)/sum) * (result(4,1)/sum) * 0.032 + (result(2,1)/sum) * (result(3,1)/sum) * 0.03 + (result(2,1)/sum) * (result(4,1)/sum) * 0.024 + (result(3,1)/sum) * (result(4,1)/sum) * 0.02
  Cells(2, 10).Value = risk
End Sub