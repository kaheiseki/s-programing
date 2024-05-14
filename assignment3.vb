Function EuroCallPrice(s0 As Integer, r As Double, sigma As Double, T As Double, K As Integer, m As Integer, count As Currency) As Double
  Dim Path As Variant
  ReDim Path(0 To m)
  Dim sum As Currency
  Dim random_value As Double
  sum = 0
  Path(0) = s0
  For i = 1 To count
    For j = 1 To m
      random_value = 0
      Do while random_value = 0
        random_value = Rnd()
      Loop
      Path(j) = Path(j - 1) + r * Path(j - 1) * (T / m) + sigma * Path(j - 1) * Sqr(T / m) * worksheetfunction.normsinv(random_value)
    Next j
    If Path(m) - K > 0 Then
      sum = sum + Path(m) - K
    End If
  Next i

  EuroCallPrice = sum / count * exp(-r * T)
End Function


Function DownAndOutCall(s0 As Integer, r As Double, sigma As Double, T As Double, K As Integer, B As Integer, m As Integer, count As Currency) As Double
  Dim Path As Variant
  ReDim Path(0 To m)
  Dim sum As Currency
  Dim exec_flag As Boolean
  Dim random_value As Double
  sum = 0
  Path(0) = s0
  For i = 1 To count
    exec_flag = True
    For j = 1 To m
      random_value = 0
      Do while random_value = 0
        random_value = Rnd()
      Loop
      Path(j) = Path(j - 1) * exp((r - (sigma * sigma) / 2) * (T / m) + sigma * Sqr(T / m) * worksheetfunction.normsinv(random_value))
      If Path(j) <= B Then
        exec_flag = False
      End If
    Next j
    If exec_flag And Path(m) - K > 0 Then
      sum = sum + Path(m) - K
    End If
  Next i

  DownAndOutCall = sum / count * exp(-r * T)
End Function

  Function AvgCallPrice(s0 As Integer, r As Double, sigma As Double, T As Double, K As Integer, m As Integer, count As Currency) As Double
    Dim Path As Variant
    ReDim Path(0 To m)
    Dim n As Currency
    Dim sum As Double
    Dim random_value As Double
    n = 0
    Path(0) = s0
    For i = 1 To count
      sum = 0
      For j = 1 To m
        random_value = 0
        Do while random_value = 0
          random_value = Rnd()
        Loop
        Path(j) = Path(j - 1) * exp((r - (sigma * sigma) / 2) * (T / m) + sigma * Sqr(T / m) * worksheetfunction.normsinv(random_value))
        sum = sum + Path(j)
      Next j
      If (sum / m) - K > 0 Then
        n = n + (sum / m) - K
      End If
    Next i
  
    AvgCallPrice = n / count * exp(-1 * r * T)
  End Function

Function DownAndOutCallAnalytical(s0 As Integer, r As Double, sigma As Double, T As Double, K As Integer, B As Integer) As Double
  Dim d1 As Double 
  Dim d2 As Double 
  Dim factor As Double
  Dim return_value As Double
  Dim Nd1 As Double
  Dim Nd2 As Double
  Dim Nd1SigmaSqrT As Double
  Dim Nd2SigmaSqrT As Double

  factor = (B / s0) ^ ((2 * r) / (sigma * sigma) - 1)
  d1 = (Log(s0 / K) + (r - ((sigma * sigma) / 2)) * T) / (sigma * Sqr(T))
  d2 = (Log((B * B) / (K * s0)) + (r - (sigma * sigma) / 2) * T) / (sigma * Sqr(T))
  Nd1 = worksheetfunction.normsdist(d1)
  Nd2 = worksheetfunction.normsdist(d2)
  Nd1SigmaSqrT = worksheetfunction.normsdist(d1 + sigma * Sqr(T))
  Nd2SigmaSqrT = worksheetfunction.normsdist(d2 + sigma * Sqr(T))

  return_value = s0 * Nd1SigmaSqrT - K * Exp(-r * T) * Nd1 - factor * (((B * B) / s0) * Nd2SigmaSqrT - K * Exp(-r * T) * Nd2)
  DownAndOutCallAnalytical = return_value
End Function