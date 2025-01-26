Function calculateSum(arr)
  Dim sum, i
  sum = 0
  For i = 0 To UBound(arr)
    sum = sum + arr(i)
  Next
  calculateSum = sum
End Function