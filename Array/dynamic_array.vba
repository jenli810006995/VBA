Option Explicit

Sub dynamic()
  Dim mark() as Long
  Dim lastrow as Long
  
  lastrow = ShNumbers.range("A" & ShNumbers.rows.count).End(xlUp).Row
  
  ' dynamic array
  Redim mark(1 to lastrow)
  
  Dim i as long
  for i = 1 to lastrow
    mark(i) = ShNumbers.range("A" & i)
  next i
  
  for i = LBound(mark) to UBound(mark)
    Debug.Print mark(i)
  next i

End Sub
