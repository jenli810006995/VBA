Option Explicit

Sub arrays()
  dim mark(1 to 2) as long
  dim i as long
  
  for i = 1 to 2
    mark(i) = ShNumbers.Range("A" & i)
  next i
  
  for i = 1 to 2
    Debug.Print mark(i)
  next i

End Sub
