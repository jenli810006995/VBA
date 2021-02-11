Option Explicit

Sub TwoDim()
  Dim marks(1 to 2, 1 to 5) as long
  
    marks(1, 1) = 67
    marks(2, 5) = 75
  
  Dim i, j as long
  
  for i = LBound(marks, 1) to UBound(marks, 1)
    for j = LBound(marks, 1) to UBound(marks, 1)
      Debug.Print i, j, marks(1, j)
    next j
  next i
  
End Sub
  
