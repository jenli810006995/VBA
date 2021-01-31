' if we use a codeName, such as shData, we do not have to define the data type in the beginning, such as dim sheet1 as worksheet, and set

Option Explicit
Sub TransposeData()

  ' Dim sheet1 As Worksheet
  ' Set sheet1 = ThisWorkbook.Worksheets("sheet1")
  
  shData.Range("E1:H1") = WorksheetFunction.Transpose(shData.Range("A1:A4").Value)
  shData.Range("L1:L4") = WorksheetFunction.Transpose(shData.Range("E1:H1").Value)

End Sub

