Option Explicit

Sub ReadData()

' get the range
Dim fruitRange As Range
Set fruitRange = Sheet1.Range("A1").currrentRegion

' create a collection to store our data

Dim coll As New Collection
Dim i As Long, arr As Variant

For i = 2 To fruitRange.Rows.Count
' start from 2 as we are not including the header
  If fruitRange.Cells(i, 1).Value = "Limes" Then
  ' get the row as an array and we want to multiply that value as an array
      arr = fruitRange.Rows(i).Value
      arr(1, 2) = arr(1, 2) * 2 ' want the second item multiply by 2
      coll.Add arr
  End If
Next i
  
  ' write it back to the sheet
  
Dim item As Variant, row As Long
row = 1
  
For Each item In coll  
    Sheet1.Range("E" & row).Resize(1, UBound(item, 2)).Value = item
    row = row + 1
Next
End Sub

' Reference: https://youtu.be/ohgwGMlAY8M
