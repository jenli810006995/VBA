Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
' this sub works when any change happens in this worksheet
  
  ' if more than one cell changes, we dont have to do anything
  If Target.Cells.Count > 1 Then Exit Sub
  
  ' if the criteria is only one row, meaning there is no filter
  
  If shData.Range("A1").CurrentRegion.Rows.Count = 1 Then
  ' clear the filter
    clearfilter

  ElseIf Not Application.Intersect(Target, shData.Range("A1").CurrentRegion) Is Nothing Then
  ' this means if a target cell within our criteria, then we want to call the advanced filter
      UseAdvancedFilterInPlace
  End If
  
  ' this checks if two changes interset it returns a row or a range

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
