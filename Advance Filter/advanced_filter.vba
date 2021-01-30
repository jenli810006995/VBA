Option Explicit

Sub UseAdvancedFilterCopy()

  Dim rg As Range

  ' clear all the data but keep the header
  
  Set rg = ThisWorkbook.Worksheets("Output").Range("A1").CurrentRegion
  
  ' we dont want to delete everthing, we want to keep the header, so we can change the header
  ' and output the selected columns in the output
  
  rg.Offset(1).ClearContents
  
  Dim rgData As Range, rgCriteria As Range, rgOutput As Range
  
  Set rgData = ThisWorkbook.Worksheets("Data").Range("A5").CurrentRegion ' this returns back all current data
  
  Set rgCriteria = ThisWorkbook.Worksheets("Data").Range("A1").CurrentRegion
  
  Set rgOutput = ThisWorkbook.Worksheets("Output").Range("A1").CurrentRegion
  ' currentregion gets us back all the data we need
  
  rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgOutput

End Sub

Sub UseAdvancedFilterInPlace()
  
  Dim rgData As Range, rgCriteria As Range
  
  Set rgData = ThisWorkbook.Worksheets("Data").Range("A5").CurrentRegion ' this returns back all current data
  Set rgCriteria = ThisWorkbook.Worksheets("Data").Range("A1").CurrentRegion
  
  rgData.AdvancedFilter xlFilterCopy, rgCriteria

End Sub


Sub clearfilter()

  If ThisWorkbook.Worksheets("Data").FilterMode = True Then
      ThisWorkbook.Worksheets("Data").ShowAllData
      
  End If

End Sub

' Reference: https://youtu.be/0YNhxVu2a5s

