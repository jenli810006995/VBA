Option Explicit

' Create a sub called Top5Report to write the data in all the columns from the top 5 countries to the Top 5 section in the Report worksheet. This is the range starting at B3 on the Report worksheet. Use the code name to refers to the worksheets.

Sub Top5Report()

  shReport.Range("B3:E7") = shCountries.Range("B1:E6").Value

End Sub

' Create a sub call AreaReport to write all the areas size to the All the Areas section in the Report worksheet. This is the range H3:H30. Use the worksheet name to refer to the worksheets.

Sub AreaReport()

  shReport.Range("H3:H30") = shCountries.Range("D2:D29").Value
  
End Sub

' Create a sub called ImmediateReport as follows, read the area and population from Russia to two variables. Print the population per square kilometre(pop/area) to the Immediate Window.

Sub ImmediateReport()

Dim RussiaPop As Long
Dim RussiaArea As Long

RussiaPop = shCountries.Range("D2")
RussiaArea = shCountries.Range("E2")

Debug.Print "Population Per Square Kilometer for Russia is " & RussiaPop / RussiaArea

End Sub

' Create a new worksheet and call it areas. Set the code name to be shAreas. Create a sub called RowsToCols that reads all the areas in D2:D11 from Countries worksheet and writes them to the range A1:J1 in the new worksheet Areas.

Sub RowsToCols()

shAreas.Range("A1:J1").Value = WorksheetFunction.Transpose(shCountries.Range("D2:D11").Value)

End Sub

