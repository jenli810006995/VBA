Option Explicit

Enum ReadColumns
    rcInvoice = 1 ' column 1, incase we want to change value in the future, just change this
    rcCompany = rcInvoice + 1
    rcAmount = rcCompany + 1
End Enum

Enum WriteColumns
    wcCompany = 6
    wcAmount = wcCompany + 1
End Enum

Enum WriteRow
    wrHeader = 1
    wrStartRow = 2
End Enum


Public Sub CreateReport()

  ' sub should be small function
  
  Dim dict As Dictionary
  
  
  ' Read the data
  
  Set dict = Readdata
  
  Debug.Assert Not (dict Is Nothing)
  Debug.Assert dict.Count > 0
  
  
  Call PrintDictionary(dict, "After Read Data")
  
  ' Apply the discount
  
  Set dict = ApplyDiscount(dict)
  
  Debug.Assert Not (dict Is Nothing)
  Debug.Assert dict.Count > 0
  
  Call PrintDictionary(dict, "After Apply Discount")
  
  ' Write the data
  
  Call WriteData(dict)
  
End Sub

  ' Read the data
  ' Create a function, bc function returns value
  
Private Function Readdata() As Dictionary

  Dim dict As New Dictionary
  
  ' read data
  
   Dim rg As Range
  ' change range to an array
  ' for large data, reading from arr is much faster
  
  ' Dim arr As Variant
  
  Set rg = ShData.ListObjects("tbCompany").DataBodyRange
'  Set arr = ShData.ListObjects("tbCompany").DataBodyRange.Value
  
  ' ListObjects is a table
  
  Dim company As String, amount As Long
  
  Dim i As Long
  For i = 1 To rg.Rows.Count
  'For i = LBound(arr, 1) To UBound(arr, 1)
      company = rg.Cells(i, rcCompany).Value
      amount = rg.Cells(i, rcAmount).Value
  ' use 2 to skip header
  ' After use ListObjects we can use i = 1
'      If dict.Exists = False Then
'          dict.Add company, amount
'      End If
      ' the above if end if is optional, bc if dict not exist, it would auto add it
      
      dict(company) = dict(company) + amount
  
  Next i
  
  ' shdata is the call name. If users change worksheet name it is still work
  
  ' return dictionary
  
  Set Readdata = dict
  
  ' Debug.Print "Readdata()"
  
End Function
  
  ' Apply the discount
  
Private Function ApplyDiscount(ByVal dict As Dictionary) As Dictionary

    Dim key As Variant, amount As Long
    For Each key In dict
        amount = dict(key)
        If amount > 15000 Then
            dict(key) = CLng(amount - (amount * 0.1)) ' round to a long integer
        End If
  
  Next key
  
  ' Apply discount
  
  Set ApplyDiscount = dict
  
  ' Debug.Print to see if the code is runing OK
  ' Debug.Print "ApplyDiscount()"
  
End Function
  
  ' Write the data
  
Private Sub WriteData(ByVal dict As Dictionary)
  
  'Debug.Print "WriteData()"
  ' clear the data
  
    ShData.Cells(wrStartRow, wcAmount).CurrentRegion.Offset(1).ClearContents
    ' offset(1) to avoid the header
  
    Dim row As Long
    row = wrStartRow
    
    Dim key As Variant, amount As Long
    
    For Each key In dict
        ShData.Cells(row, wcCompany).Value = key
        ShData.Cells(row, wcAmount).Value = dict(key)
        row = row + 1 ' it would write the value to the next row
  
  Next key
  
End Sub


' Reference: https://youtu.be/sN8kEbGlxUs
