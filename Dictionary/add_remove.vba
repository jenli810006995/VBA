Option Explicit

  Sub AddRemoveCount()
    Dim dict As New Scripting.Dictionary
    
    ' Add some items
    dict.Add "Orange", 55
    dict.Add "Peach", 55
    dict.Add "Plum", 55
    Debug.Print "The number of item is " & dict.Count
    
    ' Remove one item
    dict.Remove "Orange"
    Debug.Print "The number of item is " & dict.Count
    
    ' Remove all items
    dict.RemoveAll
    Debug.Print "The number of item is " & dict.Count
  
  End Sub
