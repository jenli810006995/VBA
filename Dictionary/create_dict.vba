Option Explicit

  Sub CheckFruit()
  
  ' Select Tools-> References from the Visual Basic menu
  ' Check box inside "Microsoft Scripting Runtime" in the list
    Dim dict As New Scripting.Dictionary
  ' the above is called Early Binding
    
  ' Add to fruit to Dictionary
  
    dict.Add Key:="Apple", Item:=51
    dict.Add Key:="Peach", Item:=34
    dict.Add Key:="Plum", Item:=43
  
    Dim sFruit As String
    
    ' ask user to enter fruit
    sFruit = InputBox("Please enter the name of a fruit")
    
    If dict.Exists("sFruit") Then
      MsgBox sFruit & " exists and has value " & dict(sFruit)
    Else
      MsgBox sFruit & " does not exist. "
      
    End If
    
    Set dict = Nothing
  
  End Sub
