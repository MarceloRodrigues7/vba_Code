Sub findReturn()
  Sheets("Plan2").Select
  Collumns("D:D").Select
  Set C = Cells.Find(What:="Ghost", After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, _
  SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not (C Is Nothing) Then
    Sheets("Plan1").Select
    Range("A2").Value = "True"
    Else
    Sheets("Plan1").Select
    Range("A2").Value = "False"
    End If
    
End Sub
