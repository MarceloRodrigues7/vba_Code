''Form Basic
Private Sub Workbook_Open() ''Workbook Initialization
    Application.Visible = False
    Planilha1.Visible = False
    Login.Show
End Sub

Private Sub btnAcess_Click()  ''Login Button Code
''Check Acc e Pass para Aceesar
  If Me.txtAcc.Value = "user2" And Me.txtPass.Value = "a1a2a3" Then
        MsgBox "Welcome " & txtAcc.Text, vbOKOnly, "Sucess"
    Application.Visible = True
    Sheets("Home").Select
    Range("cellAcc").Value = txtAcc.Text
    Unload Login
  Exit Sub
  End If
''Add New Account    
''Check Acc e Pass para Aceesar
  If Me.txtAcc.Value = "NewUser2" And Me.txtPass.Value = "NewPass2" Then
        MsgBox "Welcome " & txtAcc.Text, vbOKOnly, "Sucess"
    Application.Visible = True
    Sheets("Home").Select
    Range("cellAcc").Value = txtAcc.Text
    Unload Login
  Exit Sub
    ''Fail Acess(No Login)
  Else
        MsgBox "Access denied", vbCritical, "Fail"
        Application.Quit
  End If
Enb Sub
