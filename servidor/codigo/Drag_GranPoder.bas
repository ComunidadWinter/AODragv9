Attribute VB_Name = "Drag_GranPoder"
Option Explicit

Public GranPoder As Integer


Sub OtorgarGranPoder(UserIndex As Integer)
  Dim LoopC As Integer
  Dim EncontroIdeal As Boolean
  
  If LastUser = 0 Then Exit Sub
  
  If UserIndex = 0 Then
    GranPoder = 0
    Do While EncontroIdeal = False And LoopC < 500
      LoopC = LoopC + 1
      UserIndex = RandomNumber(1, LastUser)
      
      If UserList(UserIndex).flags.Privilegios = PlayerType.User And _
      UserList(UserIndex).flags.UserLogged = True And UserList(UserIndex).flags.Muerto = 0 And EsNewbie(UserIndex) = False Then
      
      EncontroIdeal = True
      Exit Do
      
    End If
  Loop
  If Not EncontroIdeal Then
    UserIndex = 0
    GranPoder = 0
  End If
End If

'Si hay menos de 50 usuarios, no hay gran poder
If NumUsers < 10 Then Exit Sub

If UserIndex > 0 Then
  If UserList(UserIndex).flags.Muerto = 1 Then
    Call OtorgarGranPoder(0)
  Else
    GranPoder = UserIndex
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Los Dioses le otorgan el Gran Poder a " & UserList(UserIndex).Name & " en el Mapa " & UserList(UserIndex).Pos.Map, FontTypeNames.FONTTYPE_DIOS))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FxGranPoder, 0))
    frmMain.QuienGP.Caption = UserList(UserIndex).Name
  End If
End If
End Sub




