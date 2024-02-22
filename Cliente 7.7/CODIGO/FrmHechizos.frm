VERSION 5.00
Begin VB.Form FrmHechizos 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   3585
   ClientLeft      =   8925
   ClientTop       =   2040
   ClientWidth     =   4590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleMode       =   0  'User
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FrmHechizos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HechiClk As Integer
Private Sub Form_Load()
Call MakeWindowTransparent(FrmHechizos.hWnd, 100)
End Sub

Private Sub hlst_DblClick()
'AKI2
    If (hlst.ListIndex = -1) Then
        MsgBox ("Debes Seleccionar un hechizo")
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "Selecciona un hueco", 255, 255, 255, True, False, False)
        FrmHechizos.hlst.MousePointer = 2
        HechiClk = FrmHechizos.hlst.ListIndex + 1
    End If
End Sub
Private Sub hlst_Click()
'AKI2
    If (HechiClk <> 0 And FrmHechizos.hlst.ListIndex <> -1) Then
        FrmHechizos.hlst.MousePointer = vbCustom
        FrmHechizos.hlst.MouseIcon = LoadPicture(App.Path & "\graficos\diablo.ico")

        'frmMain.MousePointer = vbNormal
        SendData "CZ" & HechiClk & "," & (FrmHechizos.hlst.ListIndex + 1)
        HechiClk = 0
    End If
  
End Sub
Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       
    'pluto:2.5.0
'If SendTxt.Visible = True And KeyCode <> 13 Then Exit Sub

'If (Not SendTxt.Visible) And _
   ((KeyCode >= 65 And KeyCode <= 90) Or _
   (KeyCode >= 48 And KeyCode <= 57)) Then
   
Select Case KeyCode
             'pluto:2.3
                Case vbKeyQ:

                    If (frmEquipo.Visible = True) Then
                       Unload frmEquipo
                    Else
                        frmEquipo.Visible = True
                    End If
                
            
                
                Case vbKeyH:
                    '[MerLiNz:MAPA]
                    If (frmMap.Visible = True) Then
                        Unload frmMap
                    Else
                        frmMap.Visible = True
                    End If
                    '[\END]
                'pluto:2-3-04
                Case vbKeyX:
                

               ' Case vbKeyM:
                    'If Not IsPlayingCheck Then
                       ' Musica = 0
                       ' Play_Midi
                    'Else
                       ' Musica = 1
                        'Stop_Midi
                    'End If
                Case vbKeyA:
    SendData "AG" & ShTime
    bInvMod = True
               ' Case vbKeyC:
                   ' Call SendData("TAB")
                    'If IScombate = True Then
                    'frmMain.Flash2.Visible = True
                    'frmMain.Flash2.FrameNum = -1
                    'frmMain.Flash2.Play
                    'Else
                    
                    'frmMain.Flash2.Visible = False
                    
                    'End If
                    'IScombate = Not IScombate
 'FrmHechizos.SetFocus
                Case vbKeyE:
    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then _
    'pluto:2.4.5
    SendData "EQUI" & ItemElegido & ",O," & ShTime
    bInvMod = True
End If
                Case vbKeyN:
                    Nombres = Not Nombres
                Case vbKeyD
                    If Not NoPuedeMagia Then Call SendData("UK" & Domar)
                Case vbKeyR:
                    If Not NoPuedeMagia Then Call SendData("UK" & Robar)
                'Case vbKeyS:
                   'Call SendData("SEG")
                Case vbKeyZ:
                    Call SendData("ACT")
                Case vbKeyO:
                    Call SendData("UK" & Ocultarse)
                'pluto:2-3-04
                Case vbKeyP:
                SendData "/DRAGPUNTOS"
                Case vbKeyT:
                     If (ItemElegido > 0 And ItemElegido < MAX_INVENTORY_SLOTS + 1) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 And seguroobjetos = False Then
        SendData "TI" & ItemElegido & "," & 1
        Else
           'If UserInventory(ItemElegido).Amount > 1 Then
            frmCantidad.Show
           End If
        'End If
    End If

    bInvMod = True
                'pluto:2.3
                Case vbKeyK:
                    SendData "XX" & ItemElegido
 bInvMod = True
                Case vbKeyU:
              If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                       If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & ItemElegido
    bInvMod = True
                    End If
            'PLUTO:2.8.0
            Case vbKeyV:
                SendData "/VAMPIRO"
            
            End Select
      '  End If


        
        
        Select Case KeyCode
            Case vbKeyReturn:
                    hechi = 1
               If Not frmCantidad.Visible Then 'And SendTxt.Visible = False Then
                    frmMain.SendTxt.Visible = True
                    FrmHechizos.Visible = False
                    frmMain.picInv.Visible = False
                    frmMain.DespInv(0).Visible = False
                     frmMain.DespInv(1).Visible = False
                    frmMain.SendTxt.SetFocus
                
                   
           End If
                    
            
            
            Case vbKeyF1:
            SendData "/resucitar"
            Case vbKeyF2:
            SendData "/meditar"
            Case vbKeyF3:
                SendData "/comerciar"
             Case vbKeyF4:
                FPSFLAG = Not FPSFLAG
            Case vbKeyF5:
                SendData "/online"
            Case vbKeyF6:
                SendData "/onlineclan"
            'PLUTO:2.4.7
            Case vbKeyF7:
            SendData "/angel"
            Case vbKeyF8:
           SendData "/demonio"
            Case vbKeyF9:
            SendData "/torneo"
            Case vbKeyF12:
            SendData "/salir"
            
    'para gms --------------------------------------
            'Case vbKeyF1:
               ' SendData "/dest"
            'Case vbKeyF8:
                'SendData "/mata"
            'Case vbKeyF3
             '   SendData "/teleploc"
           
            'Case vbKeyF5
             '   SendData "/invisible"
           'Case vbKeyF6
            '    SendData "/show sos"
            'Case vbKeyF7:
             '   SendData "/online"
      '------------------------------------------
      Case vbKeyControl:
                If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "AT"
                        UserCanAttack = 0
                End If

Case vbKeySpace:
If CurMap <> 192 Then Exit Sub



Dim aa As Byte
Dim a As Integer, b As Integer

If CharList(UserCharIndex).Heading = 1 And MapData(UserPos.x, UserPos.Y - 1).CharIndex > 0 Then
If CharList(MapData(UserPos.x, UserPos.Y - 1).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList(MapData(UserPos.x, UserPos.Y - 1).CharIndex).Body.Walk(1).GrhIndex < 4525 Then a = 0: b = -1
End If

If CharList(UserCharIndex).Heading = 2 And MapData(UserPos.x + 1, UserPos.Y).CharIndex > 0 Then
If CharList(MapData(UserPos.x + 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList(MapData(UserPos.x + 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex < 4525 Then a = 1: b = 0
End If

If CharList(UserCharIndex).Heading = 3 And MapData(UserPos.x, UserPos.Y + 1).CharIndex > 0 Then
If CharList(MapData(UserPos.x, UserPos.Y + 1).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList(MapData(UserPos.x, UserPos.Y + 1).CharIndex).Body.Walk(1).GrhIndex < 4525 Then a = 0: b = 1
End If

If CharList(UserCharIndex).Heading = 4 And MapData(UserPos.x - 1, UserPos.Y).CharIndex > 0 Then
If CharList(MapData(UserPos.x - 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList(MapData(UserPos.x - 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex < 4525 Then a = -1: b = 0
End If

If a = 0 And b = 0 Then Exit Sub
Call PlayWaveDS("145.wav")

For aa = 1 To 12
SendData ("BOLL" & CharList(UserCharIndex).Heading & "," & MapData(UserPos.x + a, UserPos.Y + b).CharIndex)
Next
'pluto:2.8.0


End Select
' FrmHechizos.Visible = False
        'If KeyCode = vbKeyReturn Then
        'frmMain.SendTxt.Visible = True
        'frmMain.SendTxt.SetFocus
        'hechi = 1
       ' Else
        'frmMain.picInv.Visible = True
       ' frmMain.picInv.SetFocus
       ' hechi = 0
        'End If
       KeyCode = 0
       
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
       
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
       
End Sub
