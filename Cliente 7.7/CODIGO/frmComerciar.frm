VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   1905
   ClientTop       =   0
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComerciar.frx":0000
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Text            =   "1"
      Top             =   6030
      Width           =   600
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   750
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   840
      Width           =   555
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4230
      Index           =   1
      Left            =   3615
      MouseIcon       =   "frmComerciar.frx":285A1
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1560
      Width           =   3090
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4230
      Index           =   0
      Left            =   240
      MouseIcon       =   "frmComerciar.frx":2926B
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1560
      Width           =   3045
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Botón derecho sobre el objeto para más infomación."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2400
      MouseIcon       =   "frmComerciar.frx":29F35
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   3030
      TabIndex        =   8
      Top             =   5760
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   1
      Left            =   3855
      MouseIcon       =   "frmComerciar.frx":2ABFF
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6000
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   0
      Left            =   615
      MouseIcon       =   "frmComerciar.frx":2B8C9
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6000
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   210
      Index           =   3
      Left            =   3990
      TabIndex        =   6
      Top             =   1215
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Index           =   4
      Left            =   3960
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Index           =   2
      Left            =   2400
      TabIndex        =   4
      Top             =   1170
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   900
      Width           =   120
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Public LastIndex1 As Integer
Public LastIndex2 As Integer


Private Sub cantidad_Change()
If Val(cantidad.Text) < 0 Then
    cantidad.Text = 1
End If

If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
    cantidad.Text = 1
End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub





Private Sub Form_Deactivate()
Me.SetFocus
End Sub


Private Sub Form_Load()
'La transparencia es graduable modificando el alphaamount en este caso esta en 150 mientras menor es este valor mas transparente se torna.




'Cargamos la interfase
Me.Picture = LoadPicture(App.Path & "\Graficos\comerciar.jpg")
 Dim ax As Byte
 ax = RandomNumber(1, 9)
 Call audio.PlayWave("comerciante" & ax & ".wav")
'Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotonComprar.jpg")
'Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botonvender.jpg")
'Call MakeWindowTransparent(frmComerciar.hWnd, 200)

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = 0 Then
    'Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotonComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    'Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botonvender.jpg")
    Image1(1).Tag = 1
End If
End Sub

Private Sub Image1_Click(Index As Integer)

Call audio.PlayWave(SND_CLICK)

If List1(Index).List(List1(Index).ListIndex) = "Nada" Or _
   List1(Index).ListIndex < 0 Then Exit Sub

Select Case Index
    Case 0
        frmComerciar.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
          If UserGLD >= NPCInventory(List1(0).ListIndex + 1).valor * Val(cantidad) Then
                SendData ("COMP" & "," & List1(0).ListIndex + 1 & "," & cantidad.Text)
                
        Else
            AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If
   Case 1
        LastIndex2 = List1(1).ListIndex
        If UserInventory(List1(1).ListIndex + 1).Equipped = 0 Then
            SendData ("VEND" & "," & List1(1).ListIndex + 1 & "," & cantidad.Text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes vender el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
                
End Select
List1(0).Clear

List1(1).Clear

NPCInvDim = 0
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        If Image1(0).Tag = 1 Then
                'Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotonComprarApretado.jpg")
                Image1(0).Tag = 0
               ' Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botonvender.jpg")
                Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
               ' Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botonvenderapretado.jpg")
                Image1(1).Tag = 0
               ' Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotonComprar.jpg")
                Image1(0).Tag = 1
        End If
        
End Select
End Sub

Private Sub Image2_Click()
SendData ("FINCOM")
End Sub

Private Sub List1_Click(Index As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

Select Case Index
    Case 0
        'Label1(0).Caption = NPCInventory(List1(0).ListIndex + 1).name
        Label1(1).Caption = NPCInventory(List1(0).ListIndex + 1).valor
        Label1(2).Caption = NPCInventory(List1(0).ListIndex + 1).Amount
        Select Case NPCInventory(List1(0).ListIndex + 1).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe:" & NPCInventory(List1(0).ListIndex + 1).MaxHIT
                Label1(4).Caption = "Min Golpe:" & NPCInventory(List1(0).ListIndex + 1).MinHIT
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3
                Label1(3).Visible = False
                'Label1(4).Caption = "Defensa:" & NPCInventory(List1(0).ListIndex + 1).DefMin & "/" & NPCInventory(List1(0).ListIndex + 1).DefMax
                Label1(4).Caption = "Defensa:" & NPCInventory(List1(0).ListIndex + 1).DefMin & "/" & NPCInventory(List1(0).ListIndex + 1).DefMax
                'Label1(4).Caption = "Defensa Cuerpo:" & UserInventory(List1(0).ListIndex + 1).DefCuerpo + 5 & vbCrLf & "Defensa Mágica:" & UserInventory(List1(0).ListIndex + 1).DefMagica
                Label1(4).Visible = True
        End Select
        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, NPCInventory(List1(0).ListIndex + 1).GrhIndex, SR, DR)
    Case 1
        'Label1(0).Caption = UserInventory(List1(1).ListIndex + 1).name
        Label1(1).Caption = Int(UserInventory(List1(1).ListIndex + 1).valor / 3)
        Label1(2).Caption = UserInventory(List1(1).ListIndex + 1).Amount
        Select Case UserInventory(List1(1).ListIndex + 1).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe:" & UserInventory(List1(1).ListIndex + 1).MaxHIT
                Label1(4).Caption = "Min Golpe:" & UserInventory(List1(1).ListIndex + 1).MinHIT
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & NPCInventory(List1(1).ListIndex + 1).DefMin & "/" & NPCInventory(List1(1).ListIndex + 1).DefMax
                'Label1(4).Caption = "Defensa Cuerpo:" & UserInventory(List1(1).ListIndex + 1).DefCuerpo + 5 & vbCrLf & "Defensa Mágica:" & UserInventory(List1(1).ListIndex + 1).DefMagica
                Label1(4).Visible = True
        End Select
        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, UserInventory(List1(1).ListIndex + 1).GrhIndex, SR, DR)
End Select
Picture1.Refresh

End Sub
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = 0 Then
    'Image1(0).Picture = LoadPicture(App.Path & "\Graficos\botonComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
   ' Image1(1).Picture = LoadPicture(App.Path & "\Graficos\botonvender.jpg")
    Image1(1).Tag = 1
End If
End Sub


Private Sub List1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If (Button = vbRightButton) Then
   frmMain.TimerLabel.Enabled = True
    
Dim Esti As Byte
Select Case Index
    Case 0
    Esti = List1(0).ListIndex + 1
    Case 1
    Esti = List1(1).ListIndex + 1
End Select
    
 
    Dim aux As Integer
    Dim Lainv As String
    Dim ob As Integer
    Dim n As Byte
    
  aux = Esti

    'If aux > 0 And aux < MAX_INVENTORY_SLOTS Then _
       ' picInv.ToolTipText = UserInventory(aux).Name
        
        
       If aux = 0 Then Exit Sub
   If IndiceLabel <> aux Then
                
        'Call ReestablecerLabel
                        
        'picInv.FontUnderline = True
        'picInv.ForeColor = vbGreen


'------------------------------------------------------------
Lainv = ""
'pluto:6.0A

If Index = 1 Then ob = UserInventory(Esti).OBJIndex
If Index = 0 Then ob = NPCInventory(Esti).OBJIndex
If ob = 0 Then Exit Sub
If ObjData(ob).MaxHIT > 0 Then Lainv = Lainv & vbNewLine & "Máximo Golpe: " & ObjData(ob).MaxHIT
If ObjData(ob).MinHIT > 0 Then Lainv = Lainv & vbNewLine & "Mínimo Golpe: " & ObjData(ob).MinHIT
If ObjData(ob).MaxDef > 0 Then Lainv = Lainv & vbNewLine & "Máxima Defensa: " & ObjData(ob).MaxDef
If ObjData(ob).MinDef > 0 Then Lainv = Lainv & vbNewLine & "Mínima Defensa: " & ObjData(ob).MinDef

    If ObjData(ob).MaxHIT > 0 Then

If ObjData(ob).Apuñala > 0 Then
Lainv = Lainv & vbNewLine & "Apuñala: Sí"
Else
Lainv = Lainv & vbNewLine & "Apuñala: No"
End If
If ObjData(ob).Envenena > 0 Then
Lainv = Lainv & vbNewLine & "Veneno: Sí"
Else
Lainv = Lainv & vbNewLine & "Veneno: No"
End If

    End If


If ObjData(ob).Magia > 0 Then Lainv = Lainv & vbNewLine & "Mejora Magias: " & ObjData(ob).Magia & "%"
If ObjData(ob).MaxModificador > 0 Then Lainv = Lainv & vbNewLine & "Efecto Máximo: " & ObjData(ob).MaxModificador
If ObjData(ob).MinModificador > 0 Then Lainv = Lainv & vbNewLine & "Efecto Mínimo: " & ObjData(ob).MinModificador
If ObjData(ob).DuracionEfecto > 0 Then Lainv = Lainv & vbNewLine & "Duración Efecto: " & ObjData(ob).DuracionEfecto
If ObjData(ob).MinSed > 0 Then Lainv = Lainv & vbNewLine & "Recupera Sed: " & ObjData(ob).MinSed
If ObjData(ob).MinHam > 0 Then Lainv = Lainv & vbNewLine & "Recupera Hambre: " & ObjData(ob).MinHam
If ObjData(ob).MinSta > 0 Then Lainv = Lainv & vbNewLine & "Recupera Energía: " & ObjData(ob).MinSta

If ObjData(ob).MinSkill > 0 Then
Dim ala As Integer
ala = ObjData(ob).MinSkill
If UCase$(UserClase) <> "PIRATA" And UCase$(UserClase) <> "PESCADOR" Then ala = ala * 2
Lainv = Lainv & vbNewLine & "Skill Mínimo: " & ala
End If

If ObjData(ob).objetoespecial > 0 Then
Select Case ObjData(ob).objetoespecial
Case 1
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Ahorro 33% Flechas"
Case 53
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Ahorro 50% Flechas"
Case 54
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Ahorro 75% Flechas"
Case 2
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Fuerza +5"
Case 3
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Fuerza +2"
Case 4
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Fuerza +3"
Case 5
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Agilidad +5"
Case 6
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Agilidad +2"
Case 7
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Agilidad +3"
Case 8
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Mana +100"
Case 9
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Mana +200"
Case 10
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Mana +300"
End Select
End If
If ObjData(ob).nocaer > 0 Then Lainv = Lainv & vbNewLine & "Habilidad Mágica: No se cae al morir."

'peso
If Index = 1 Then
Lainv = Lainv & vbNewLine & "Cantidad: " & UserInventory(aux).Amount
Lainv = Lainv & vbNewLine & "Peso Unidad: " & ObjData(ob).peso & " kg."
Lainv = Lainv & vbNewLine & "Peso Total: " & ObjData(ob).peso * UserInventory(aux).Amount & " Kg."
Else
Lainv = Lainv & vbNewLine & "Cantidad: " & NPCInventory(aux).Amount
Lainv = Lainv & vbNewLine & "Peso Unidad: " & ObjData(ob).peso & " kg."
'Lainv = Lainv & vbNewLine & "Peso Total: " & ObjData(ob).peso * NPCInventory(aux).Amount & " Kg."
End If


If ObjData(ob).SkArco > 0 And ObjData(ob).proyectil > 0 Then Lainv = Lainv & vbNewLine & "Skill Mínimo: " & ObjData(ob).SkArco
If ObjData(ob).SkArma > 0 And ObjData(ob).proyectil = 0 Then Lainv = Lainv & vbNewLine & "Skill Mínimo: " & ObjData(ob).SkArma

If ObjData(ob).Vendible = 0 Then
Lainv = Lainv & vbNewLine & "Vendible: Sí"
Else
Lainv = Lainv & vbNewLine & "Vendible: No"
End If

If ObjData(ob).razaelfa > 0 Then Lainv = Lainv & vbNewLine & "Raza: Elfos."
If ObjData(ob).RazaEnana > 0 Then Lainv = Lainv & vbNewLine & "Raza: Enanos."
If ObjData(ob).razahumana > 0 Then Lainv = Lainv & vbNewLine & "Raza: Humanos."
If ObjData(ob).razaorca > 0 Then Lainv = Lainv & vbNewLine & "Raza: Orcos."
If ObjData(ob).razavampiro > 0 Then Lainv = Lainv & vbNewLine & "Raza: Vampiros."
If ObjData(ob).Real > 0 Then Lainv = Lainv & vbNewLine & "Armada: Armada Real."
If ObjData(ob).Caos > 0 Then Lainv = Lainv & vbNewLine & "Armada: Armada del Caos."
If ObjData(ob).Hombre > 0 Then Lainv = Lainv & vbNewLine & "Sexo: Hombres."
If ObjData(ob).Mujer > 0 Then Lainv = Lainv & vbNewLine & "Sexo: Mujeres."
If ObjData(ob).ObjetoClan <> "" Then Lainv = Lainv & vbNewLine & "Clan: " & ObjData(ob).ObjetoClan
If ObjData(ob).HechizoIndex > 0 Then
Dim afeti As String
Lainv = Lainv & vbNewLine & "Skill de Magia Necesario: " & Hechizos(ObjData(ob).HechizoIndex).MinSkill
Lainv = Lainv & vbNewLine & "Mana Necesario: " & Hechizos(ObjData(ob).HechizoIndex).ManaRequerido
If Hechizos(ObjData(ob).HechizoIndex).MaxHP > 0 Then Lainv = Lainv & vbNewLine & "P.Máximo: " & Hechizos(ObjData(ob).HechizoIndex).MaxHP
If Hechizos(ObjData(ob).HechizoIndex).MinHP > 0 Then Lainv = Lainv & vbNewLine & "P.Mínimo: " & Hechizos(ObjData(ob).HechizoIndex).MinHP
If Hechizos(ObjData(ob).HechizoIndex).MaxFuerza > 0 Then Lainv = Lainv & vbNewLine & "F.Máximo: " & Hechizos(ObjData(ob).HechizoIndex).MaxFuerza
If Hechizos(ObjData(ob).HechizoIndex).MinFuerza > 0 Then Lainv = Lainv & vbNewLine & "F.Mínimo: " & Hechizos(ObjData(ob).HechizoIndex).MinFuerza
If Hechizos(ObjData(ob).HechizoIndex).MaxAgilidad > 0 Then Lainv = Lainv & vbNewLine & "A.Máximo: " & Hechizos(ObjData(ob).HechizoIndex).MaxAgilidad
If Hechizos(ObjData(ob).HechizoIndex).MinAgilidad > 0 Then Lainv = Lainv & vbNewLine & "A.Mínimo: " & Hechizos(ObjData(ob).HechizoIndex).MinAgilidad
If Hechizos(ObjData(ob).HechizoIndex).MaxHam > 0 Then Lainv = Lainv & vbNewLine & "H.Máximo: " & Hechizos(ObjData(ob).HechizoIndex).MaxHam
If Hechizos(ObjData(ob).HechizoIndex).MinHam > 0 Then Lainv = Lainv & vbNewLine & "H.Mínimo: " & Hechizos(ObjData(ob).HechizoIndex).MinHam
If Hechizos(ObjData(ob).HechizoIndex).MaxSed > 0 Then Lainv = Lainv & vbNewLine & "S.Máximo: " & Hechizos(ObjData(ob).HechizoIndex).MaxSed
If Hechizos(ObjData(ob).HechizoIndex).MinSed > 0 Then Lainv = Lainv & vbNewLine & "S.Mínimo: " & Hechizos(ObjData(ob).HechizoIndex).MinSed

'target
If Hechizos(ObjData(ob).HechizoIndex).Target = 1 Then afeti = "Sólo Usuarios."
If Hechizos(ObjData(ob).HechizoIndex).Target = 2 Then afeti = "Sólo Npc´s."
If Hechizos(ObjData(ob).HechizoIndex).Target = 3 Then afeti = "Usuarios y Npc´s."
If Hechizos(ObjData(ob).HechizoIndex).Target = 4 Then afeti = "Terreno."
Lainv = Lainv & vbNewLine & "Objetivo: " & afeti

End If


Dim pit As Byte
For n = 1 To 20
If UCase$(ObjData(ob).ClaseProhibida(n)) = UCase$(UserClase) Then pit = 1
Next
If pit = 1 Then
Lainv = Lainv & vbNewLine & "El " & UserClase & " NO puede usarlo."
Else
Lainv = Lainv & vbNewLine & "El " & UserClase & " puede usarlo."
End If
pit = 0

'-----------------------------------------------------------
'-----------------------------------------------------------
'fabricación---------
If ObjData(ob).LingH > 0 Or ObjData(ob).LingP > 0 Or ObjData(ob).LingO > 0 Or ObjData(ob).Madera > 0 Or ObjData(ob).Diamantes Or ObjData(ob).Gemas > 0 Then
Lainv = Lainv & vbNewLine & vbNewLine & "Se puede Fabricar Con:"
If ObjData(ob).LingO > 0 Then Lainv = Lainv & vbNewLine & "Lingotes Oro: " & ObjData(ob).LingO
If ObjData(ob).LingP > 0 Then Lainv = Lainv & vbNewLine & "Lingotes Plata: " & ObjData(ob).LingP
If ObjData(ob).LingH > 0 Then Lainv = Lainv & vbNewLine & "Lingotes Hierro: " & ObjData(ob).LingH
If ObjData(ob).Madera > 0 Then Lainv = Lainv & vbNewLine & "Madera: " & ObjData(ob).Madera
If ObjData(ob).Diamantes > 0 Then Lainv = Lainv & vbNewLine & "Diamantes: " & ObjData(ob).Diamantes
If ObjData(ob).Gemas > 0 Then Lainv = Lainv & vbNewLine & "Gemas: " & ObjData(ob).Gemas
End If
'--------------------
    
    If Index = 1 Then
    Call Mostrar_ToolTip(Form2, UserInventory(aux).name & vbNewLine & String(Len(UserInventory(aux).name), "_") & vbNewLine & Lainv, _
    &H80000018, vbBlack, App.Path & "\Graficos\aodrag.ico")
    Else
    Call Mostrar_ToolTip(Form2, NPCInventory(aux).name & vbNewLine & String(Len(NPCInventory(aux).name), "_") & vbNewLine & Lainv, _
    &H80000018, vbBlack, App.Path & "\Graficos\aodrag.ico")

    End If
            
        
        
                
        IndiceLabel = aux
        
    End If
    
    End If 'boton derecho

End Sub
