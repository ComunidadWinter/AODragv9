VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   2460
   ClientTop       =   285
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBancoObj.frx":0000
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
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3120
      TabIndex        =   6
      Text            =   "1"
      Top             =   5985
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
      Left            =   3600
      MouseIcon       =   "frmBancoObj.frx":285A1
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
      MouseIcon       =   "frmBancoObj.frx":2926B
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1560
      Width           =   3090
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   3840
      MouseIcon       =   "frmBancoObj.frx":29F35
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   600
      MouseIcon       =   "frmBancoObj.frx":2ABFF
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      MouseIcon       =   "frmBancoObj.frx":2B8C9
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   210
      Left            =   3000
      TabIndex        =   7
      Top             =   5760
      Width           =   840
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
      TabIndex        =   5
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
      Left            =   3990
      TabIndex        =   4
      Top             =   840
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
      ForeColor       =   &H00000040&
      Height          =   210
      Index           =   2
      Left            =   2730
      TabIndex        =   3
      Top             =   1170
      Width           =   105
   End
End
Attribute VB_Name = "frmBancoObj"
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

Private Sub Command1_Click(Index As Integer)

End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub


Private Sub Form_Load()
'Cargamos la interfase
Me.Picture = LoadPicture(App.Path & "\Graficos\comerciar.jpg")
'Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotonComprar.jpg")
'Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botonvender.jpg")
'Call MakeWindowTransparent(frmBancoObj.hWnd, 200)
End Sub


Private Sub Image1_Click(Index As Integer)
Call audio.PlayWave(SND_CLICK)

If List1(Index).List(List1(Index).ListIndex) = "Nada" Or _
   List1(Index).ListIndex < 0 Then Exit Sub

Select Case Index
    Case 0
        frmBancoObj.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
        
        SendData ("RETI" & "," & List1(0).ListIndex + 1 & "," & cantidad.Text)
                
   Case 1
        LastIndex2 = List1(1).ListIndex
        If UserInventory(List1(1).ListIndex + 1).Equipped = 0 Then
            SendData ("DEPO" & "," & List1(1).ListIndex + 1 & "," & cantidad.Text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes depositar el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
                
End Select
List1(0).Clear

List1(1).Clear

NPCInvDim = 0
End Sub

Private Sub Image2_Click()
SendData ("FINBAN")
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
        'Label1(0).Caption = UserBancoInventory(List1(0).ListIndex + 1).name
        Label1(2).Caption = UserBancoInventory(List1(0).ListIndex + 1).Amount
        Select Case UserBancoInventory(List1(0).ListIndex + 1).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe:" & UserBancoInventory(List1(0).ListIndex + 1).MaxHIT
                Label1(4).Caption = "Min Golpe:" & UserBancoInventory(List1(0).ListIndex + 1).MinHIT
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & UserInventory(List1(0).ListIndex + 1).DefMin & "/" & UserInventory(List1(0).ListIndex + 1).DefMax
                'Label1(4).Caption = "Defensa:" & UserBancoInventory(List1(0).ListIndex + 1).DefMin & "/" & UserBancoInventory(List1(0).ListIndex + 1).DefMax
                'Label1(4).Caption = "Defensa Cuerpo:" & UserBancoInventory(List1(0).ListIndex + 1).DefCuerpo + 5 & vbCrLf & "Defensa Mágica:" & UserBancoInventory(List1(0).ListIndex + 1).DefMagica
                Label1(4).Visible = True
        End Select
        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, UserBancoInventory(List1(0).ListIndex + 1).GrhIndex, SR, DR)
    Case 1
        'Label1(0).Caption = UserInventory(List1(1).ListIndex + 1).name
        Label1(2).Caption = UserInventory(List1(1).ListIndex + 1).Amount
        Select Case UserInventory(List1(1).ListIndex + 1).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe:" & UserInventory(List1(1).ListIndex + 1).MaxHIT
                Label1(4).Caption = "Min Golpe:" & UserInventory(List1(1).ListIndex + 1).MinHIT
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & UserInventory(List1(1).ListIndex + 1).DefMin & "/" & UserInventory(List1(1).ListIndex + 1).DefMax
                'Label1(4).Caption = "Defensa Cuerpo:" & UserBancoInventory(List1(1).ListIndex + 1).DefCuerpo + 5 & vbCrLf & "Defensa Mágica:" & UserBancoInventory(List1(1).ListIndex + 1).DefMagica
                Label1(4).Visible = True
        End Select
        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, UserInventory(List1(1).ListIndex + 1).GrhIndex, SR, DR)
End Select
Picture1.Refresh

End Sub

