VERSION 5.00
Begin VB.Form frmComerciarUsu 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   3000
   ClientTop       =   0
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3180
      Left            =   3600
      TabIndex        =   7
      Top             =   2460
      Width           =   2610
   End
   Begin VB.TextBox txtCant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Text            =   "1"
      Top             =   5820
      Width           =   975
   End
   Begin VB.OptionButton optQue 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3720
      MouseIcon       =   "frmComerciarUsu.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1980
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.OptionButton optQue 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   5280
      MouseIcon       =   "frmComerciarUsu.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1980
      Width           =   195
   End
   Begin VB.ListBox List2 
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
      Height          =   2340
      Left            =   600
      TabIndex        =   2
      Top             =   2220
      Width           =   2730
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   540
      Left            =   900
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   1605
      Width           =   540
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   3960
      Picture         =   "frmComerciarUsu.frx":1994
      Top             =   6600
      Width           =   2220
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   4200
      Picture         =   "frmComerciarUsu.frx":675D
      Top             =   6240
      Width           =   1650
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   1080
      Picture         =   "frmComerciarUsu.frx":A9E6
      Top             =   6600
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1080
      Picture         =   "frmComerciarUsu.frx":ED6A
      Top             =   6240
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   5820
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad: 0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   5820
      Width           =   2415
   End
   Begin VB.Label lblEstadoResp 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando respuesta..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   2850
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'[Alejo]

Private Sub Form_Deactivate()
Me.SetFocus
Picture1.SetFocus

End Sub

Private Sub Form_Load()
'Carga las imagenes...?
frmComerciarUsu.Picture = LoadPicture(DirGraficos & "ComerciarUsu.jpg")

lblEstadoResp.Visible = False
'Call MakeWindowTransparent(frmComerciarUsu.hWnd, 200)

End Sub

Private Sub Form_LostFocus()
Me.SetFocus
Picture1.SetFocus

End Sub

Private Sub Image1_Click()
Call SendData("COMUSUOK")
End Sub

Private Sub Image2_Click()
Call SendData("COMUSUNO")
End Sub

Private Sub Image3_Click()
If optQue(0).Value = True Then
    If List1.ListIndex < 0 Then Exit Sub
    If List1.ItemData(List1.ListIndex) <= 0 Then Exit Sub
    
    If Val(txtCant.Text) > List1.ItemData(List1.ListIndex) Or _
        Val(txtCant.Text) <= 0 Then Exit Sub
ElseIf optQue(1).Value = True Then
    If Val(txtCant.Text) > UserGLD Then Exit Sub
End If

If optQue(0).Value = True Then
    Call SendData("OFRECER" & List1.ListIndex + 1 & "," & Trim(Val(txtCant.Text)))
Image3.Visible = False
ElseIf optQue(1).Value = True Then
    Call SendData("OFRECER" & FLAGORO & "," & Trim(Val(txtCant.Text)))
Image3.Visible = False
Else
    Exit Sub
End If

lblEstadoResp.Visible = True

End Sub

Private Sub Image4_Click()
Call SendData("FINCOMUSU")
End Sub

Private Sub List1_Click()
'DibujaGrh UserInventory(List1.ListIndex + 1).GrhIndex

End Sub

Public Sub DibujaGrh(Grh As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, Grh, SR, DR)

End Sub

Private Sub List2_Click()
If List2.ListIndex >= 0 Then
    DibujaGrh OtroInventario(List2.ListIndex + 1).GrhIndex
    Label3.Caption = "Cantidad: " & List2.ItemData(List2.ListIndex)
    Image1.Visible = True
    Image2.Visible = True
Else
    Image1.Visible = False
    Image2.Visible = False
End If

End Sub

Private Sub optQue_Click(Index As Integer)
Select Case Index
Case 0
    List1.Enabled = True
Case 1
    List1.Enabled = False
End Select

End Sub

Private Sub txtCant_KeyDown(KeyCode As Integer, Shift As Integer)
If Not (KeyCode >= 48 And KeyCode <= 57) Then
    KeyCode = 0
End If

End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
    KeyAscii = 0
End If

End Sub

'[/Alejo]

