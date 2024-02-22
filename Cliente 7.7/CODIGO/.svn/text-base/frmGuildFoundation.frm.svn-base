VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6120
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      BackColor       =   &H00000040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox txtClanName 
      BackColor       =   &H00000040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3480
      MouseIcon       =   "frmGuildFoundation.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmGuildFoundation.frx":0CCA
      Top             =   6360
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   840
      MouseIcon       =   "frmGuildFoundation.frx":519A
      MousePointer    =   99  'Custom
      Picture         =   "frmGuildFoundation.frx":5E64
      Top             =   6360
      Width           =   1680
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sitio Web Del Clan"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Del Clan:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGuildFoundation.frx":A1C4
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Form_Load()
frmGuildFoundation.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
End Sub

Private Sub Image1_Click()

If Len(txtClanName.Text) <= 16 And Len(txtClanName.Text) >= 4 Then
    If Not AsciiValidos(txtClanName) Then
        MsgBox "Nombre invalido."
        Exit Sub
    End If
Else
        MsgBox "Nombre demasiado extenso o corto (de 4 a 12 caracteres)."
        Exit Sub
End If

ClanName = txtClanName
Site = Text2
Unload Me
frmGuildDetails.Show
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
