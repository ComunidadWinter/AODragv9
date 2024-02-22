VERSION 5.00
Begin VB.Form frmMensaje 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2160
      MouseIcon       =   "frmMensaje.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmMensaje.frx":0CCA
      Top             =   6360
      Width           =   1680
   End
   Begin VB.Label msg 
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
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   5295
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmMensaje.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
End Sub



Private Sub Form_Deactivate()
    If Visible Then Me.SetFocus
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

