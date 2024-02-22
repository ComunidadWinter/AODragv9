VERSION 5.00
Begin VB.Form frmEntrenador 
   BorderStyle     =   0  'None
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6765
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEntrenador.frx":0000
   ScaleHeight     =   4365
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2130
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   3900
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3480
      MouseIcon       =   "frmEntrenador.frx":2038D
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1080
      MouseIcon       =   "frmEntrenador.frx":21057
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   2295
   End
End
Attribute VB_Name = "frmEntrenador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmEntrenador.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
 Call audio.PlayWave("177.wav")
End Sub
Private Sub Command1_Click()
Call SendData("ENTR" & lstCriaturas.ListIndex + 1)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Image1_Click()
Call SendData("ENTR" & lstCriaturas.ListIndex + 1)
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
