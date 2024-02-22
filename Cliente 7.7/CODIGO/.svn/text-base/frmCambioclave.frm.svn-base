VERSION 5.00
Begin VB.Form frmCambioclave 
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3360
      MouseIcon       =   "frmCambioclave.frx":0000
      Picture         =   "frmCambioclave.frx":0CCA
      Top             =   5400
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   840
      MouseIcon       =   "frmCambioclave.frx":4F83
      Picture         =   "frmCambioclave.frx":5C4D
      Top             =   5400
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva Clave para la Cuenta:"
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
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   4095
   End
End
Attribute VB_Name = "frmCambioclave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
frmCambioclave.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
Text1.Text = ""
End Sub



Private Sub Image1_Click()
If Ergs = "/cambioclave" Then
Dim j$
Dim stx As String
j$ = MD5String(Text1)
stx = "/PASSWD " & j$
Call SendData(stx)
frmCambioclave.Visible = False
End If
End Sub

Private Sub Image2_Click()
Ergs = ""
Unload Me
End Sub
