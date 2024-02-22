VERSION 5.00
Begin VB.Form frmRecuperar 
   BorderStyle     =   0  'None
   Caption         =   "Recuperar Personaje"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   LinkTopic       =   "Form3"
   ScaleHeight     =   7200
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000040&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3240
      MouseIcon       =   "frmRecuperar.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmRecuperar.frx":0CCA
      Top             =   6600
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   960
      MouseIcon       =   "frmRecuperar.frx":4E86
      MousePointer    =   99  'Custom
      Picture         =   "frmRecuperar.frx":5B50
      Top             =   6600
      Width           =   1680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRecuperar.frx":9EB0
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Introduce Nombre del Personaje "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
   End
End
Attribute VB_Name = "frmRecuperar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If frmRecuperar.Text1 = "" Then
MsgBox "Nombre personaje no válido"
Exit Sub
End If
Call SendData("RPERSS" & frmRecuperar.Text1.Text)
frmRecuperar.Visible = False
End Sub

Private Sub Command2_Click()
frmRecuperar.Visible = False
End Sub

Private Sub Form_Load()
frmRecuperar.Picture = LoadPicture(App.Path & "\Graficos\ventanas.jpg")
End Sub


Private Sub Image1_Click()
If frmRecuperar.Text1 = "" Then
MsgBox "Nombre personaje no válido"
Exit Sub
End If
Call SendData("RPERSS" & frmRecuperar.Text1.Text)
frmRecuperar.Visible = False
End Sub

Private Sub Image2_Click()
frmRecuperar.Visible = False
End Sub
