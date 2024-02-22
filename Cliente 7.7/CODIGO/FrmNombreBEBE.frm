VERSION 5.00
Begin VB.Form FrmNombreBEBE 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3720
      MouseIcon       =   "FrmNombreBEBE.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "FrmNombreBEBE.frx":0CCA
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1080
      MouseIcon       =   "FrmNombreBEBE.frx":4F83
      MousePointer    =   99  'Custom
      Picture         =   "FrmNombreBEBE.frx":5C4D
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre para tu futuro descendiente"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   5175
   End
End
Attribute VB_Name = "FrmNombreBEBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Generobebe As String


Private Sub Form_Load()
FrmNombreBEBE.Picture = LoadPicture(DirGraficos & "ventanas.jpg")

If Generagenero = 0 Then
a = RandomNumber(1, 2)
Generagenero = a
End If
If Generagenero = 1 Then Generobebe = "Hombre"
If Generagenero = 2 Then Generobebe = "Mujer"
Label1.Caption = Generobebe
End Sub
'End Sub

Private Sub Image1_Click()
Dim nbebe As String
nbebe = Text1.Text
If Not AsciiValidos(nbebe) Then
MsgBox "Nombre con caracteres invalidos."
Exit Sub
End If

SendData "NBEB" & nbebe & "," & Generobebe
Generagenero = 0
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
