VERSION 5.00
Begin VB.Form frmCiudad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox list1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCiudad.frx":0000
      Left            =   1080
      List            =   "frmCiudad.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   4035
      Top             =   4200
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   735
      Top             =   4215
      Width           =   960
   End
   Begin VB.Label Descrip 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione su hogar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1995
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   2985
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione su hogar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1425
   End
End
Attribute VB_Name = "frmCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Integer

For i = 1 To NUMCIUDADES
 list1.AddItem Ciudades(i)
Next i

list1.Text = list1.List(1)
Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónVolver.jpg")
Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguiente.jpg")
End Sub

Private Sub list1_Click()
Descrip.Caption = CityDesc(list1.ListIndex + 1)
End Sub
