VERSION 5.00
Begin VB.Form Cerrando 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Cerrando AodraG"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Cerrando Server AodraG. Espere unos segundos........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   6255
   End
End
Attribute VB_Name = "Cerrando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Visible = True
End Sub
