VERSION 5.00
Begin VB.Form frmFAMA 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   LinkTopic       =   "Form4"
   Picture         =   "frmFAMA.frx":0000
   ScaleHeight     =   7215
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2280
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Fama de tu Personaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
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
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel de Fama:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFAMA.frx":26121
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos Popularidad:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4320
      Width           =   3135
   End
End
Attribute VB_Name = "frmFAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Label3.Caption = frmMain.FamaLabel.Caption
Select Case Val(Label3.Caption)


Case Is > 700001
Label5.Caption = "Leyenda del Dragón"
Case 500001 To 700000
Label5.Caption = "Héroe del AodraG"
Case 350001 To 500000
Label5.Caption = "Campeón"
Case 200001 To 350000
Label5.Caption = "Señor"
Case 120001 To 200000
Label5.Caption = "Admirable"
Case 60001 To 120000
Label5.Caption = "Caballero"
Case 30001 To 60000
Label5.Caption = "Montaraz"
Case 15001 To 30000
Label5.Caption = "Jinete"
Case 8001 To 15000
Label5.Caption = "Afamado"
Case 4001 To 8000
Label5.Caption = "Honorable"
Case 2001 To 4000
Label5.Caption = "Aprendiz"
Case 1001 To 2000
Label5.Caption = "Escudero"
Case 601 To 1000
Label5.Caption = " Aldeano"
Case 301 To 600
Label5.Caption = "Mendigo"
Case 100 To 300
Label5.Caption = "Persigue-Pollos"
Case Is < 100
Label5.Caption = "Recoge-Manzanas"

End Select
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

