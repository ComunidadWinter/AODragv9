VERSION 5.00
Begin VB.Form frmQuest 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Misiones"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   Picture         =   "frmQuest.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   300
      Left            =   4080
      MouseIcon       =   "frmQuest.frx":5FAB
      MousePointer    =   99  'Custom
      Picture         =   "frmQuest.frx":6C75
      Top             =   6600
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   600
      MouseIcon       =   "frmQuest.frx":AE95
      MousePointer    =   99  'Custom
      Picture         =   "frmQuest.frx":BB5F
      Top             =   6600
      Width           =   2250
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NO CUMPLES TODOS LOS REQUISITOS PARA PODER COMENZAR ESTA MISIÓN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   2880
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Nivelminimo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Nivelmaximo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel Mínimo:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel Máximo:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Misiones 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Tiempo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo Restante:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Misiones Completadas:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00000080&
      Height          =   1575
      Left            =   3240
      TabIndex        =   5
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Recompensa"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00000080&
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Objetivos"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
SendData ("ABORTQ")
Unload frmQuest
End Sub

Private Sub Label10_Click()



End Sub

Private Sub Image2_Click()
Unload frmQuest
End Sub

Private Sub Label7_Click()

End Sub
