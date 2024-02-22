VERSION 5.00
Begin VB.Form frmNaci 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Encuesta"
   ClientHeight    =   11520
   ClientLeft      =   90
   ClientTop       =   -420
   ClientWidth     =   12000
   LinkTopic       =   "Form4"
   Picture         =   "frmNaci.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4440
      TabIndex        =   0
      Top             =   3720
      Width           =   2535
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Europa"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "América"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "¿Desde donde Juegas?"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3120
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   4800
      MouseIcon       =   "frmNaci.frx":75742
      MousePointer    =   99  'Custom
      Picture         =   "frmNaci.frx":7640C
      Top             =   5640
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNaci.frx":7A76C
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   9255
   End
End
Attribute VB_Name = "frmNaci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Option1.Value = False
Option2.Value = False
End Sub



Private Sub Image1_Click()
If Option1.Value = True Then
Naci = 1
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Server", 1)
End If

If Option2.Value = True Then
Naci = 2
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Server", 2)
End If


If Naci > 0 Then
frmNaci.Visible = False
End If
End Sub
