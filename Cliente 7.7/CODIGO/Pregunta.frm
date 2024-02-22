VERSION 5.00
Begin VB.Form Pregunta 
   Caption         =   "Comprobando Macro Asistido"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Pregunta.frx":0000
   LinkTopic       =   "Form4"
   Picture         =   "Pregunta.frx":1CCA
   ScaleHeight     =   7170
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Responder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MouseIcon       =   "Pregunta.frx":2EC4C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   4800
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   360
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Respuesta Incorrecta!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Introduce aquí el resultado y pulsa el botón RESPONDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1800
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo Restante:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Responde rápidamente esta pregunta....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   5655
   End
End
Attribute VB_Name = "Pregunta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Byte
Dim b As Byte
Dim c As Byte
Public Tmacro As Byte

Private Sub Command1_Click()
If Val(Pregunta.Text1.Text) = c Then
Pregunta.Timer1.Enabled = False
Tmacro = 20
Unload Me
Else
Beep
Pregunta.Label6.Visible = True
End If

End Sub

Private Sub Form_Load()
Tmacro = 20
'Pregunta.Timer1.Enabled = True
a = RandomNumber(1, 4)
b = RandomNumber(1, 4)
c = a + b
Label2.Caption = "¿Cuanto suman " & a & " + " & b & " ?"
End Sub

Private Sub Timer1_Timer()

Tmacro = Tmacro - 1

If Tmacro < 1 Then
SendData ("B2")

Tmacro = 20
Unload Pregunta
Pregunta.Timer1.Enabled = False
End If
Label3.Caption = Tmacro
End Sub
