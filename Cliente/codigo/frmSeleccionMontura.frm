VERSION 5.00
Begin VB.Form frmSeleccionMontura 
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3975
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
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "Ver estadisticas"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox listMonturas 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona la montura que quieres ver..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmSeleccionMontura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSiguiente_Click()
    MonturaSeleccionada = listMonturas.ListIndex + 1
    Call WriteSolInfMonturaClient(1)
    Unload Me
End Sub

Private Sub cmdVolver_Click()
    listMonturas.Clear
    Unload Me
End Sub

