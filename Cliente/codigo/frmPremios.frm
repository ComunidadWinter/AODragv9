VERSION 5.00
Begin VB.Form frmPremios 
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5145
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstPremios 
      Height          =   3960
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdCanjear 
      Caption         =   "Canjear"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   530
      Left            =   2820
      ScaleHeight     =   525
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Width           =   500
   End
   Begin VB.Label lblCantidad 
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblDespues 
      Caption         =   "Puntos despues del canje: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label lblDisponibles 
      Caption         =   "Puntos disponibles: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblCoste 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coste: 0"
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
      Height          =   195
      Left            =   2760
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "frmPremios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCanjear_Click()
    Call WritePremios(1, lstPremios.ListIndex + 1)
End Sub

Private Sub cmdVolver_Click()
    lstPremios.Clear
    Unload Me
End Sub

Private Sub Form_Load()
    lblDisponibles.Caption = "Puntos disponibles: " & UserDragCreditos
End Sub

Private Sub lstPremios_Click()
    lblDespues.Caption = "Puntos despues del canje: " & UserDragCreditos - PremiosInv(lstPremios.ListIndex + 1).Puntos
    lblCoste.Caption = "Coste: " & PremiosInv(lstPremios.ListIndex + 1).Puntos
    lblCantidad.Caption = "Cantidad: " & PremiosInv(lstPremios.ListIndex + 1).Cantidad
End Sub
