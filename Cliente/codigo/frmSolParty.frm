VERSION 5.00
Begin VB.Form frmSolParty 
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdRechazas 
      Caption         =   "Rechazar"
      Height          =   360
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Text 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmSolParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    Call WriteRespuestaPeticionAParty(True)
    Unload Me
End Sub

Private Sub cmdRechazas_Click()
    Call WriteRespuestaPeticionAParty(False)
    Unload Me
End Sub
