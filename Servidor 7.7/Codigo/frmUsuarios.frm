VERSION 5.00
Begin VB.Form frmUsuarios 
   BackColor       =   &H00000000&
   Caption         =   "Usuarios"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   6135
      Left            =   120
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------
'               Argentum Online Server
'                       Por
'               Pablo Ignacio Márquez
'        pablomarquez@argentum-online.com.ar
'
' El codigo fuente de Argentum Online no es de dominio
' publico, el codigo es propiedad intelectual de Pablo
' Ignacio Márquez y esta terminantemente prohibido su
' uso, copia, modificacion o difusion.
'-----------------------------------------------------

Private Sub Command1_Click()
Me.Visible = False
End Sub

Private Sub Form_Deactivate()
Me.Visible = False
End Sub

