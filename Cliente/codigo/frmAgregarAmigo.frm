VERSION 5.00
Begin VB.Form frmAgregarAmigo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   2580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtNameFriend 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Escribe el nombre de tu amigo:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmAgregarAmigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call WriteaddFriend(txtNameFriend.Text)
    Unload Me
End Sub
