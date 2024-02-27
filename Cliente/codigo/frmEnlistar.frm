VERSION 5.00
Begin VB.Form frmEnlistar 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3840
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   990
   End
   Begin VB.CommandButton cmdEnlistar 
      Caption         =   "¡Enlistar!"
      Height          =   360
      Left            =   2400
      TabIndex        =   0
      Top             =   2160
      Width           =   990
   End
   Begin VB.Label lblEstasApunto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEnlistar.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3420
   End
   Begin VB.Label lblATENCIÓN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡ATENCIÓN!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   1770
   End
End
Attribute VB_Name = "frmEnlistar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnlistar_Click()
    Call WriteEnlist
End Sub

Private Sub cmdVolver_Click()
    Unload Me
End Sub
