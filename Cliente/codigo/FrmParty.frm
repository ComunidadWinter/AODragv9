VERSION 5.00
Begin VB.Form FrmParty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de party"
   ClientHeight    =   3570
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   3360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Abandonar party"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3090
      Begin VB.CommandButton Command1 
         Caption         =   "Invitar"
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Nuevo lider"
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Expulsar"
         Height          =   375
         Left            =   210
         TabIndex        =   2
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   1035
         ItemData        =   "FrmParty.frx":0000
         Left            =   120
         List            =   "FrmParty.frx":0002
         TabIndex        =   1
         Top             =   360
         Width           =   2850
      End
   End
End
Attribute VB_Name = "FrmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    Call ChangeCursorMain(cur_Action, frmMain)
    SolicitudParty = True
End Sub

Private Sub Command2_Click()
    Call WriteExpulsarDeParty(FrmParty.List1)
    Call WriteActualizarParty
End Sub

Private Sub Command3_Click()
    Call WriteNuevoLiderDeParty(FrmParty.List1)
    Call WriteActualizarParty
End Sub

Private Sub Command4_Click()
    Call WriteAbandonarParty
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    FrmParty.List1.Clear
    Call WriteActualizarParty
End Sub
