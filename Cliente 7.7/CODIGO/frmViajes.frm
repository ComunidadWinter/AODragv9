VERSION 5.00
Begin VB.Form frmViajes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Viajes AOdrag"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H80000004&
      Height          =   1785
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H80000004&
      Height          =   1785
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   3840
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   4200
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   4350
      Left            =   0
      Picture         =   "frmViajes.frx":0000
      Top             =   0
      Width           =   6750
   End
End
Attribute VB_Name = "frmViajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
List1.Clear
List2.Clear
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Image3_Click()
If List1.Text = "Ullathorpe" Then SendData ("/VIAJAR ULLA")
If List1.Text = "Nix" Then SendData ("/VIAJAR NIX")
If List1.Text = "Banderville" Then SendData ("/VIAJAR BANDER")
If List1.Text = "Lindos" Then SendData ("/VIAJAR LINDOS")
If List1.Text = "Arghal" Then SendData ("/VIAJAR ARGHAL")
If List1.Text = "Nueva Esperanza" Then SendData ("/VIAJAR ESPERANZA")
If List1.Text = "Atlantis" Then SendData ("/VIAJAR ATLANTIS")
If List1.Text = "Ciudad Caos" Then SendData ("/VIAJAR CAOS")
If List1.Text = "Desierto de Rinkel" Then SendData ("/VIAJAR RINKEL")
If List1.Text = "Ciudad Descanso" Then SendData ("/VIAJAR DESCANSO")
End Sub

Private Sub List1_Click()
Label1.Caption = "Valor: " & List2.List(List1.ListIndex)
Label2.Caption = "Ciudad: " & List1.Text
End Sub


