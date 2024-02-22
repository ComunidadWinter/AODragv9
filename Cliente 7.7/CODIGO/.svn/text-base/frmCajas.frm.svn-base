VERSION 5.00
Begin VB.Form frmCajas 
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   Picture         =   "frmCajas.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   300
      Left            =   2280
      MouseIcon       =   "frmCajas.frx":5FAB
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":6C75
      Top             =   6240
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   5
      Left            =   3600
      MouseIcon       =   "frmCajas.frx":AF2E
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":BBF8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   4
      Left            =   960
      MouseIcon       =   "frmCajas.frx":C3BC
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":D086
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   3
      Left            =   3600
      MouseIcon       =   "frmCajas.frx":D84A
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":E514
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   2
      Left            =   960
      MouseIcon       =   "frmCajas.frx":ECD8
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":F9A2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   1
      Left            =   3480
      MouseIcon       =   "frmCajas.frx":10166
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":10E30
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   0
      Left            =   960
      MouseIcon       =   "frmCajas.frx":115F4
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":122BE
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "frmCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim n As Byte
For n = 0 To 3
frmCajas.Image1(n).Picture = LoadPicture(DirGraficos & "baul2.jpg")
Next
End Sub

Private Sub Image1_Click(Index As Integer)
Dim index2 As Byte
index2 = Index + 1
SendData ("/BOVEDA" & index2)
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
frmBanquero.Show vbModal
End Sub
