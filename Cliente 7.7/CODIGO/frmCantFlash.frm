VERSION 5.00
Begin VB.Form frmCantFlash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCantFlash.frx":0000
   ScaleHeight     =   2985
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   1460
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   960
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   960
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "frmCantFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Ergs = "/ingresar" Then SendData ("/DEPOSITAR " & Val(Text1))
If Ergs = "/retirar" Then SendData ("/RETIRAR " & Val(Text1))
frmCantFlash.Visible = False
End Sub

Private Sub Command2_Click()
frmCantFlash.Visible = False
Ergs = ""
End Sub

Private Sub Form_Load()
frmCantFlash.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Image1_Click()
If Ergs = "/ingresar" Then SendData ("/DEPOSITAR " & Val(Text1))
If Ergs = "/retirar" Then SendData ("/RETIRAR " & Val(Text1))
frmCantFlash.Visible = False
End Sub

Private Sub Image2_Click()
frmCantFlash.Visible = False
Ergs = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (Index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub
