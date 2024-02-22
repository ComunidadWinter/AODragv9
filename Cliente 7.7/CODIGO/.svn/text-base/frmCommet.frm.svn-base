VERSION 5.00
Begin VB.Form frmCommet 
   BorderStyle     =   0  'None
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6915
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
   ScaleHeight     =   3915
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   1000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4975
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3840
      MouseIcon       =   "frmCommet.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   840
      MouseIcon       =   "frmCommet.frx":0CCA
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   2295
   End
End
Attribute VB_Name = "frmCommet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nombre As String
Private Sub Form_Load()
frmCommet.Picture = LoadPicture(DirGraficos & "Commet.jpg")
End Sub
Private Sub Command1_Click()


If Text1 = "" Then
    MsgBox "Debes redactar un mensaje solicitando la paz al lider de " & nombre
    Exit Sub
End If

Call SendData("PEACEOFF" & nombre & "," & Replace(Text1, vbCrLf, "º"))
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()

If Text1 = "" Then
    MsgBox "Debes redactar un mensaje solicitando la paz al lider de " & nombre
    Exit Sub
End If

Call SendData("PEACEOFF" & nombre & "," & Replace(Text1, vbCrLf, "º"))
Unload Me
End Sub
