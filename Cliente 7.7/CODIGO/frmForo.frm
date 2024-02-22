VERSION 5.00
Begin VB.Form frmForo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmForo.frx":0000
   ScaleHeight     =   7155
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox MiMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   5295
      Index           =   1
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox MiMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   345
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   5025
      Index           =   0
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmForo.frx":5FAB
      Top             =   1440
      Visible         =   0   'False
      Width           =   5025
   End
   Begin VB.ListBox List 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   5295
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   3960
      Picture         =   "frmForo.frx":5FB1
      Top             =   6600
      Width           =   1680
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   2280
      Picture         =   "frmForo.frx":A26A
      Top             =   6600
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   600
      Picture         =   "frmForo.frx":E3D2
      Top             =   6600
      Width           =   1650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmForo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ForoIndex As Integer
Private Sub Form_Load()
frmForo.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
End Sub






Private Sub Form_Deactivate()
Me.SetFocus
End Sub


Private Sub Image1_Click()
Dim i
For Each i In Text
    i.Visible = False
Next

If Not MiMensaje(0).Visible Then
    List.Visible = False
    MiMensaje(0).Visible = True
    MiMensaje(1).Visible = True
    MiMensaje(0).SetFocus
    'Command1.Enabled = False
    Label1.Visible = True
    Label2.Visible = True
Else
    Call SendData("DEMSG" & MiMensaje(0).Text & Chr(176) & MiMensaje(1).Text)
    List.AddItem MiMensaje(0).Text
    load Text(List.ListCount)
    Text(List.ListCount - 1).Text = MiMensaje(1).Text
    List.Visible = True
    
    MiMensaje(0).Visible = False
    MiMensaje(1).Visible = False
    'Command1.Enabled = True
    Label1.Visible = False
    Label2.Visible = False
End If
End Sub

Private Sub Image2_Click()
MiMensaje(0).Visible = False
MiMensaje(1).Visible = False
'Command1.Enabled = True
Label1.Visible = False
Label2.Visible = False
Dim i
For Each i In Text
    i.Visible = False
Next
List.Visible = True
End Sub

Private Sub Image3_Click()
Unload Me
'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus
End Sub

Private Sub List_Click()
List.Visible = False
Text(List.ListIndex).Visible = True

End Sub

Private Sub MiMensaje_Change(Index As Integer)
If Len(MiMensaje(0).Text) <> 0 And Len(MiMensaje(1).Text) <> 0 Then
'Image1.Enabled = True
End If

End Sub

