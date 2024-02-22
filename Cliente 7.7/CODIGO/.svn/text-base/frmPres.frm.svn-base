VERSION 5.00
Begin VB.Form frmPres 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmPres.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   Picture         =   "frmPres.frx":0CCA
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1125
      Top             =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   2175
      Left            =   2520
      TabIndex        =   0
      Top             =   5640
      Width           =   7095
   End
End
Attribute VB_Name = "frmPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_KeyPress(KeyAscii As Integer)
'Me.Picture = LoadPicture(App.Path & "\Graficos\into.jpg")

'If KeyAscii = 27 Then finpres = True
End Sub

Private Sub Timer1_Timer()
Static ticks As Long

ticks = ticks + 1

If ticks = 1 Then
Call audio.PlayWave("aodrag.wav")
'frmPres.Navegador.Visible = False
    'ElseIf ticks = 2 Then
    'Me.Picture = LoadPicture(App.Path & "\Graficos\datafull.jpg")
ElseIf ticks = 2 Then
'frmPres.Navegador.Visible = True




   ' Me.Picture = LoadPicture(App.Path & "\Graficos\ls.jpg")
ElseIf ticks = 4 Then
 finpres = True
End If

End Sub
