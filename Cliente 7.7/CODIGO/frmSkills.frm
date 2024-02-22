VERSION 5.00
Begin VB.Form frmSkills 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   4
      Left            =   3405
      TabIndex        =   4
      Top             =   3420
      Width           =   480
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   3
      Left            =   3405
      TabIndex        =   3
      Top             =   2910
      Width           =   480
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   2
      Left            =   3405
      TabIndex        =   2
      Top             =   2310
      Width           =   480
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   1
      Left            =   3405
      TabIndex        =   1
      Top             =   1725
      Width           =   480
   End
   Begin VB.Label text1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   3405
      TabIndex        =   0
      Top             =   1125
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   2
      Left            =   2310
      Top             =   3975
      Width           =   1830
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   465
      Top             =   4485
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   5010
      Top             =   4455
      Width           =   960
   End
End
Attribute VB_Name = "frmSkills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Graficos\skills.jpg")
Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónVolver.jpg")
Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguiente.jpg")
Image1(2).Picture = LoadPicture(App.Path & "\Graficos\TirarDados.jpg")
Dim j
For Each j In Image1()
    j.Tag = "0"
Next

Call TirarDados
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = "1" Then
            Image1(0).Tag = "0"
            Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguiente.jpg")
End If
If Image1(1).Tag = "1" Then
            Image1(1).Tag = "0"
            Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónVolver.jpg")
End If
If Image1(2).Tag = "1" Then
            Image1(2).Tag = "0"
            Image1(2).Picture = LoadPicture(App.Path & "\Graficos\TirarDados.jpg")
End If
End Sub

Private Sub Image1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)
Select Case Index
    Case 0
    Case 1
        Unload Me
        frmCiudad.Show vbModal
    Case 2
        Call PlayWaveDS(SND_DICE)
        Call TirarDados
End Select
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        If Image1(0).Tag = "0" Then
            Image1(0).Tag = "1"
            Call PlayWaveDS(SND_OVER)
            Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguienteApretado.jpg")
        End If
    Case 1
        If Image1(1).Tag = "0" Then
            Image1(1).Tag = "1"
            Call PlayWaveDS(SND_OVER)
            Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónVolverApretado.jpg")
        End If
    Case 2
        If Image1(2).Tag = "0" Then
            Image1(2).Tag = "1"
            Call PlayWaveDS(SND_OVER)
            Image1(2).Picture = LoadPicture(App.Path & "\Graficos\TirarDadosApretado.jpg")
        End If
End Select
End Sub
