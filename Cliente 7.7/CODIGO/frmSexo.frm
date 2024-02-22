VERSION 5.00
Begin VB.Form frmSexo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox clist1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      ItemData        =   "frmSexo.frx":0000
      Left            =   1695
      List            =   "frmSexo.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   330
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   375
      Top             =   840
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   3480
      Top             =   855
      Width           =   780
   End
End
Attribute VB_Name = "frmSexo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
clist1.AddItem "Hombre"
clist1.AddItem "Mujer"
clist1.Text = clist1.List(0)
Me.Picture = LoadPicture(App.Path & "\Graficos\sexo.jpg")
Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónVolver.jpg")
Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguiente.jpg")
Dim j
For Each j In Image1()
    j.Tag = "0"
Next

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = "1" Then
            Image1(0).Tag = "0"
            Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguiente.jpg")
End If
If Image1(1).Tag = "1" Then
            Image1(1).Tag = "0"
            Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónVolver.jpg")
End If
End Sub

Private Sub Image1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)
Select Case Index
    Case 0
        UserSexo = clist1.Text
        Me.Visible = False
        frmClase.Show vbModal
    Case 1
        Unload Me
        frmRaza.Show vbModal
End Select
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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
End Select
End Sub

