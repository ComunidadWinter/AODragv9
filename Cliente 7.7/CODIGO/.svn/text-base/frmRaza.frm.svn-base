VERSION 5.00
Begin VB.Form frmRaza 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox clRaza 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmRaza.frx":0000
      Left            =   2385
      List            =   "frmRaza.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   645
      Width           =   3045
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   4305
      Top             =   3300
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   990
      Top             =   3315
      Width           =   960
   End
   Begin VB.Label lDESC 
      BackStyle       =   0  'Transparent
      Caption         =   "Raza"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1530
      Left            =   810
      TabIndex        =   1
      Top             =   1395
      Width           =   4710
   End
End
Attribute VB_Name = "frmRaza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clRaza_Click()
Select Case clRaza.Text
    Case "Humano"
      lDESC.Caption = "Los Humanos son la raza mas comun en Argentum. Son muy buenos Comerciantes y Guerreros."
    Case "Elfo"
      lDESC.Caption = "Los Elfos pueblan los bosques de Argentum. Son exelentes ladrones y cazadores"
    Case "Elfo Oscuro"
      lDESC.Caption = "En general son malvados, son muy odiados a lo largo de todo Argentum. Son muy buenos con las armas."
    Case "Gnomo"
      lDESC.Caption = "Son una raza extraña, tienen mucha suerte y ciertos gnomos tiene propiedades magicas"
    Case "Enano"
      lDESC.Caption = "Son muy bueno luchadores. Habitan las zonas montañosas. Existe una gran enemistad entre gnomos y enanos"
End Select
End Sub



Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Graficos\raza.jpg")
Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónVolver.jpg")
Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguiente.jpg")

Dim j
For Each j In Image1()
    j.Tag = "0"
Next


lDESC.WordWrap = True
Dim i As Integer
For i = LBound(ListaRazas) To UBound(ListaRazas)
 clRaza.AddItem ListaRazas(i)
Next i
clRaza.Text = clRaza.List(LBound(ListaRazas))
Call clRaza_Click
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
        UserRaza = clRaza.Text
        Me.Visible = False
        frmSexo.Show vbModal
    Case 1
        Unload Me
        frmCrearCaracter.Show vbModal
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

