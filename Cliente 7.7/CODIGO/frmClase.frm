VERSION 5.00
Begin VB.Form frmClase 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   990
      Top             =   3300
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   4305
      Top             =   3285
      Width           =   960
   End
   Begin VB.Label lDESC 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1530
      Left            =   825
      TabIndex        =   0
      Top             =   1395
      Width           =   4635
   End
End
Attribute VB_Name = "frmClase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer
For i = LBound(ListaClases) To UBound(ListaClases)
 list1.AddItem ListaClases(i)
Next i
list1.Text = list1.List(LBound(ListaClases))
Me.Picture = LoadPicture(App.Path & "\Graficos\clase.jpg")
Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónVolver.jpg")
Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguiente.jpg")




End Sub

Private Sub Image1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)
Select Case Index
    Case 0
        UserClase = list1.Text
        Me.Visible = False
        frmCiudad.Show vbModal
    Case 1
        Unload Me
        frmSexo.Show vbModal
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
End Select
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
End Sub

