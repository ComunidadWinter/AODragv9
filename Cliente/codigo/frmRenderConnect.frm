VERSION 5.00
Begin VB.Form frmRenderConnect 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   Icon            =   "frmRenderConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image btnConsejo 
      Height          =   255
      Left            =   6840
      Top             =   9480
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmRenderConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnConsejo_Click()
    Call Audio.General_Set_Wav(SND_CLICK)
    ConsejoSeleccionado = Consejos(RandomNumber(1, UBound(ListaConsejos())))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Call Audio.General_Set_Wav(SND_CLICK)
        Call CloseClient
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub Form_Load()
    Me.Caption = Form_Caption
End Sub
