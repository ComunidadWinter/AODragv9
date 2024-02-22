VERSION 5.00
Begin VB.Form frmMiniMap 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   240
   ClientTop       =   6705
   ClientWidth     =   1500
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   1500
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MiniMap 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      Begin VB.Timer tMinimap 
         Interval        =   800
         Left            =   120
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmMiniMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Constantes
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Public Sub Formulario(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub
Public Function Transparencia(ByVal hWnd As Long, valor As Integer) As Long

On Local Error GoTo ErrSub

Dim Estilo As Long

If valor < 0 Or valor > 255 Then
Transparencia = 1
Else

Estilo = GetWindowLong(hWnd, GWL_EXSTYLE)
Estilo = Estilo Or WS_EX_LAYERED

SetWindowLong hWnd, GWL_EXSTYLE, Estilo

'Aplica el nuevo estilo con la transparencia
SetLayeredWindowAttributes hWnd, 0, valor, LWA_ALPHA

Transparencia = 0
End If

If Err Then
Transparencia = 2
End If

Exit Function

'Error
ErrSub:

MsgBox Err.Description, vbCritical, "Error"

End Function

Private Sub Redondear_Formulario(El_Form As Form, Radio As Long)
  
Dim Region As Long
Dim Ret As Long
Dim ancho As Long
Dim alto As Long
Dim old_Scale As Integer
       
    ' guardar la escala
    old_Scale = El_Form.ScaleMode
       
    ' cambiar la escala a pixeles
    El_Form.ScaleMode = vbPixels
       
    'Obtenemos el ancho y alto de la region del Form
    ancho = El_Form.ScaleWidth
    alto = El_Form.ScaleHeight
  
    'Pasar el ancho alto del formualrio y el valor de redondeo .. es decir el radio
    Region = CreateRoundRectRgn(0, 0, ancho, alto, Radio, Radio)
  
    ' Aplica la región al formulario
    Ret = SetWindowRgn(El_Form.hWnd, Region, True)
       
    ' restaurar la escala
    El_Form.ScaleMode = old_Scale
  
End Sub
  
Private Sub Form_Load()
Dim i As Integer
    ' Le pasamos el formulario y el radio de redondeo
    Call Redondear_Formulario(Me, 10)
If Not Transparencia(Me.hWnd, 0) = 0 Then
    MsgBox " El Minimapa no es visible perfectamente en Sistemas Operativos" _
    & "anteriores a windows 2000", vbCritical
    Me.Show
Else
    Me.Enabled = False
    Me.Show

Call Transparencia(Me.hWnd, 190)

'reactiva la ventana
    Me.Enabled = True

End If
 'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus
End Sub


Private Sub tMinimap_Timer()
Mod_TileEngine.DibujarMiniMapa MiniMap 'minimap
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Formulario Me
End Sub
Private Sub MiniMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Formulario Me
'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus
End Sub



