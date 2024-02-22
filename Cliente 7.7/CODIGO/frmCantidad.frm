VERSION 5.00
Begin VB.Form frmCantidad 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   770
      TabIndex        =   0
      Top             =   1475
      Width           =   2205
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   960
      MouseIcon       =   "frmCantidad.frx":0000
      MousePointer    =   99  'Custom
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   960
      MouseIcon       =   "frmCantidad.frx":0CCA
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmCantidad.Picture = LoadPicture(DirGraficos & "cantidadflash.jpg")
End Sub




Private Sub Form_Deactivate()
Unload Me
'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus
End Sub

Private Sub Image1_Click()
Unload Me
Exit Sub
'frmCantidad.Visible = False
'SendData "TI" & ItemElegido & "," & frmCantidad.Text1.Text
'frmCantidad.Text1.Text = "0"

'pluto:2.4.5
frmCantidad.Visible = False
If Not NoPuedeTirar Then
 NoPuedeTirar = True
 SendData "TI" & ItemElegido & "," & frmCantidad.Text1.Text
frmCantidad.Text1.Text = "0"
End If




End Sub

Private Sub Image2_Click()
frmCantidad.Visible = False
If Not NoPuedeTirar Then
 NoPuedeTirar = True
 SendData "TI" & ItemElegido & "," & frmCantidad.Text1.Text
frmCantidad.Text1.Text = "0"
End If

'frmCantidad.Visible = False
'If Not NoPuedeTirar Then
' NoPuedeTirar = True
'If ItemElegido <> FLAGORO Then
'SendData "TI" & ItemElegido & "," & UserInventory(ItemElegido).Amount
'Else
'SendData "TI" & ItemElegido & "," & UserGLD
'End If
'End If
'frmCantidad.Text1.Text = "0"

'frmCantidad.Visible = False
'If ItemElegido <> FLAGORO Then
   ' SendData "TI" & ItemElegido & "," & UserInventory(ItemElegido).Amount
'Else
  '  SendData "TI" & ItemElegido & "," & UserGLD
'End If

'frmCantidad.Text1.Text = "0"

End Sub

Private Sub Text1_Change()

If Val(Text1.Text) < 0 Then
    Text1.Text = MAX_INVENTORY_OBJS
End If

If Val(Text1.Text) > MAX_INVENTORY_OBJS And ItemElegido <> FLAGORO Then
    Text1.Text = 1
End If

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (Index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub
