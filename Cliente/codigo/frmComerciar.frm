VERSION 5.00
Begin VB.Form frmComerciar 
   BorderStyle     =   0  'None
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   478
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3210
      TabIndex        =   6
      Text            =   "1"
      Top             =   5940
      Width           =   600
   End
   Begin VB.PictureBox PicInv 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   3600
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   1
      Top             =   1650
      Width           =   2970
   End
   Begin VB.PictureBox picInvNpc 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   360
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   0
      Top             =   1650
      Width           =   2970
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Image cComprar 
      Height          =   300
      Left            =   990
      Top             =   5925
      Width           =   1680
   End
   Begin VB.Image cVender 
      Height          =   300
      Left            =   4290
      Top             =   5925
      Width           =   1680
   End
   Begin VB.Image cCerrar 
      Height          =   300
      Left            =   2670
      Top             =   6570
      Width           =   1680
   End
   Begin VB.Label lblCantidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2160
      TabIndex        =   4
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label lblValor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1680
      TabIndex        =   3
      Top             =   915
      Width           =   120
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   420
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public LasActionBuy As Boolean
Private ClickNpcInv As Boolean
Private lIndex As Byte

Private Sub cantidad_Change()
'on error resume next
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If
    
    If ClickNpcInv Then
        If InvComNpc.SelectedItem <> 0 Then
            'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
            lblValor.Caption = PonerPuntos(CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text)))  'No mostramos numeros reales
        End If
    Else
        If InvComUsu.SelectedItem <> 0 Then
            lblValor.Caption = PonerPuntos(CalculateBuyPrice(Inventario.Valor(InvComUsu.SelectedItem), Val(cantidad.Text)))  'No mostramos numeros reales
        End If
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
'on error resume next
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cCerrar_Click()
    Call WriteCommerceEnd
    Unload Me
End Sub

Private Sub cCerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Sound.Sound_Play(SND_CLICK)
    cCerrar.Picture = General_Load_Picture_From_Resource("25.gif")
    cCerrar.Tag = "1"
End Sub

Private Sub cCerrar_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cCerrar.Tag = "0" Then
        cCerrar.Picture = General_Load_Picture_From_Resource("24.gif")
        cCerrar.Tag = "1"
    End If
End Sub

Private Sub cComprar_Click()
    
    ' Debe tener seleccionado un item para comprarlo.
    If InvComNpc.SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    
    Call Sound.Sound_Play(SND_CLICK)
    
    LasActionBuy = True
    If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text)) Then
        Call WriteCommerceBuy(InvComNpc.SelectedItem, Val(cantidad.Text))
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficiente oro.", 2, 51, 223, 1, 1)
        Exit Sub
    End If
End Sub

Private Sub cVender_Click()

    ' Debe tener seleccionado un item para comprarlo.
    If InvComUsu.SelectedItem = 0 Then Exit Sub

    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    
    Call Sound.Sound_Play(SND_CLICK)
    
    LasActionBuy = False

    Call WriteCommerceSell(InvComUsu.SelectedItem, Val(cantidad.Text))
End Sub

Private Sub cVender_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Sound.Sound_Play(SND_CLICK)
    cVender.Picture = General_Load_Picture_From_Resource("18.gif")
    cVender.Tag = "1"
End Sub

Private Sub cComprar_Mousedown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Call Sound.Sound_Play(SND_CLICK)
    cComprar.Picture = General_Load_Picture_From_Resource("19.gif")
    cComprar.Tag = "1"
End Sub

Private Sub cVender_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cVender.Tag = "0" Then
        cVender.Picture = General_Load_Picture_From_Resource("16.gif")
        cVender.Tag = "1"
    End If
End Sub

Private Sub cComprar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cComprar.Tag = "0" Then
        cComprar.Picture = General_Load_Picture_From_Resource("17.gif")
        cComprar.Tag = "1"
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cVender.Tag = "1" Then
        cVender.Picture = LoadPicture("")
        cVender.Tag = "0"
    End If
    
    If cComprar.Tag = "1" Then
        cComprar.Picture = LoadPicture("")
        cComprar.Tag = "0"
    End If
    
    If cCerrar.Tag = "1" Then
        cCerrar.Picture = LoadPicture("")
        cCerrar.Tag = "0"
    End If
End Sub

Private Sub Form_Activate()
'on error resume next
    InvComUsu.DrawInv
    InvComNpc.DrawInv
End Sub

Private Sub Form_GotFocus()
'on error resume next
    InvComUsu.DrawInv
    InvComNpc.DrawInv
End Sub

Private Sub Form_Load()
    Me.Picture = General_Load_Picture_From_Resource("15.gif")
    
    cVender.Picture = LoadPicture("")
    cVender.Tag = "0"
    cComprar.Picture = LoadPicture("")
    cComprar.Tag = "0"
    cCerrar.Picture = LoadPicture("")
    cCerrar.Tag = "0"
End Sub

''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Private Function CalculateSellPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * objAmount + 0.5
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function
''
' Calculates the buying price of an item (The price that a merchant will buy you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.
Private Function CalculateBuyPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * objAmount)
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'on error resume next
    InvComUsu.DrawInv
    InvComNpc.DrawInv
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Comerciando = False
End Sub

Private Sub picInvNpc_Click()
'on error resume next
    Dim ItemSlot As Byte
    
    Call Sound.Sound_Play(SND_CLICK)
    
    InvComNpc.DrawInv
    
    ItemSlot = InvComNpc.SelectedItem
    If ItemSlot = 0 Then Exit Sub
    
    ClickNpcInv = True
    InvComUsu.DeselectItem
    lblNombre.Caption = NPCInventory(ItemSlot).name
    lblValor.Caption = PonerPuntos(CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text))) 'No mostramos numeros reales
    
    If NPCInventory(ItemSlot).amount <> 0 Then
    
        Select Case NPCInventory(ItemSlot).OBJType
            Case eObjType.otWeapon
                lblInfo.Caption = "Máx Golpe:" & NPCInventory(ItemSlot).MaxHit & "/Mín Golpe:" & NPCInventory(ItemSlot).MinHit
                lblInfo.Visible = True
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                lblInfo.Caption = "Máx Defensa:" & NPCInventory(ItemSlot).MaxDef & "/Mín Defensa:" & NPCInventory(ItemSlot).MinDef
                lblInfo.Visible = True
            Case Else
                lblInfo.Visible = False
        End Select
    Else
        lblInfo.Visible = False
    End If
End Sub

Private Sub PicInv_Click()
'on error resume next
    Dim ItemSlot As Byte
    
    Call Sound.Sound_Play(SND_CLICK)
    
    InvComUsu.DrawInv
    
    ItemSlot = InvComUsu.SelectedItem
    
    If ItemSlot = 0 Then Exit Sub
    
    ClickNpcInv = False
    InvComNpc.DeselectItem
    
    lblNombre.Caption = Inventario.ItemName(ItemSlot)
    lblValor.Caption = PonerPuntos(CalculateBuyPrice(Inventario.Valor(ItemSlot), Val(cantidad.Text))) 'No mostramos numeros reales
    
    If Inventario.amount(ItemSlot) <> 0 Then
    
        Select Case Inventario.OBJType(ItemSlot)
            Case eObjType.otWeapon
                lblInfo.Caption = "Máx Golpe:" & Inventario.MaxHit(ItemSlot) & "/Mín Golpe:" & Inventario.MinHit(ItemSlot)
                lblInfo.Visible = True
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                lblInfo.Caption = "Máx Defensa:" & Inventario.MaxDef(ItemSlot) & "/Mín Defensa:" & Inventario.MinDef(ItemSlot)
                lblInfo.Visible = True
            Case Else
                lblInfo.Visible = False
        End Select
    Else
        lblInfo.Visible = False
    End If
End Sub
