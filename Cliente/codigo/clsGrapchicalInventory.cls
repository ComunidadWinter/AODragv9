VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item
Private RGB(3) As Long
Private RGB_A(3) As Long

Private LastRender As Long  ' GSZAO
Private FixY As Boolean ' GSZAO (cuando esta True, es para dibujar el inventario del jugador, sino es otro inventario, de comercio, etc)
Private ItemEmpty As Direct3DTexture8 ' GSZAO
Private rgbRed(3) As Long ' GSZAO

Private slotDragged As Integer
 
Public Event dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer

Private ShowValue As Boolean
Private AlphaItems As Byte

Private MouseX As Single
Private MouseY As Single

Private Inventory() As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
    'Redraw old item to deselect it
    If prevSelectedItem <> FLAGORO Then
        Call DrawInventory
    End If
End Sub

Public Sub DeselectItem()
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Deselect the currently selected item
'***************************************************
    If InvSelectedItem = 0 Then Exit Sub
    
    Dim ClearSlot As Byte
    
    ClearSlot = InvSelectedItem
    
    'Select nothing
    InvSelectedItem = 0
    
    'Redraw old item to deselect it
    Call DrawInventory
End Sub

Public Sub ChangeSlotItemAmount(ByVal slot As Byte, ByVal amount As Long)
'***************************************************
'Author: ZaMa
'Created Date: 30/11/2009
'Last Modify Date: 09/12/2009 (Budi)
'Change the amount of the required item, and update the slot.
'***************************************************
    
    Inventory(slot).amount = amount
    
    'Update Amount
    Call DrawInventory
End Sub

' Retrieves the selected item index
Public Property Get SelectedItem() As Byte: SelectedItem = InvSelectedItem: End Property
' Retrieves the max hit of the selected item
Public Property Get MaxHit(ByVal slot As Byte) As Integer: MaxHit = Inventory(slot).MaxHit: End Property
' Retrieves the min hit of the selected item
Public Property Get MinHit(ByVal slot As Byte) As Integer: MinHit = Inventory(slot).MinHit: End Property
' Retrieves the defense of the selected item
Public Property Get MaxDef(ByVal slot As Byte) As Integer: MaxDef = Inventory(slot).MaxDef: End Property
' Retrieves the defense of the selected item
Public Property Get MinDef(ByVal slot As Byte) As Integer: MinDef = Inventory(slot).MinDef: End Property
' Retrieves the grh index of the selected item
Public Property Get GrhIndex(ByVal slot As Byte) As Integer: GrhIndex = Inventory(slot).GrhIndex: End Property
' Retrieves the value of the selected item
Public Property Get Valor(ByVal slot As Byte) As Long: Valor = Inventory(slot).Valor: End Property
' Retrieves the selected item's object index
Public Property Get OBJIndex(ByVal slot As Byte) As Integer: OBJIndex = Inventory(slot).OBJIndex: End Property
' Retrieves the selected item's object type
Public Property Get OBJType(ByVal slot As Byte) As Integer: OBJType = Inventory(slot).OBJType: End Property
' Retrieves the selected item's name
Public Property Get ItemName(ByVal slot As Byte) As String
    If slot > 0 Then
      ItemName = Inventory(slot).name
     End If
End Property
' Retrieves True if the item at the given pos is eqiupped
Public Property Get Equipped(ByVal slot As Byte) As Boolean: Equipped = Inventory(slot).Equipped: End Property
'Puede usar
Public Property Get PuedeUsar(ByVal slot As Byte) As Byte: PuedeUsar = Inventory(slot).PuedeUsar: End Property

Public Property Get amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        amount = UserGLD
    ElseIf slot >= LBound(Inventory) And slot <= UBound(Inventory) Then
        amount = Inventory(slot).amount
    End If
End Property

Public Property Get MaxObjs() As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Retrieves the capacity of the Inventory
'***************************************************
    MaxObjs = UBound(Inventory)
End Property

Public Sub ToggleShowValue(ByVal bShow As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 11/12
'Indicates if the amount of the items in the inventory is shown
'***************************************************
    ShowValue = bShow
    
    Dim slot As Integer
    ' Update Inventory
    For slot = 1 To Me.MaxObjs
        DrawInventory
    Next slot
    
End Sub

Public Sub ClearAllSlots() ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modify Date: 23/11/2011 - ^[GS]^
'Clear All Slots.
'***************************************************
'on error resume next

    Dim slot As Long
    For slot = 1 To Me.MaxObjs
        Call SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0)
    Next slot

End Sub

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Long, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eMaxDef As Integer, ByVal eMinDef As Integer, ByVal eValor As Single, ByVal eName As String, ByVal ePuedeUsar As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 10/08/2014 - ^[GS]^
'Sets all data for a given inventory slot
'***************************************************
On Error GoTo Fallo:

    If slot < 1 Or slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(slot)
        .amount = eAmount
        .MaxDef = eMaxDef
        .MinDef = eMinDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
        .PuedeUsar = ePuedeUsar
    End With
    
    'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
    
    'Render inventory slot (we don't need to render the whole inventory)
    Call DrawInventory
    
    Exit Sub
Fallo:
    
    Call LogError("SetItem::Error " & Err.number & " - " & Err.Description & " - Slot: " & slot & " - OBJIndex: " & eOBJIndex)

End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 08/08/2012 - ^[GS]^
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    If Comerciando = False Then
        TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 - temp_y
    Else
        TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    End If
    
    If (temp_y > 0) And FixY = True Then
        TempItem = TempItem + (temp_y) ' Corrección de filas
    End If
    
    'Make sure it's within limits
    If TempItem <= UBound(Inventory) Then
        'Make sure slot isn't empty
        If Inventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
    
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        ClickItem = TempItem
    End If
    
End Function

Private Sub DrawInventory()
'on error resume next
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 02/05/2013 - ^[GS]^
'Renders the inventory to the given PictureBox
'***************************************************
    Dim loopc As Long
    Dim SrcRect As RECT
    Dim SrcRecti As RECT
    Dim TempRect As RECT
    Dim DestRect As RECT
    
    With TempRect
        .Bottom = InventoryWindow.Height
        .Right = InventoryWindow.Width
    End With
    
    DestRect = TempRect
       
    'Clear the inventory window
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    DirectDevice.BeginScene
    
    For loopc = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)

            startdX = ((loopc - 1) Mod (InventoryWindow.Width / 32)) * 32
            startdY = ((loopc - 1) \ (InventoryWindow.Width / 32)) * 32
            
            If Inventory(loopc).GrhIndex Then
                    
                'Render a box around the selected item
                If InvSelectedItem = loopc Then
                    Call DDrawTransGrhIndextoSurface(2, startdX, startdY, 0, LightIluminado())
                End If
                                                                        
                'Dibujamos el Grh del item
                If Inventory(loopc).PuedeUsar = 1 Then
                    Call DDrawTransGrhIndextoSurface(Inventory(loopc).GrhIndex, startdX, startdY, 0, LightIluminado())
                Else
                    Call DDrawTransGrhIndextoSurface(Inventory(loopc).GrhIndex, startdX, startdY, 0, NoPuedeUsar)
                End If
                        
                'Si esta equipado le ponemos una marca para indicarlo
                If Inventory(loopc).Equipped Then
                    Call texto.Engine_Text_Draw(startdX + 21, startdY, "+", -65536)
                End If
                                        
                'Dibujamos la cantidad de ese item
                Call texto.Engine_Text_Draw(startdX, startdY + 18, Inventory(loopc).amount, -1)
            End If
        Next loopc
    DirectDevice.EndScene
    DirectDevice.Present DestRect, ByVal 0, InventoryWindow.hwnd, ByVal 0
    
End Sub

Public Sub DrawInv()

    If GetTickCount - LastRender > 56 Then ' GSZAO
        LastRender = GetTickCount
        DoEvents
        'aca
        Call DrawInventory
    End If
    
End Sub

Public Sub Initialize(ByRef DirectD3D As D3DX8, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, _
                            Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 32, _
                            Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, _
                            Optional ByVal startY As Integer = 0, Optional ByVal bShowText As Boolean = True, _
                            Optional ByVal bFixY As Boolean = False)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 03/12/09
'Sets the reference to the picture box on which to render the inventory
'03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
'***************************************************
    Set InventoryWindow = InvPic
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    'initialize DX stuff
    'Dim SurfaceDesc As DDSURFACEDESC2
    
    'Make sure DirectD3D was correctly initialized
    If DirectD3D Is Nothing Then Exit Sub

    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
    FixY = bFixY ' GSZAO
    ShowValue = bShowText
        
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 30/06/2012 - ^[GS]^
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
        
    MouseX = X
    MouseY = Y
        
    If Button = vbRightButton Then
        
        'Call DrawInventory
        
    Else

        temp_x = X \ ownTilePixelWidth
        temp_y = Y \ ownTilePixelHeight
        
        If Comerciando = False Then
            TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 - temp_y
        Else
            TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
        End If
        
        If (temp_y > 0) And FixY = True Then
            TempItem = TempItem + (temp_y) ' Corrección de filas
        End If
        
        If TempItem <= UBound(Inventory) Then
            If FixY = True Then ' inventario del usuario
                frmMain.lblItem.Caption = Inventory(TempItem).name ' Que nos diga el nombre del objeto cuando pasamos el mouse sobre él
            Else ' "otros" inventarios
                InventoryWindow.ToolTipText = Inventory(TempItem).name ' Que nos diga el nombre del objeto cuando pasamos el mouse sobre él
            End If
        End If
    
    End If
    
    'Call DrawInventory
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
   
    'Exit if it got outside the control's area
    If X < 0 Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        X = 0
    ElseIf Y < 0 Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        Y = 0
    ElseIf X > InventoryWindow.Width Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        X = InventoryWindow.Width
    ElseIf Y > InventoryWindow.Height Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        Y = InventoryWindow.Height
    End If
   
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
   
    prevSelItem = InvSelectedItem
   
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
       
    If Not frmMain.UsabaDrag Then InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    'aca
    'Call DrawInventory
        
    If Button = vbRightButton Then
        If InvSelectedItem <> slotDragged Then
            'RaiseEvent dragDone(slotDragged, InvSelectedItem)
            Call frmMain.picInv_dragDone(slotDragged, InvSelectedItem)
            frmMain.UsandoDrag = False
        Else
            frmMain.UsandoDrag = False
        End If
    End If
        

End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer

    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
       
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
       
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
       
    If Button = vbRightButton Then
        If Comerciando = False Then
            TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 - temp_y
        Else
            TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
        End If
        
        If (temp_y > 0) And FixY = True Then
            TempItem = TempItem + (temp_y) ' Corrección de filas
        End If
        
        slotDragged = TempItem
    End If

End Sub


