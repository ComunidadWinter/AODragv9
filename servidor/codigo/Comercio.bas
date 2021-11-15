Attribute VB_Name = "modSistemaComercio"
'*****************************************************
'Sistema de Comercio para Argentum Online
'Programado por Nacho (Integer)
'integer-x@hotmail.com
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
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Enum eModoComercio
    Compra = 1
    Venta = 2
End Enum

Public Const REDUCTOR_PRECIOVENTA As Byte = 3

''
' Makes a trade. (Buy or Sell)
'
' @param Modo The trade type (sell or buy)
' @param UserIndex Specifies the index of the user
' @param NpcIndex specifies the index of the npc
' @param Slot Specifies which slot are you trying to sell / buy
' @param Cantidad Specifies how many items in that slot are you trying to sell / buy
Public Sub Comercio(ByVal Modo As eModoComercio, ByVal UserIndex As Integer, ByVal NPCIndex As Integer, ByVal Slot As Integer, ByVal Cantidad As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
'  - 06/13/08 (NicoNZ)
'*************************************************
    Dim Precio As Long
    Dim Objeto As Obj
    
    If Cantidad < 1 Or Slot < 1 Then Exit Sub
    
    If Modo = eModoComercio.Compra Then
        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Sub
        ElseIf Cantidad > MAX_INVENTORY_OBJS Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
            Call Ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados items:" & Cantidad)
            UserList(UserIndex).flags.Ban = 1
            Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        ElseIf Not NPCList(NPCIndex).Invent.Object(Slot).Amount > 0 Then
            Exit Sub
        End If
        
        If Cantidad > NPCList(NPCIndex).Invent.Object(Slot).Amount Then Cantidad = NPCList(UserList(UserIndex).flags.TargetNPC).Invent.Object(Slot).Amount
        
        Objeto.Amount = Cantidad
        Objeto.ObjIndex = NPCList(NPCIndex).Invent.Object(Slot).ObjIndex
        
        'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
        'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
        
        Precio = CLng((ObjData(NPCList(NPCIndex).Invent.Object(Slot).ObjIndex).Valor) * Cantidad) + 0.5

        If UserList(UserIndex).Stats.GLD < Precio Then
           Call WriteMultiMessage(UserIndex, eMessages.NoCantidad)
            Exit Sub
        End If
        
        
        If MeterItemEnInventario(UserIndex, Objeto) = False Then
            'Call WriteConsoleMsg(UserIndex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        End If
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
        
        Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(Slot), Cantidad)
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).Name & " compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
        ElseIf Objeto.Amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).Name & " compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
            End If
        End If
        
        'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
        If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
            Call WriteVar(DatPath & "NPCs.dat", "NPC" & NPCList(NPCIndex).Numero, "obj" & Slot, Objeto.ObjIndex & "-0")
            Call logVentaCasa(UserList(UserIndex).Name & " compro " & ObjData(Objeto.ObjIndex).Name)
        End If
        
    ElseIf Modo = eModoComercio.Venta Then
        
        If Cantidad > UserList(UserIndex).Invent.Object(Slot).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Slot).Amount
        
        Objeto.Amount = Cantidad
        Objeto.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        If Objeto.ObjIndex = 0 Then
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Newbie = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.NoComerciar)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        ElseIf (NPCList(NPCIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And NPCList(NPCIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
            Call WriteMultiMessage(UserIndex, eMessages.NoComerciar)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).NoComerciable = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.NoComerciar)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        ElseIf UserList(UserIndex).Invent.Object(Slot).Amount < 0 Or Cantidad = 0 Then
            Exit Sub
        ElseIf Slot < LBound(UserList(UserIndex).Invent.Object()) Or Slot > UBound(UserList(UserIndex).Invent.Object()) Then
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        End If
        
        Call QuitarUserInvItem(UserIndex, Slot, Cantidad)
        
        'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
        Precio = Fix(SalePrice(ObjData(Objeto.ObjIndex).Valor) * Cantidad)
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Precio
        
        If UserList(UserIndex).Stats.GLD > MAXORO Then _
            UserList(UserIndex).Stats.GLD = MAXORO
        
        Dim NpcSlot As Integer
        NpcSlot = SlotEnNPCInv(NPCIndex, Objeto.ObjIndex, Objeto.Amount)
        
        If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
            'Mete el obj en el slot
            NPCList(NPCIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
            NPCList(NPCIndex).Invent.Object(NpcSlot).Amount = NPCList(NPCIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount
            If NPCList(NPCIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
                NPCList(NPCIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS
            End If
        End If
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).Name & " vendió al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
        ElseIf Objeto.Amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).Name & " vendió al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
            End If
        End If
        
    End If
    
    Call UpdateUserInv(True, UserIndex, 0)
    Call WriteUpdateUserStats(UserIndex)
    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
    Call WriteTradeOK(UserIndex)
End Sub

Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
        Call WriteTradeOK(UserIndex)
    UserList(UserIndex).flags.Comerciando = True
    Call WriteCommerceInit(UserIndex)
End Sub

Private Function SlotEnNPCInv(ByVal NPCIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    SlotEnNPCInv = 1
    Do Until NPCList(NPCIndex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto _
      And NPCList(NPCIndex).Invent.Object(SlotEnNPCInv).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
        SlotEnNPCInv = SlotEnNPCInv + 1
        If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
        
    Loop
    
    If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then
    
        SlotEnNPCInv = 1
        
        Do Until NPCList(NPCIndex).Invent.Object(SlotEnNPCInv).ObjIndex = 0
        
            SlotEnNPCInv = SlotEnNPCInv + 1
            If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
            
        Loop
        
        If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then NPCList(NPCIndex).Invent.NroItems = NPCList(NPCIndex).Invent.NroItems + 1
    
    End If
    
End Function

''
' Send the inventory of the Npc to the user
'
' @param userIndex The index of the User
' @param npcIndex The index of the NPC

Private Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last Modified: 06/14/08
'Last Modified By: Nicolás Ezequiel Bouhid (NicoNZ)
'*************************************************
    Dim Slot As Byte
    Dim val As Single
    
    For Slot = 1 To MAX_INVENTORY_SLOTS
        If NPCList(NPCIndex).Invent.Object(Slot).ObjIndex > 0 Then
            Dim thisObj As Obj
            thisObj.ObjIndex = NPCList(NPCIndex).Invent.Object(Slot).ObjIndex
            thisObj.Amount = NPCList(NPCIndex).Invent.Object(Slot).Amount
            val = (ObjData(NPCList(NPCIndex).Invent.Object(Slot).ObjIndex).Valor)
            
            Call WriteChangeNPCInventorySlot(UserIndex, Slot, thisObj, val)
        Else
            Dim DummyObj As Obj
            Call WriteChangeNPCInventorySlot(UserIndex, Slot, DummyObj, 0)
        End If
    Next Slot
End Sub

''
' Devuelve el valor de venta del objeto
'
' @param valor  El valor de compra de objeto

Public Function SalePrice(ByVal Valor As Long) As Single
'*************************************************
'Author: Nicolás (NicoNZ)
'
'*************************************************
    SalePrice = Valor / REDUCTOR_PRECIOVENTA
End Function
