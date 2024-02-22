Attribute VB_Name = "modBanco"
Option Explicit

'MODULO PROGRAMADO POR NEB
'Kevin Birmingham
'kbneb@hotmail.com

Sub IniciarDeposito(ByVal UserIndex As Integer)
On Error GoTo fallo

'Hacemos un Update del inventario del usuario
'Pluto:7.0 añado caja
Call UpdateBanUserInv(True, UserIndex, 0)
'Atcualizamos el dinero
Call SendUserStatsOro(UserIndex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
SendData2 ToIndex, UserIndex, 0, 11
UserList(UserIndex).flags.Comerciando = True

Exit Sub
fallo:
Call LogError("iniciardeposito " & Err.number & " D: " & Err.Description)


End Sub

Sub SendBanObj(UserIndex As Integer, Slot As Byte, Object As UserOBJ)

On Error GoTo fallo
Dim Caja As Byte
Caja = UserList(UserIndex).flags.NCaja
UserList(UserIndex).BancoInvent(Caja).Object(Slot) = Object
'pluto:6.0A
If Object.ObjIndex > 0 Then
    Call SendData2(ToIndex, UserIndex, 0, 33, Slot & "," & Object.ObjIndex & "," & Object.Amount)
Else
    Call SendData2(ToIndex, UserIndex, 0, 33, Slot & "," & "0") ' & "," & "(None)" & "," & "0" & "," & "0")
End If

Exit Sub
fallo:
Call LogError("senbanobj " & Err.number & " D: " & Err.Description)

End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo fallo
Dim NullObj As UserOBJ
Dim loopc As Byte
Dim Caja As Byte
Caja = UserList(UserIndex).flags.NCaja
'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).BancoInvent(Caja).Object(Slot).ObjIndex > 0 Then
        Call SendBanObj(UserIndex, Slot, UserList(UserIndex).BancoInvent(Caja).Object(Slot))
    Else
        Call SendBanObj(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
'pluto:7.0
    For loopc = 1 To MAX_BANCOINVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(UserIndex).BancoInvent(Caja).Object(loopc).ObjIndex > 0 Then
            Call SendBanObj(UserIndex, loopc, UserList(UserIndex).BancoInvent(Caja).Object(loopc))
        Else
            
            Call SendBanObj(UserIndex, loopc, NullObj)
            
        End If

    Next loopc

End If
Exit Sub
fallo:
Call LogError("Updatebanuserinv " & Err.number & " D: " & Err.Description)

End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
On Error GoTo fallo


If Cantidad < 1 Then Exit Sub
Dim Caja As Byte
Caja = UserList(UserIndex).flags.NCaja
Call SendUserStatsOro(UserIndex)
   
       If UserList(UserIndex).BancoInvent(Caja).Object(i).Amount > 0 Then
            If Cantidad > UserList(UserIndex).BancoInvent(Caja).Object(i).Amount Then Cantidad = UserList(UserIndex).BancoInvent(Caja).Object(i).Amount
            'Agregamos el obj que compro al inventario
            Call UserReciveObj(UserIndex, CInt(i), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el banco
            Call UpdateBanUserInv(True, UserIndex, 0)
            'Actualizamos la ventana de comercio
            Call UpdateVentanaBanco(i, 0, UserIndex)
       End If

Exit Sub
fallo:
Call LogError("userretiraitem " & Err.number & " D: " & Err.Description)


End Sub

Sub UserReciveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
On Error GoTo fallo
Dim Slot As Integer
Dim obji As Integer

'pluto:2.15
'If UserList(UserIndex).flags.TargetNpcTipo = 25 Then Exit Sub
Dim Caja As Byte
Caja = UserList(UserIndex).flags.NCaja

If UserList(UserIndex).BancoInvent(Caja).Object(ObjIndex).Amount <= 0 Then Exit Sub



obji = UserList(UserIndex).BancoInvent(Caja).Object(ObjIndex).ObjIndex


'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji And _
   UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
    Slot = Slot + 1
    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

'Sino se fija por un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1
            If Slot > MAX_INVENTORY_SLOTS Then
            Call SendData(ToIndex, UserIndex, 0, "P7")
            Exit Sub
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If



'Mete el obj en el slot
If UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
    'Menor que MAX_INV_OBJS
    UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad
'pluto:2.4.5
UserList(UserIndex).Stats.Peso = UserList(UserIndex).Stats.Peso + (ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Peso * Cantidad)
Call SendUserStatsPeso(UserIndex)
    
    Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), Cantidad)
Else
    Call SendData(ToIndex, UserIndex, 0, "P7")
    End If

Exit Sub
fallo:
Call LogError("Userrecibeobj " & Err.number & " D: " & Err.Description)

End Sub

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)


On Error GoTo fallo
Dim Caja As Byte
Caja = UserList(UserIndex).flags.NCaja
Dim ObjIndex As Integer
ObjIndex = UserList(UserIndex).BancoInvent(Caja).Object(Slot).ObjIndex

    'Quita un Obj

       UserList(UserIndex).BancoInvent(Caja).Object(Slot).Amount = UserList(UserIndex).BancoInvent(Caja).Object(Slot).Amount - Cantidad
        
        If UserList(UserIndex).BancoInvent(Caja).Object(Slot).Amount <= 0 Then
            'UserList(UserIndex).BancoInvent(Caja).NroItems = UserList(UserIndex).BancoInvent.NroItems - 1
            UserList(UserIndex).BancoInvent(Caja).Object(Slot).ObjIndex = 0
            UserList(UserIndex).BancoInvent(Caja).Object(Slot).Amount = 0
        End If

    
  Exit Sub
fallo:
Call LogError("quitarbancoinvitem " & Err.number & " D: " & Err.Description)
  
End Sub

Sub UpdateVentanaBanco(ByVal Slot As Integer, ByVal NpcInv As Byte, ByVal UserIndex As Integer)
   On Error GoTo fallo
   Call SendData2(ToIndex, UserIndex, 0, 71, Slot & "," & NpcInv)
Exit Sub
fallo:
Call LogError("updateventanabanco " & Err.number & " D: " & Err.Description)
End Sub


Sub UserDepositaItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

On Error GoTo fallo

'El usuario deposita un item
Call SendUserStatsOro(UserIndex)
'pluto:2.3
If ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).OBJType = 60 And UserList(UserIndex).flags.TargetNpcTipo = 4 Then
UserList(UserIndex).flags.Comerciando = False
Call SendData2(ToIndex, UserIndex, 0, 9)
Call SendData(ToIndex, UserIndex, 0, "||No puedes dejar Mascotas en la Bóveda." & "´" & FontTypeNames.FONTTYPE_WARNING)
Exit Sub
End If

'pluto:6.3
If ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).OBJType = 42 And UserList(UserIndex).flags.Montura > 0 Then
UserList(UserIndex).flags.Comerciando = False
Call SendData2(ToIndex, UserIndex, 0, 9)
Call SendData(ToIndex, UserIndex, 0, "||No puedes dejar la Ropa mientras cabalgas." & "´" & FontTypeNames.FONTTYPE_WARNING)
Exit Sub
End If


If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then
            
            If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
            'Agregamos el obj que compro al inventario
            Call UserDejaObj(UserIndex, CInt(Item), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el inventario del banco
            Call UpdateBanUserInv(True, UserIndex, 0)
            'Actualizamos la ventana del banco
            
            Call UpdateVentanaBanco(Item, 1, UserIndex)
            
End If

Exit Sub
fallo:
Call LogError("USErdepositaitem UI:" & UserIndex & " D: " & Err.Description & " Item: " & Item & " Can: " & Cantidad)

End Sub

Sub UserDejaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
On Error GoTo fallo
Dim Slot As Integer
Dim obji As Integer

If Cantidad < 1 Then Exit Sub
Dim Caja As Byte
Caja = UserList(UserIndex).flags.NCaja

obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex

'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(UserIndex).BancoInvent(Caja).Object(Slot).ObjIndex = obji And _
         UserList(UserIndex).BancoInvent(Caja).Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            'pluto:7.0
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
Loop

'Sino se fija por un slot vacio antes del slot devuelto
'pluto:7.0
If Slot > MAX_BANCOINVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).BancoInvent(Caja).Object(Slot).ObjIndex = 0
            Slot = Slot + 1
            'pluto:7.0
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "||No tienes mas espacio en el banco!!" & "´" & FontTypeNames.FONTTYPE_info)
                Exit Sub
                Exit Do
            End If
        Loop
        'If Slot <= MAX_BANCOINVENTORY_SLOTS Then UserList(UserIndex).BancoInvent(Caja).NroItems = UserList(UserIndex).BancoInvent(caja).NroItems + 1
        
        
End If

If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If UserList(UserIndex).BancoInvent(Caja).Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
        'Menor que MAX_INV_OBJS
        UserList(UserIndex).BancoInvent(Caja).Object(Slot).ObjIndex = obji
        UserList(UserIndex).BancoInvent(Caja).Object(Slot).Amount = UserList(UserIndex).BancoInvent(Caja).Object(Slot).Amount + Cantidad
        
        Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)

    Else
        Call SendData(ToIndex, UserIndex, 0, "||El banco no puede cargar tantos objetos." & "´" & FontTypeNames.FONTTYPE_info)
    End If

Else
    Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
End If
Exit Sub
fallo:
Call LogError("Userdejaobj " & Err.number & " D: " & Err.Description)

End Sub


