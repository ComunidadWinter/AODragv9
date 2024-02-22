Attribute VB_Name = "Comercio"
Option Explicit

Sub UserCompraObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal NpcIndex As Integer, ByVal Cantidad As Integer)
On Error GoTo fallo
Dim infla As Long
Dim Descuento As String
Dim unidad As Long, monto As Long
Dim Slot As Integer
Dim obji As Integer
Dim Encontre As Boolean

If (Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(ObjIndex).Amount <= 0) Then Exit Sub

obji = Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(ObjIndex).ObjIndex

If ObjData(obji).OBJType = OBJTYPE_LLAVES And LlaveCuenta(UserIndex) = 0 Then
    Cuentas(UserIndex).Llave = ObjData(obji).Clave
    infla = (Npclist(NpcIndex).Inflacion * ObjData(obji).Valor) \ 100
'pluto:2.17------------
    If MapInfo(UserList(UserIndex).Pos.Map).Dueño = 1 And Criminal(UserIndex) Then infla = infla * 10
    If MapInfo(UserList(UserIndex).Pos.Map).Dueño = 2 And Not Criminal(UserIndex) Then infla = infla * 10
'----------------------
    
    Descuento = UserList(UserIndex).flags.Descuento
    If Descuento = 0 Then Descuento = 1 'evitamos dividir por 0!
    unidad = ((ObjData(Npclist(NpcIndex).Invent.Object(ObjIndex).ObjIndex).Valor + infla) / Descuento)
    If unidad < 1 Then unidad = 1
    monto = unidad * Cantidad
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - monto
    Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNpc, CByte(ObjIndex), Cantidad)
    Call logVentaCasa(UserList(UserIndex).Name & " compro " & ObjData(obji).Name)
    Call SendData(ToIndex, UserIndex, 0, "||Has comprado una casita :P" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
    Exit Sub
End If

If ObjData(obji).OBJType = OBJTYPE_LLAVES Then
    Call SendData(ToIndex, UserIndex, 0, "||Ya tenes una casa." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
    Exit Sub
End If
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
    'pluto:2.4.1
    If UserList(UserIndex).Stats.Peso + (Cantidad * ObjData(obji).Peso) > UserList(UserIndex).Stats.PesoMax Then
    Call SendData(ToIndex, UserIndex, 0, "P6")
    Exit Sub
    End If

    'Menor que MAX_INV_OBJS
    UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad
        
    UserList(UserIndex).Stats.Peso = UserList(UserIndex).Stats.Peso + (Cantidad * ObjData(obji).Peso)
Call SendUserStatsPeso(UserIndex)
    'pluto:2-3-04
    If Npclist(NpcIndex).Comercia = 1 Then
    'Le sustraemos el valor en oro del obj comprado
    infla = (Npclist(NpcIndex).Inflacion * ObjData(obji).Valor) \ 100
'pluto:2.17------------
    If MapInfo(UserList(UserIndex).Pos.Map).Dueño = 1 And Criminal(UserIndex) Then infla = infla * 10
    If MapInfo(UserList(UserIndex).Pos.Map).Dueño = 2 And Not Criminal(UserIndex) Then infla = infla * 10
'----------------------

    Descuento = UserList(UserIndex).flags.Descuento
    If Descuento = 0 Then Descuento = 1 'evitamos dividir por 0!
    unidad = ((ObjData(Npclist(NpcIndex).Invent.Object(ObjIndex).ObjIndex).Valor + infla) / Descuento)
   'pluto:6.8-------------
If EventoDia = 5 Then
unidad = unidad - Porcentaje(unidad, 20)
End If
 '-------------------------------
    If unidad < 1 Then unidad = 1
    monto = unidad * Cantidad
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - monto
    'tal vez suba el skill comerciar ;-)
    Call SubirSkill(UserIndex, Comerciar)
    End If
    If Npclist(NpcIndex).Comercia = 2 Then
    'Le sustraemos el valor en puntos del obj comprado
    infla = (Npclist(NpcIndex).Inflacion * ObjData(obji).Valor) \ 100
'pluto:2.17------------
If MapInfo(UserList(UserIndex).Pos.Map).Dueño = 1 And Criminal(UserIndex) Then infla = infla * 10
If MapInfo(UserList(UserIndex).Pos.Map).Dueño = 2 And Not Criminal(UserIndex) Then infla = infla * 10
'----------------------

    Descuento = UserList(UserIndex).flags.Descuento
    If Descuento = 0 Then Descuento = 1 'evitamos dividir por 0!
    unidad = ((ObjData(Npclist(NpcIndex).Invent.Object(ObjIndex).ObjIndex).Valor + infla) / Descuento)
'pluto:6.8-------------
If EventoDia = 5 Then
unidad = unidad - Porcentaje(unidad, 20)
End If
 '-------------------------------
    monto = unidad * Cantidad
    UserList(UserIndex).Stats.Puntos = UserList(UserIndex).Stats.Puntos - monto
    End If
    
'    If UserList(UserIndex).Stats.GLD < 0 Then UserList(UserIndex).Stats.GLD = 0
    
    Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNpc, CByte(ObjIndex), Cantidad)
Else
      Call SendData(ToIndex, UserIndex, 0, "P7")
      End If

Exit Sub
fallo:
Call LogError("USERCOMPRAOBJ " & UserList(UserIndex).Name & "npc: " & Npclist(NpcIndex).Name & " obj: " & ObjIndex & " can: " & Cantidad & " " & Err.number & " D: " & Err.Description)


End Sub


Sub NpcCompraObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
On Error GoTo fallo
Dim Slot As Integer
Dim obji As Integer
Dim NpcIndex As Integer
Dim infla As Long
Dim monto As Long
      
If Cantidad < 1 Then Exit Sub

NpcIndex = UserList(UserIndex).flags.TargetNpc
obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex

'pluto:2-3-04
If Npclist(NpcIndex).Comercia <> 1 Then
 Call SendData(ToIndex, UserIndex, 0, "||No compro Objetos." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

If ObjData(obji).Newbie = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||No comercio objetos para newbies." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
End If

If Npclist(NpcIndex).TipoItems <> OBJTYPE_CUALQUIERA And Npclist(NpcIndex).TipoItems <> 888 Then
    '¿Son los items con los que comercia el npc?
    If Npclist(NpcIndex).TipoItems <> ObjData(obji).OBJType Then
            Call SendData(ToIndex, UserIndex, 0, "||El npc no esta interesado en comprar ese objeto." & "´" & FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
    End If
End If
'pluto:2.17
If Npclist(NpcIndex).TipoItems = 888 And (ObjData(obji).Real = 0 Or ObjData(obji).Vendible = 1) Then
Call SendData(ToIndex, UserIndex, 0, "||El npc no esta interesado en comprar ese objeto." & "´" & FontTypeNames.FONTTYPE_WARNING)
Exit Sub
End If

'pluto:2.4.1
If ObjData(obji).OBJType = 60 Then
Call SendData(ToIndex, UserIndex, 0, "||El npc no esta interesado en comprar ese objeto." & "´" & FontTypeNames.FONTTYPE_WARNING)
Exit Sub
End If
'pluto:2.8.0
If ObjData(obji).Vendible = 1 Then
Call SendData(ToIndex, UserIndex, 0, "||El npc no esta interesado en comprar ese objeto." & "´" & FontTypeNames.FONTTYPE_WARNING)
Exit Sub
End If




'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = obji And _
         Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
        
            If Slot > MAX_INVENTORY_SLOTS Then
                Exit Do
            End If
Loop

'Sino se fija por un slot vacio antes del slot devuelto
If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
'                Call SendData(ToIndex, NpcIndex, 0, "||El npc no puede cargar mas objetos." & FONTTYPENAMES.FONTTYPE_INFO)
'                Exit Sub
                Exit Do
            End If
        Loop
        If Slot <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
        
        
End If

If Slot <= MAX_INVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
        'Menor que MAX_INV_OBJS
        Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = obji
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad
        
        Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        'Le sumamos al user el valor en oro del obj vendido
        monto = ((ObjData(obji).Valor \ 3 + infla) * Cantidad)
        Call AddtoVar(UserList(UserIndex).Stats.GLD, monto, MAXORO)
        'tal vez suba el skill comerciar ;-)
        Call SubirSkill(UserIndex, Comerciar)
    
    Else
        Call SendData(ToIndex, UserIndex, 0, "||El npc no puede cargar tantos objetos." & "´" & FontTypeNames.FONTTYPE_info)
    End If

Else
    Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
    'Le sumamos al user el valor en oro del obj vendido
    monto = ((ObjData(obji).Valor \ 3 + infla) * Cantidad)
    Call AddtoVar(UserList(UserIndex).Stats.GLD, monto, MAXORO)
End If
Exit Sub
fallo:
Call LogError("NPCCOMPRAOBJ" & Err.number & " D: " & Err.Description)

End Sub

Sub IniciarCOmercioNPC(ByVal UserIndex As Integer)
On Error GoTo fallo

'Mandamos el Inventario
Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
'Hacemos un Update del inventario del usuario
Call UpdateUserInv(True, UserIndex, 0)
'Atcualizamos el dinero
Call SendUserStatsOro(UserIndex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
SendData2 ToIndex, UserIndex, 0, 10
UserList(UserIndex).flags.Comerciando = True

Exit Sub
fallo:
Call LogError("INICIARCOMERCIONPC" & Err.number & " D: " & Err.Description)


End Sub

Sub NPCVentaItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer, ByVal NpcIndex As Integer)
On Error GoTo fallo

Dim infla As Long
Dim val As Long
Dim Desc As String

'pluto:2.10
If Cantidad < 1 Or NpcIndex < 1 Or UserIndex < 1 Or i < 0 Or i > 20 Then Exit Sub

'NPC VENDE UN OBJ A UN USUARIO
Call SendUserStatsOro(UserIndex)
'Calculamos el valor unitario
infla = Int((Npclist(NpcIndex).Inflacion * ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor) / 100)
'pluto:2.17------------
If MapInfo(UserList(UserIndex).Pos.Map).Dueño = 1 And Criminal(UserIndex) Then infla = infla * 10
If MapInfo(UserList(UserIndex).Pos.Map).Dueño = 2 And Not Criminal(UserIndex) Then infla = infla * 10
'----------------------

Desc = Descuento(UserIndex)
If Desc = 0 Then Desc = 1 'evitamos dividir por 0!
val = (ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor + infla) / Desc

'pluto:6.8-------------
If EventoDia = 5 Then
val = val - Porcentaje(val, 20)
End If
 '-------------------------------

If val < 1 Then val = 1


'pluto:2-3-04
If (UserList(UserIndex).Stats.GLD >= (val * Cantidad) And Npclist(NpcIndex).Comercia = 1) Or (UserList(UserIndex).Stats.Puntos >= (val * Cantidad) And Npclist(NpcIndex).Comercia = 2) Then

       If Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(i).Amount > 0 Then
            If UserList(UserIndex).flags.Privilegios > 0 And UserList(UserIndex).flags.Privilegios < 3 Then Exit Sub
            If Cantidad > Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(i).Amount Then Cantidad = Npclist(UserList(UserIndex).flags.TargetNpc).Invent.Object(i).Amount
            'Agregamos el obj que compro al inventario
            Call UserCompraObj(UserIndex, CInt(i), UserList(UserIndex).flags.TargetNpc, Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el oro
            Call SendUserStatsOro(UserIndex)
            'Actualizamos la ventana de comercio
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
            Call UpdateVentanaComercio(i, 0, UserIndex)
        
       End If
Else
    'pluto:2-3-04
    If Npclist(NpcIndex).Comercia = 1 Then Call SendData(ToIndex, UserIndex, 0, "||No tenes suficiente Oro." & "´" & FontTypeNames.FONTTYPE_info) Else Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes DraGPuntos." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
End If


Exit Sub
fallo:
'pluto:2.10
Call LogError("NPCVENTAITEM Npc:" & Npclist(NpcIndex).Name & " NpcIndex: " & NpcIndex & " Jug: " & UserList(UserIndex).Name & " " & Err.number & " D: " & Err.Description & "Obj: " & i & "Cant: " & Cantidad & "TipoNpc: " & Npclist(NpcIndex).NPCtype)


End Sub
Sub NPCCompraItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

On Error GoTo fallo

'NPC COMPRA UN OBJ A UN USUARIO
Call SendUserStatsOro(UserIndex)
'pluto:vender oro
If UserList(UserIndex).Invent.Object(Item).ObjIndex = 12 Then Exit Sub
'pluto:fin vender oro


If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then

            If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
            'Agregamos el obj que compro al inventario
            Call NpcCompraObj(UserIndex, CInt(Item), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el oro
            Call SendUserStatsOro(UserIndex)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
            'Actualizamos la ventana de comercio
            
            Call UpdateVentanaComercio(Item, 1, UserIndex)
            
End If

Exit Sub
fallo:
Call LogError("NPCCOMPRAITEM" & Err.number & " D: " & Err.Description)


End Sub

Sub UpdateVentanaComercio(ByVal Slot As Integer, ByVal NpcInv As Byte, ByVal UserIndex As Integer)
On Error GoTo fallo
Call SendData2(ToIndex, UserIndex, 0, 70, Slot & "," & NpcInv)

Exit Sub
fallo:
Call LogError("UPDATEVENTANACOMERCIO" & Err.number & " D: " & Err.Description)


End Sub

Function Descuento(ByVal UserIndex As Integer) As String
On Error GoTo fallo
'Establece el descuento en funcion del skill comercio
Dim PtsComercio As Integer
PtsComercio = CInt(UserList(UserIndex).Stats.UserSkills(Comerciar) / 2)

If PtsComercio <= 10 And PtsComercio > 5 Then
    UserList(UserIndex).flags.Descuento = 1.1
    Descuento = 1.1
ElseIf PtsComercio <= 20 And PtsComercio >= 11 Then
    UserList(UserIndex).flags.Descuento = 1.2
    Descuento = 1.2
ElseIf PtsComercio <= 30 And PtsComercio >= 19 Then
    UserList(UserIndex).flags.Descuento = 1.3
    Descuento = 1.3
ElseIf PtsComercio <= 40 And PtsComercio >= 29 Then
    UserList(UserIndex).flags.Descuento = 1.4
    Descuento = 1.4
ElseIf PtsComercio <= 50 And PtsComercio >= 39 Then
    UserList(UserIndex).flags.Descuento = 1.5
    Descuento = 1.5
ElseIf PtsComercio <= 60 And PtsComercio >= 49 Then
    UserList(UserIndex).flags.Descuento = 1.6
    Descuento = 1.6
ElseIf PtsComercio <= 70 And PtsComercio >= 59 Then
    UserList(UserIndex).flags.Descuento = 1.7
    Descuento = 1.7
ElseIf PtsComercio <= 80 And PtsComercio >= 69 Then
    UserList(UserIndex).flags.Descuento = 1.8
    Descuento = 1.8
ElseIf PtsComercio <= 99 And PtsComercio >= 79 Then
    UserList(UserIndex).flags.Descuento = 1.9
    Descuento = 1.9
ElseIf PtsComercio <= 999999 And PtsComercio >= 99 Then
    UserList(UserIndex).flags.Descuento = 2
    Descuento = 2
Else
    UserList(UserIndex).flags.Descuento = 0
    Descuento = 0
End If
Exit Function
fallo:
Call LogError("DESCUENTO" & Err.number & " D: " & Err.Description)

End Function



Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
On Error GoTo fallo
'Enviamos el inventario del npc con el cual el user va a comerciar...
Dim i As Integer
Dim infla As Long
Dim Desc As String
Dim val As Long
Desc = Descuento(UserIndex)
If Desc = 0 Then Desc = 1 'evitamos dividir por 0!

For i = 1 To MAX_INVENTORY_SLOTS
  If Npclist(NpcIndex).Invent.Object(i).ObjIndex > 0 Then
        'Calculamos el porc de inflacion del npc
        infla = (Npclist(NpcIndex).Inflacion * ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor) / 100
'pluto:2.17------------
    If MapInfo(UserList(UserIndex).Pos.Map).Dueño = 1 And Criminal(UserIndex) Then infla = infla * 10
    If MapInfo(UserList(UserIndex).Pos.Map).Dueño = 2 And Not Criminal(UserIndex) Then infla = infla * 10
'----------------------


'-----
val = (ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor + infla) / Desc
   'pluto:6.8-------------
   If EventoDia = 5 Then
val = val - Porcentaje(val, 20)
End If
 '-------------------------------
 If val < 1 Then val = 1
        'pluto:6.0A
        Call SendData2(ToIndex, UserIndex, 0, 45, Npclist(NpcIndex).Invent.Object(i).ObjIndex & "," & Npclist(NpcIndex).Invent.Object(i).Amount & "," & val)
        'SendData2 ToIndex, UserIndex, 0, 45, _
        'ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Name _
       ' & "," & Npclist(NpcIndex).Invent.Object(i).Amount & _
        '"," & val _
       ' & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).GrhIndex _
       ' & "," & Npclist(NpcIndex).Invent.Object(i).ObjIndex _
       ' & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).OBJType _
       ' & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).MaxHIT _
       ' & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).MinHIT _
       ' & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).MaxDef
  Else
       Call SendData2(ToIndex, UserIndex, 0, 45, 0)
       ' SendData2 ToIndex, UserIndex, 0, 45, _
        '"Nada" _
        '& "," & 0 & _
        '"," & 0 _
        '& "," & 0 _
        '& "," & 0 _
        '& "," & 0 _
        '& "," & 0 _
        '& "," & 0 _
       ' & "," & 0 _
       ' & "," & 0 _
       ' & "," & 0 _
       ' & "," & 0 _
       ' & "," & 0 _
       ' & "," & 0 _
       ' & "," & 0 _
       ' & "," & 0
  End If
  
Next

Exit Sub
fallo:
Call LogError("ENVIARNPCINV" & Err.number & " D: " & Err.Description)


End Sub
