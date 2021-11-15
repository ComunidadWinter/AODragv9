Attribute VB_Name = "InvUsuario"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

If ObjIndex = 0 Then Exit Function

Dim flag As Boolean

'Admins can use ANYTHING!
If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
        Dim i As Integer
        For i = 1 To NUMCLASES
            If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
                ClasePuedeUsarItem = False
                Exit Function
            End If
        Next i
    End If
End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Function ClasePuedeUsarHechizo(ByVal UserIndex As Integer, ByVal SpellIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

If SpellIndex = 0 Then Exit Function

Dim flag As Boolean

'Los admins pueden usar todo
If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    If Not Hechizos(SpellIndex).ExclusivoClase(1) = 0 Then
        Dim i As Integer
        For i = 1 To NUMCLASES
            If Hechizos(SpellIndex).ExclusivoClase(i) = UserList(UserIndex).clase Then
            Debug.Print "hechizo:" & Hechizos(SpellIndex).ExclusivoClase(i) & " Usuario:" & UserList(UserIndex).clase
                ClasePuedeUsarHechizo = True
                Exit Function
            End If
        Next i
    Else
        ClasePuedeUsarHechizo = True
    End If
Else
    ClasePuedeUsarHechizo = True
End If

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarHechizo")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next j

'[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
If UCase$(MapInfo(UserList(UserIndex).Pos.Map).Restringir) = "NEWBIE" Then
    
    Call WarpUserChar(UserIndex, 1, 44, 88, True)

End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Dim j As Integer
        For j = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(j).ObjIndex = 0
                .Invent.Object(j).Amount = 0
                .Invent.Object(j).Equipped = 0
                
        Next
        
        .Invent.NroItems = 0
        
        .Invent.ArmourEqpObjIndex = 0
        .Invent.ArmourEqpSlot = 0
        
        .Invent.WeaponEqpObjIndex = 0
        .Invent.WeaponEqpSlot = 0
        
        .Invent.CascoEqpObjIndex = 0
        .Invent.CascoEqpSlot = 0
        
        .Invent.EscudoEqpObjIndex = 0
        .Invent.EscudoEqpSlot = 0
        
        .Invent.AnilloEqpObjIndex = 0
        .Invent.AnilloEqpSlot = 0
        
        .Invent.MunicionEqpObjIndex = 0
        .Invent.MunicionEqpSlot = 0
        
        .Invent.BarcoObjIndex = 0
        .Invent.BarcoSlot = 0
        
        .Invent.MonturaObjIndex = 0
        .Invent.MonturaSlot = 0
    End With

End Sub

Sub TirarOro(ByVal cantidad As Long, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo errHandler

Debug.Print "LA"
'If Cantidad > 100000 Then Exit Sub
If UserList(UserIndex).Lac.LTirar.Puedo = False Then Exit Sub

'SI EL Pjta TIENE ORO LO TIRAMOS
If (cantidad > 0) And (cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon (guardo el oro tirado si supera los 50k)
        If cantidad > 50000 Then
            Dim j As Integer
            Dim k As Integer
            Dim M As Integer
            Dim Cercanos As String
            M = UserList(UserIndex).Pos.Map
            For j = UserList(UserIndex).Pos.X - 10 To UserList(UserIndex).Pos.X + 10
                For k = UserList(UserIndex).Pos.Y - 10 To UserList(UserIndex).Pos.Y + 10
                    If InMapBounds(M, j, k) Then
                        If MapData(M, j, k).UserIndex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(M, j, k).UserIndex).Name & ","
                        End If
                    End If
                Next k
            Next j
            Call LogDesarrollo(UserList(UserIndex).Name & " tira oro. Cercanos: " & Cercanos)
        End If
        '/Seguridad
        Dim Extra As Long
        Dim TeniaOro As Long
        TeniaOro = UserList(UserIndex).Stats.GLD
        If cantidad > 500000 Then 'Para evitar explotar demasiado
            Extra = cantidad - 500000
            cantidad = 500000
        End If
        
        Do While (cantidad > 0)
            
            If cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                cantidad = cantidad - MiObj.Amount
            Else
                MiObj.Amount = cantidad
                cantidad = cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            
            If EsGM(UserIndex) Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
            Dim AuxPos As worldPos
            
            If UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, False)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.Amount
                End If
            Else
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, True)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.Amount
                End If
            End If
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
        If TeniaOro = UserList(UserIndex).Stats.GLD Then Extra = 0
        If Extra > 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Extra
        End If
    
End If

Exit Sub

errHandler:

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal cantidad As Integer)

On Error GoTo errHandler

    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
        If .Amount <= cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, Slot)
        End If
        
        'Quita un objeto
        .Amount = .Amount - cantidad
        '¿Quedan mas?
        If .Amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .ObjIndex = 0
            .Amount = 0
        End If
    End With

Exit Sub

errHandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.Description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

On Error GoTo errHandler

Dim NullObj As UserObj
Dim LoopC As Long

With UserList(UserIndex)
    'Actualiza un solo slot
    If Not UpdateAll Then
    
        'Actualiza el inventario
        If .Invent.Object(Slot).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, Slot, .Invent.Object(Slot))
        Else
            Call ChangeUserInv(UserIndex, Slot, NullObj)
        End If
    
    Else
    
    'Actualiza todos los slots
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            'Actualiza el inventario
            If .Invent.Object(LoopC).ObjIndex > 0 Then
                Call ChangeUserInv(UserIndex, LoopC, .Invent.Object(LoopC))
            Else
                Call ChangeUserInv(UserIndex, LoopC, NullObj)
            End If
        Next LoopC
    End If
    
    Exit Sub
End With

errHandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.Description)

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim Obj As Obj

If num > 0 Then
  
  If num > UserList(UserIndex).Invent.Object(Slot).Amount Then num = UserList(UserIndex).Invent.Object(Slot).Amount

  
  'Check objeto en el suelo
  If MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.ObjIndex = 0 Or MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex Then
        Obj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If ObjData(Obj.ObjIndex).OBJType <> eOBJType.otRopaMontura Then
          If UserList(UserIndex).flags.QueMontura Then _
              Call WriteConsoleMsg(UserIndex, "No puedes tirar eso mientras estas montado.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        If num + MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount
        End If
        
        Obj.Amount = num
        
        Call MakeObj(Obj, Map, X, Y)
        Call QuitarUserInvItem(UserIndex, Slot, num)
        Call UpdateUserInv(False, UserIndex, Slot)
        
        If ObjData(Obj.ObjIndex).NoLimpiar = 0 Then ' se puede eliminar?
            If ObjData(Obj.ObjIndex).OBJType <> eOBJType.otGuita Then ' las monedas no se borran
                Call aLimpiarMundo.AddItem(Map, X, Y) ' GSZAO
            End If
        End If
        
        If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(Obj.ObjIndex).Name)
        
        'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Obj.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).Name & " tiró al piso " & Obj.Amount & " " & ObjData(Obj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
        ElseIf Obj.Amount > 5000 Then 'Es mucha cantidad? > Subí a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)
        'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Obj.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).Name & " tiró al piso " & Obj.Amount & " " & ObjData(Obj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
            End If
        End If
  Else
    Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
  End If
    
End If

End Sub

Sub EraseObj(ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - num

If MapData(Map, X, Y).ObjInfo.Amount <= 0 Then
    MapData(Map, X, Y).ObjInfo.ObjIndex = 0
    MapData(Map, X, Y).ObjInfo.Amount = 0
    
    Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))
End If

End Sub

Sub MakeObj(ByRef Obj As Obj, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

    If MapData(Map, X, Y).ObjInfo.ObjIndex = Obj.ObjIndex Then
        MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount + Obj.Amount
    Else
        MapData(Map, X, Y).ObjInfo = Obj
        
        Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(Obj.ObjIndex).GrhIndex, X, Y))
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errHandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
errHandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj
Dim ObjPos As String

With UserList(UserIndex)

    '¿Hay algun obj?
    If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex > 0 Then
        '¿Esta permitido agarrar este obj?
        If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
            Dim X As Integer
            Dim Y As Integer
            Dim Slot As Byte
            
            
            
            X = .Pos.X
            Y = .Pos.Y
            Obj = ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex)
            MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
            MiObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
            
            If Obj.OBJType = otGuita Then
                .Stats.GLD = .Stats.GLD + MiObj.Amount
                Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(.Pos.X, .Pos.Y, MiObj.Amount, 4))
                Call WriteUpdateUserStats(UserIndex)
                Call WriteConsoleMsg(UserIndex, "¡Obtienes " & MiObj.Amount & " monedas de oro!", FontTypeNames.FONTTYPE_oro)
                Exit Sub
            End If
            
            If MeterItemEnInventario(UserIndex, MiObj) Then
                'Quitamos el objeto
                Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                
                If ObjData(MiObj.ObjIndex).NoLimpiar = 0 Then ' juntamos algo que puede estar en la lista de limpieza
                    Call aLimpiarMundo.RemoveItem(.Pos.Map, .Pos.X, .Pos.Y)
                End If
                
                If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
                
                'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                'Es un Objeto que tenemos que loguear?
                If ObjData(MiObj.ObjIndex).Log = 1 Then
                    ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                    Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                ElseIf MiObj.Amount > MAX_INVENTORY_OBJS - 1000 Then 'Es mucha cantidad?
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                        ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                        Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                    End If
                End If
            End If
            
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)
    End If
End With
End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo errHandler

'Desequipa el item slot del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

With UserList(UserIndex)

    If (Slot < LBound(.Invent.Object)) Or (Slot > UBound(.Invent.Object)) Then
        Exit Sub
    ElseIf .Invent.Object(Slot).ObjIndex = 0 Then
        Exit Sub
    End If
    
    Obj = ObjData(.Invent.Object(Slot).ObjIndex)
    ObjIndex = .Invent.Object(Slot).ObjIndex
    
    Select Case Obj.OBJType
        Case eOBJType.otWeapon
            
            Call DesEquiparItemBonificador(UserIndex, ObjIndex)
        
            .Invent.Object(Slot).Equipped = 0
            .Invent.WeaponEqpObjIndex = 0
            .Invent.WeaponEqpSlot = 0
                
            If Not .flags.Mimetizado = 1 Then
                .Char.WeaponAnim = NingunArma
                Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
            End If
        
        Case eOBJType.otFlechas
            .Invent.Object(Slot).Equipped = 0
            .Invent.MunicionEqpObjIndex = 0
            .Invent.MunicionEqpSlot = 0
        
        Case eOBJType.otAnillo
            .Invent.Object(Slot).Equipped = 0
            .Invent.AnilloEqpObjIndex = 0
            .Invent.AnilloEqpSlot = 0
        
        Case eOBJType.otArmadura
             Call DesEquiparItemBonificador(UserIndex, ObjIndex)
        
            .Invent.Object(Slot).Equipped = 0
            .Invent.ArmourEqpObjIndex = 0
            .Invent.ArmourEqpSlot = 0
            
            Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                
        Case eOBJType.otCASCO
        
             Call DesEquiparItemBonificador(UserIndex, ObjIndex)
        
            .Invent.Object(Slot).Equipped = 0
            .Invent.CascoEqpObjIndex = 0
            .Invent.CascoEqpSlot = 0
            
            If Not .flags.Mimetizado = 1 Then
                .Char.CascoAnim = NingunCasco
                Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
            End If
        
        Case eOBJType.otESCUDO
        
             Call DesEquiparItemBonificador(UserIndex, ObjIndex)
        
            .Invent.Object(Slot).Equipped = 0
            .Invent.EscudoEqpObjIndex = 0
            .Invent.EscudoEqpSlot = 0
        
            If Not .flags.Mimetizado = 1 Then
                .Char.ShieldAnim = NingunEscudo
                Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
            End If
    End Select
    
End With
Call WriteUpdateUserStats(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub

errHandler:
    Call LogError("Error en Desquipar. Error " & Err.Number & " : " & Err.Description)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo errHandler

If ObjIndex = 0 Then Exit Function

If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Hombre
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True
    End If
Else
    SexoPuedeUsarItem = True
End If

Exit Function
errHandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjIndex = 0 Then Exit Function

If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    If ObjData(ObjIndex).Real = 1 Then
        If Not criminal(UserIndex) Then
            FaccionPuedeUsarItem = esArmada(UserIndex)
        Else
            FaccionPuedeUsarItem = False
        End If
    ElseIf ObjData(ObjIndex).Caos = 1 Then
        If criminal(UserIndex) Then
            FaccionPuedeUsarItem = esCaos(UserIndex)
        Else
            FaccionPuedeUsarItem = False
        End If
    Else
        FaccionPuedeUsarItem = True
    End If
Else
    FaccionPuedeUsarItem = True
End If
End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 01/08/2009
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
'*************************************************

On Error GoTo errHandler
With UserList(UserIndex)

    If .flags.Morph > 0 Then
        Call WriteConsoleMsg(UserIndex, "No puedes equipar/desequipar nada mientras estas transformado.", FontTypeNames.FONTTYPE_INFOBOLD)
        Exit Sub
    End If
    
    'Equipa un item del inventario
    Dim Obj As ObjData
    Dim ObjIndex As Integer
    
    ObjIndex = .Invent.Object(Slot).ObjIndex
    Obj = ObjData(ObjIndex)
    
    If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
        Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Select Case Obj.OBJType
        Case eOBJType.otWeapon
        
                If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                      FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
               
                'Si esta equipado lo quita
                If .Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    'Animacion por defecto
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = NingunArma
                    Else
                        .Char.WeaponAnim = NingunArma
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    Exit Sub
                End If
                    
                'Quitamos el elemento anterior
                If .Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                End If
                    
                Call EquiparItemBonificador(UserIndex, ObjIndex)
                
                .Invent.Object(Slot).Equipped = 1
                .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
                .Invent.WeaponEqpSlot = Slot
                
                'El sonido solo se envia si no lo produce un admin invisible
                If Not (.flags.AdminInvisible = 1) Then _
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    
                If .flags.Mimetizado = 1 Then
                    .CharMimetizado.WeaponAnim = Obj.WeaponAnim
                Else
                    .Char.WeaponAnim = Obj.WeaponAnim
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                End If
            Else
                Call WriteMultiMessage(UserIndex, eMessages.ClaseNoUsa)
            End If
        Case eOBJType.otAnillo
           If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
              FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)
                        Exit Sub
                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.AnilloEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
                    End If
            
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.AnilloEqpObjIndex = ObjIndex
                    .Invent.AnilloEqpSlot = Slot
                    
           Else
                Call WriteMultiMessage(UserIndex, eMessages.ClaseNoUsa)
           End If
        
        Case eOBJType.otFlechas
           If ClasePuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex) And _
              FaccionPuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex) Then
                    
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)
                        Exit Sub
                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.MunicionEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
                    End If
            
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.MunicionEqpObjIndex = .Invent.Object(Slot).ObjIndex
                    .Invent.MunicionEqpSlot = Slot
                    
           Else
                Call WriteMultiMessage(UserIndex, eMessages.ClaseNoUsa)
           End If
        
        Case eOBJType.otArmadura
            'If .flags.Montando = 1 Then Exit Sub
            If .flags.Navegando = 1 Then Exit Sub
            'Nos aseguramos que puede usarla
            If ClasePuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex) And _
               SexoPuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex) And _
               CheckRazaUsaRopa(UserIndex, .Invent.Object(Slot).ObjIndex) And _
               FaccionPuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex) Then
               
               'Si esta equipado lo quita
                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
                    If Not .flags.Mimetizado = 1 Then
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    Exit Sub
                End If
        
                'Quita el anterior
                If .Invent.ArmourEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
                End If
        
                Call EquiparItemBonificador(UserIndex, ObjIndex)
                
                'Lo equipa
                .Invent.Object(Slot).Equipped = 1
                .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex
                .Invent.ArmourEqpSlot = Slot

                If .flags.Mimetizado = 1 Then
                    .CharMimetizado.body = Obj.Ropaje
                Else
                    .Char.body = Obj.Ropaje
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                End If
                .flags.Desnudo = 0
                
    
            Else
                Call WriteConsoleMsg(UserIndex, "Tu clase, genero o raza no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case eOBJType.otCASCO
            If .flags.Navegando = 1 Then Exit Sub
            If ClasePuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex) Then
                'Si esta equipado lo quita
                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.CascoAnim = NingunCasco
                    Else
                        .Char.CascoAnim = NingunCasco
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    Exit Sub
                End If
        
                'Quita el anterior
                If .Invent.CascoEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
                End If

                Call EquiparItemBonificador(UserIndex, ObjIndex)
                
                'Lo equipa
                .Invent.Object(Slot).Equipped = 1
                .Invent.CascoEqpObjIndex = .Invent.Object(Slot).ObjIndex
                .Invent.CascoEqpSlot = Slot

                If .flags.Mimetizado = 1 Then
                    .CharMimetizado.CascoAnim = Obj.CascoAnim
                Else
                    .Char.CascoAnim = Obj.CascoAnim
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                End If
            Else
                Call WriteMultiMessage(UserIndex, eMessages.ClaseNoUsa)
            End If
        
        Case eOBJType.otESCUDO
            If .flags.Navegando = 1 Then Exit Sub
             If ClasePuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex) And _
                 FaccionPuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex) Then
    
                 'Si esta equipado lo quita
                 If .Invent.Object(Slot).Equipped Then
                     Call Desequipar(UserIndex, Slot)
                     If .flags.Mimetizado = 1 Then
                         .CharMimetizado.ShieldAnim = NingunEscudo
                     Else
                         .Char.ShieldAnim = NingunEscudo
                         Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                     End If
                     Exit Sub
                 End If
    
                 'Quita el anterior
                 If .Invent.EscudoEqpObjIndex > 0 Then
                     Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                 End If
         
                Call EquiparItemBonificador(UserIndex, ObjIndex)
         
                 'Lo equipa
                 .Invent.Object(Slot).Equipped = 1
                 .Invent.EscudoEqpObjIndex = .Invent.Object(Slot).ObjIndex
                 .Invent.EscudoEqpSlot = Slot
                 
                 If .flags.Mimetizado = 1 Then
                     .CharMimetizado.ShieldAnim = Obj.ShieldAnim
                 Else
                     .Char.ShieldAnim = Obj.ShieldAnim
                     
                     Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                 End If
             Else
                 Call WriteMultiMessage(UserIndex, eMessages.ClaseNoUsa)
             End If
    End Select
End With

'Actualiza
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub
errHandler:
Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.Description)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errHandler

If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    'Verifica si la raza puede usar la ropa
    If UserList(UserIndex).raza = eRaza.Humano Or _
       UserList(UserIndex).raza = eRaza.Elfo Or _
       UserList(UserIndex).raza = eRaza.Drow Or _
       UserList(UserIndex).raza = eRaza.Orco Or UserList(UserIndex).raza = eRaza.NoMuerto Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
    Else
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
    End If
    'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
   ' If (UserList(UserIndex).raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Or (UserList(UserIndex).raza <> eRaza.NoMuerto) And ObjData(ItemIndex).RazaDrow Then
   '     CheckRazaUsaRopa = False
   ' End If
Else
    CheckRazaUsaRopa = True
End If
Exit Function
errHandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 01/08/2009
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin, except to its own client
'*************************************************

Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj
Dim i As Byte

With UserList(UserIndex)

If .Invent.Object(Slot).Amount = 0 Then Exit Sub

Obj = ObjData(.Invent.Object(Slot).ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

'¿Está trabajando?
If .flags.Makro <> 0 Then
    Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Obj.OBJType = eOBJType.otWeapon Then
If UserList(UserIndex).Lac.LUsar.Puedo = False Then Exit Sub
    If Obj.proyectil = 1 Then
        
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
End If

ObjIndex = .Invent.Object(Slot).ObjIndex
.flags.TargetObjInvIndex = ObjIndex
.flags.TargetObjInvSlot = Slot

Select Case Obj.OBJType
        
    '07/11/2015 Irongete: Items contenedores
    Case eOBJType.otContenedores
    
        Dim ObjContenedor As Obj
        MiObj.Amount = 1
        MiObj.ObjIndex = 384
    
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
    
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
    
        'Actualizar el inventario
        Call UpdateUserInv(False, UserIndex, Slot)


    Case eOBJType.otUseOnce
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If

        'Usa el item
        .Stats.MinHam = .Stats.MinHam + Obj.MinHam
        If .Stats.MinHam > .Stats.MaxHam Then _
            .Stats.MinHam = .Stats.MaxHam
        .flags.Hambre = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        'Sonido
        
        If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call UpdateUserInv(False, UserIndex, Slot)

    Case eOBJType.otGuita
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        
        .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).Amount
        .Invent.Object(Slot).Amount = 0
        .Invent.Object(Slot).ObjIndex = 0
        .Invent.NroItems = .Invent.NroItems - 1
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteUpdateGold(UserIndex)
        
    Case eOBJType.otWeapon
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        
        If Not .Stats.MinSta > 0 Then
            If .genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Estas muy cansado", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Estas muy cansada", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        
        
        If ObjData(ObjIndex).proyectil = 1 Then
            If .Invent.Object(Slot).Equipped = 0 Then
                Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Call WriteWorkRequestTarget(UserIndex, Proyectiles)
        Else
            If .flags.targetObj = Leña Then
                If .Invent.Object(Slot).ObjIndex = DAGA Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call TratarDeHacerFogata(.flags.TargetObjMap, _
                         .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                End If
            End If
        End If

        
        Select Case ObjIndex
            Case CAÑA_PESCA, RED_PESCA
                Call WriteWorkRequestTarget(UserIndex, eSkill.pesca)
            Case MARTILLO_HERRERO
                Call WriteWorkRequestTarget(UserIndex, eSkill.Herreria)
            Case SERRUCHO_CARPINTERO
                Call EnivarObjConstruibles(UserIndex)
                Call WriteShowCarpenterForm(UserIndex)
        End Select
        
    
    Case eOBJType.otPociones
        If UserList(UserIndex).Lac.LPociones.Puedo = False Then Exit Sub
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        
        If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
            Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra poción!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .flags.TomoPocion = True
        .flags.TipoPocion = Obj.TipoPocion
                
        Select Case .flags.TipoPocion
        
            Case 1 'Modif la agilidad
                .flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                If UserList(UserIndex).clase = Bard Then
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 15) Then _
                        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 15)
                Else
                    If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 13) Then _
                        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 13)
                End If
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                Call WriteUpdateDexterity(UserIndex)
                
            Case 2 'Modif la fuerza
                .flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                If UserList(UserIndex).clase = Bard Then
                    If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 15) Then _
                        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 15)
                Else
                    If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 13) Then _
                        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 13)
                End If
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                Call WriteUpdateStrenght(UserIndex)
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                .Stats.MinHP = .Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If .Stats.MinHP > VidaMaxima(UserIndex) Then _
                    .Stats.MinHP = VidaMaxima(UserIndex)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                'nuevo calculo para recargar mana
                .Stats.MinMAN = .Stats.MinMAN + Porcentaje(ManaMaxima(UserIndex), 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV
                If .Stats.MinMAN > ManaMaxima(UserIndex) Then _
                    .Stats.MinMAN = ManaMaxima(UserIndex)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                
            Case 5 ' Pocion violeta
                If .flags.Envenenado = 1 Then
                    .flags.Envenenado = 0
                    Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                
            Case 6  ' Pocion Negra
                If .flags.Privilegios And PlayerType.User Then
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call UserDie(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                End If
       End Select
       Call WriteUpdateUserStats(UserIndex)
       Call UpdateUserInv(False, UserIndex, Slot)

    'Esfera de Experiencia, al tomarla nos dara un numero aleatorio de exp
    Case eOBJType.otEsferadeExp
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim random As Long
        random = RandomNumber(10000, 200000)
        .Stats.Exp = .Stats.Exp + random
        Call WriteConsoleMsg(UserIndex, "¡Felicitaciones, la esfera sagrada de experiencia te ha otorgado " & random & " puntos de experiencia", FontTypeNames.FONTTYPE_CONSEJO)
        Call WriteUpdateExp(UserIndex)
        Call CheckUserLevel(UserIndex)
        
'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)

    'Amuleto Ankh
    Case eOBJType.otPiedraResu
    
        If .flags.Muerto = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Este item solo lo puedes usar estando muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapInfo(.Pos.Map).Pk = False Then
            Call WriteConsoleMsg(UserIndex, "¡Este objeto no puede utilizarse en ciudades!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case .Pos.Map
            Case 1, 2, 3, 4, 9, 14
                Call WarpUserChar(UserIndex, 1, 31, 88, True)
            Case 6, 12, 13
                Call WarpUserChar(UserIndex, 12, 81, 82, True)
            Case 5, 7, 8, 11, 19
                Call WarpUserChar(UserIndex, 7, 66, 28, True)
            Case Else
                Call WarpUserChar(UserIndex, 1, 31, 88, True)
        End Select
        
     Case eOBJType.otBebidas
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        ' Los admin invisibles solo producen sonidos a si mismos
        If .flags.AdminInvisible = 1 Then
            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otLlaves
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        
        If .flags.targetObj = 0 Then Exit Sub
        TargObj = ObjData(.flags.targetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = Obj.clave Then
         
                        MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                        .flags.targetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                        Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = Obj.clave Then
                        MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                        Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                        .flags.targetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call WriteConsoleMsg(UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                  Exit Sub
            End If
        End If
    
    Case eOBJType.otBotellaVacia
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        If Not HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
            Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        MiObj.Amount = 1
        MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otBotellaLlena
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        MiObj.Amount = 1
        MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otPergaminos
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If

        If .flags.Hambre = 0 And _
            .flags.Sed = 0 Then
            If ClasePuedeUsarHechizo(UserIndex, Obj.HechizoIndex) Then
                Call AgregarHechizo(UserIndex, Slot)
                Call UpdateUserInv(False, UserIndex, Slot)
            Else
                Call WriteConsoleMsg(UserIndex, "Tu conocimiento de las artes arcanas no te permiten comprender este hechizo.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
        End If
    Case eOBJType.otMinerales
        If .flags.Muerto = 1 Then
             Call WriteMultiMessage(UserIndex, eMessages.Muerto)
             Exit Sub
        End If
        Call WriteWorkRequestTarget(UserIndex, FundirMetal)
       
    Case eOBJType.otInstrumentos
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        
        If Obj.Real Then '¿Es el Cuerno Real?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(.Pos.Map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                End If
                
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf Obj.Caos Then '¿Es el Cuerno Legión?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(.Pos.Map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                End If
                
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'Si llega aca es porque es o Laud o Tambor o Flauta
        ' Los admin invisibles solo producen sonidos a si mismos
        If .flags.AdminInvisible = 1 Then
            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
        End If
       
    Case eOBJType.otBarcos
        'Verifica si esta aproximado al agua antes de permitirle navegar
        If .Stats.ELV < 25 Then
            If .Stats.ELV < 20 Then
                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        If .flags.QueMontura Then
            WriteConsoleMsg UserIndex, "¡¡No puedes navegar si estás montando!!", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        If ((LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, True, False) _
                Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, True, False) _
                Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, True, False) _
                Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, True, False)) _
                And .flags.Navegando = 0) _
                Or .flags.Navegando = 1 Then
            Call DoNavega(UserIndex, Obj, Slot)
        Else
            Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
        End If
        
    Case eOBJType.otMonturas
            If .flags.Muerto = 1 Then
                Call WriteMultiMessage(UserIndex, eMessages.Muerto)
                Exit Sub
            End If
            
            '¿Tiene monturas?
            If .Stats.NUMMONTURAS = 0 Then
                Call WriteConsoleMsg(UserIndex, "¡No tienes monturas!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .flags.QueMontura = 0 Then
                If Intemperie(UserIndex) = False Then
                    Call WriteConsoleMsg(UserIndex, "¡No puedes montar aqui!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            
            If UserList(UserIndex).flags.Morph > 0 Then
                Call WriteConsoleMsg(UserIndex, "No puedes montar mientras estas transformado.", FontTypeNames.FONTTYPE_INFOBOLD)
                Exit Sub
            End If
            
            If Not Obj.skdomar <= UserList(UserIndex).Stats.UserSkills(eSkill.Domar) Then
                Call WriteConsoleMsg(UserIndex, "Necesitas " & Obj.skdomar & " en domar animales para montar esa mascota.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If Not TieneObjetos(1100, 1, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "¡No tienes ropa de montar!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If ((LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False) _
                Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False) _
                Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False) _
                Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y, True, False)) _
                And .flags.Navegando = 0) _
                Or .flags.Navegando = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes montar en el agua!", FontTypeNames.FONTTYPE_INFO)
            Else
                '¿Ya tiene esa mascota?
                For i = 1 To UserList(UserIndex).Stats.NUMMONTURAS
                    If UserList(UserIndex).flags.Montura(i).tipo = ObjData(.Invent.Object(Slot).ObjIndex).IndiceMontura Then
                        'Si la tiene, dime que slot es de tu lista
                        Call DoEquita(UserIndex, Slot, i)
                        Exit Sub
                    End If
                Next i
                'Si llego hasta aqui es por que no la tiene.
                Call WriteConsoleMsg(UserIndex, "¡No tienes esa montura!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        '<-------------> PASAJES <----------->
        
    Case eOBJType.otPasajes
            If .flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡Estás muerto!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
            If .flags.TargetNpcTipo <> Pirata Then
                Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre el marinero.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
            If Distancia(NPCList(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                Exit Sub
            End If
                
            If .Pos.Map <> Obj.DesdeMap Then
                Call WriteConsoleMsg(UserIndex, "El pasaje no lo compraste aquí! Largate!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
            If Not MapaValido(Obj.HastaMap) Then
                Call WriteConsoleMsg(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
            If .Stats.UserSkills(eSkill.Navegacion) < Obj.CantidadSkill Then
                Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje no puedo llevarte. Necesitas " & Obj.CantidadSkill & " skills para utilizar este pasaje. Consulta el manual del juego en http://www.aodrag.es/wiki/ para saber cómo conseguirlos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
            If .Stats.ELV < 10 Then
                Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, necesitas ser nivel 10 como minimo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
            Call WarpUserChar(UserIndex, Obj.HastaMap, Obj.HastaX, Obj.HastaY, True)
            Call WriteConsoleMsg(UserIndex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_CENTINELA)
            .Stats.MinAGU = 0
            .Stats.MinHam = 0
            .flags.Sed = 1
            .flags.Hambre = 1
            Call WriteUpdateHungerAndThirst(UserIndex)
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)
        
        
      Case eOBJType.otCabezaMontura
            '¿Esta muerto?
            If .flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡Estás muerto!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            '¿Tiene algun montura?
            If Not .Stats.NUMMONTURAS = 0 Then
                'Si tiene algun montura... ¿Ya tiene esa montura?
                For i = 1 To .Stats.NUMMONTURAS
                    If .flags.Montura(i).tipo = ObjData(.Invent.Object(Slot).ObjIndex).IndiceMontura Then
                        Call WriteConsoleMsg(UserIndex, "No puedes tener mas monturas de ese tipo.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                Next i
            End If
            
            '¿Tiene ya el maximo de monturas?
            If .Stats.NUMMONTURAS = MAX_MONTURAS Then
                Call WriteConsoleMsg(UserIndex, "No puedes tener mas monturas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Dim CabezaMontura As Obj
            CabezaMontura.Amount = 1
            CabezaMontura.ObjIndex = Obj.Cabeza
            
            If Not MeterItemEnInventario(UserIndex, CabezaMontura) Then _
                Call TirarItemAlPiso(UserList(UserIndex).Pos, CabezaMontura)

            'Comenzamos con el aprendizaje:
            
            'Sumamos + 1 el numero de monturas que tenemos
            .Stats.NUMMONTURAS = .Stats.NUMMONTURAS + 1
            'Ponemos el indice de la montura que estamos aprendiendo.
            .flags.Montura(.Stats.NUMMONTURAS).tipo = ObjData(.Invent.Object(Slot).ObjIndex).IndiceMontura
            'Le ponemos un nombre predeterminado
            .flags.Montura(.Stats.NUMMONTURAS).nombre = ObjData(.Invent.Object(Slot).ObjIndex).Name
            'Le ponemos nivel 1
            .flags.Montura(.Stats.NUMMONTURAS).MonturaLevel = 1
            'Le ponemos el ELU inicial y la exp
            .flags.Montura(.Stats.NUMMONTURAS).ELU = ELU_INICIAL
            .flags.Montura(.Stats.NUMMONTURAS).Exp = 0
            'Le ponemos Skills a 0
            .flags.Montura(.Stats.NUMMONTURAS).Skills = 0
            'Le ponemos el resto de atributos a 0
            .flags.Montura(.Stats.NUMMONTURAS).Ataque = 0
            .flags.Montura(.Stats.NUMMONTURAS).Defensa = 0
            .flags.Montura(.Stats.NUMMONTURAS).AtMagia = 0
            .flags.Montura(.Stats.NUMMONTURAS).DefMagia = 0
            .flags.Montura(.Stats.NUMMONTURAS).Evasion = 0
            .flags.Montura(.Stats.NUMMONTURAS).Speed = 0
            'Notificamos al usuario
            Call WriteConsoleMsg(UserIndex, "¡Felicidades, ahora tienes una " & ObjData(.Invent.Object(Slot).ObjIndex).Name & "!.", FontTypeNames.FONTTYPE_INFO)
            'Quitamos la cabeza del inventario
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)
            
            Dim query As String
            Dim fecha_creacion As String
            fecha_creacion = Ahora()
            
            query = "INSERT INTO montura (id_personaje, tipo, nombre, level, skills, elu, exp, ataque, defensa, atmagia, defmagia, evasion, fecha_creacion, speed) VALUE ('" & .id & "', '" & .flags.Montura(.Stats.NUMMONTURAS).tipo & "', '" & .flags.Montura(.Stats.NUMMONTURAS).nombre & "', '" & .flags.Montura(.Stats.NUMMONTURAS).MonturaLevel & "', '" & .flags.Montura(.Stats.NUMMONTURAS).Skills & "', '" & .flags.Montura(.Stats.NUMMONTURAS).ELU & "', '" & .flags.Montura(.Stats.NUMMONTURAS).Exp & "', '" & .flags.Montura(.Stats.NUMMONTURAS).Ataque & "', '" & .flags.Montura(.Stats.NUMMONTURAS).Defensa & "', '" & .flags.Montura(.Stats.NUMMONTURAS).AtMagia & "', '" & .flags.Montura(.Stats.NUMMONTURAS).DefMagia & "', '" & .flags.Montura(.Stats.NUMMONTURAS).Evasion & "', '" & fecha_creacion & "', '" & .flags.Montura(.Stats.NUMMONTURAS).Speed & "') "
            Set RS = SQL.Execute(query)
            
            Set RS = New ADODB.Recordset
            Set RS = SQL.Execute("SELECT id FROM montura WHERE id_personaje = '" & .id & "' AND fecha_creacion = '" & fecha_creacion & "'")
            
            .flags.Montura(.Stats.NUMMONTURAS).id = RS!id
            
            Call SaveUser(UserIndex)
            
    Case eOBJType.otManual
            '¿Esta muerto?
            If .flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡Estás muerto!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .Stats.UserSkills(Obj.IndiceSkill) >= Obj.CuantosSkill Then
                Call WriteConsoleMsg(UserIndex, "¡Tus conocimientos son superiores a los de este manual!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If Not .Stats.UserSkills(Obj.IndiceSkill) >= Obj.SkNecesarios Then
                Call WriteConsoleMsg(UserIndex, "¡No llegas a comprender este manual, necesitas tener " & Obj.SkNecesarios & " Skills para comprenderlo!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            .Stats.UserSkills(Obj.IndiceSkill) = Obj.CuantosSkill
            Call WriteConsoleMsg(UserIndex, "¡Tus conocimientos en " & SkillsNames(Obj.IndiceSkill) & " aumentaron en " & Obj.CuantosSkill & " puntos!", FontTypeNames.FONTTYPE_INFOBOLD)
            
            'Quitamos el manual del inventario
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)
            
    Case eOBJType.otCofre
        '01/03/2016 Lorwik: Cofres
            '¿Esta muerto?
            If .flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡Estás muerto!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Aqui va la futura comprobacion de si esta cerrado y tiene la llave
            
            If Obj.CantItems = 0 Then
                Call WriteConsoleMsg(UserIndex, "¡El cofre esta vacío!", FontTypeNames.FONTTYPE_INFO)
                'El cofre desaparece.
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
                Exit Sub
            End If
            
            Dim XProb As Byte
            Dim ObjCofr As Obj
            Dim Premios As Byte
            
            For i = 1 To Obj.CantItems
                'Tiramos los dados
                XProb = RandomNumber(1, 100)
            
                If XProb <= Obj.ItemCofre(i).Prob Then
                    MiObj.ObjIndex = Obj.ItemCofre(i).Obj
                    MiObj.Amount = Obj.ItemCofre(i).cant
                    
                    Premios = Premios + 1
                    
                    If Not MeterItemEnInventario(UserIndex, ObjCofr) Then
                        Call WriteConsoleMsg(UserIndex, "¡Has ganado " & ObjData(MiObj.ObjIndex).Name & "!", FontTypeNames.FONTTYPE_INFOBOLD)
                        Call TirarItemAlPiso(UserList(UserIndex).Pos, ObjCofr)
                    End If
                End If
            Next i
            
            'Finalmente se le quita el cofre
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)
            
            'Si le toco algo, se le dice cuantos items le toco.
            If Premios > 0 Then
                Call WriteConsoleMsg(UserIndex, "¡Encontraste " & Premios & " items en el cofre!", FontTypeNames.FONTTYPE_INFOBOLD)
            Else 'Si tuvo mala suerte se le dice.
                Call WriteConsoleMsg(UserIndex, "¡El cofre esta vacío!", FontTypeNames.FONTTYPE_INFO)
            End If
End Select

End With

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithWeapons(UserIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Call WriteCarpenterObjects(UserIndex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next

Dim Carcaj As Boolean

    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

    'Lorwik: Solo pierde el item de efectomagico=1
    If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
        If ObjData(UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.AnilloEqpSlot).ObjIndex).EfectoMagico = 1 Then
            Call DropObj(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot, 1, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            Exit Sub
        End If
    End If

    Call TirarTodosLosItems(UserIndex)

    Dim cantidad As Long
    If UserList(UserIndex).Stats.GLD > 20000 Then
        cantidad = UserList(UserIndex).Stats.GLD - 20000
        cantidad = Porcentaje(cantidad, 5)
        Call TirarOro(cantidad, UserIndex)
    End If
End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean

ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And _
            (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And _
            ObjData(Index).OBJType <> eOBJType.otLlaves And _
            ObjData(Index).OBJType <> eOBJType.otBarcos And _
            ObjData(Index).OBJType <> eOBJType.otBebidas And _
            ObjData(Index).OBJType <> eOBJType.otUseOnce And _
            ObjData(Index).NoSeCae = 0


End Function

Sub TirarItemsZonaPelea(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As worldPos
    Dim MiObj As Obj
   
    For i = 1 To MAX_INVENTORY_SLOTS
        With UserList(UserIndex)
        If .Invent.Object(i).ObjIndex Then
            If ItemSeCae(.Invent.Object(i).ObjIndex) Then
                MiObj.ObjIndex = .Invent.Object(i).ObjIndex
                MiObj.Amount = .Invent.Object(i).Amount
                
                NuevaPos.X = 0
                NuevaPos.Y = 0
          
                Tilelibre .Pos, NuevaPos, MiObj, False, True
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(UserIndex, i, MiObj.Amount, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
            End If
        End If
        End With
    Next i
End Sub

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As worldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    Dim Carcaj As Boolean
    
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
    
    '01/03/2016 Lorwik: Si tiene el Carcaj (efectomagico=2) solo pierde el 10% de las flechas.
    '¿Tiene anillo equipado?
    If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
        '¿El anillo equipado tiene efectomagico=2 (Carcaj)?
        If ObjData(UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.AnilloEqpSlot).ObjIndex).EfectoMagico = 2 Then
            Carcaj = True
        End If
    End If
    
    Debug.Print Carcaj & "-" & UserList(UserIndex).Invent.AnilloEqpSlot
    
    '02/03/2016 Lorwik: Desequipamos.
    Call Desequipar_Todo(UserIndex)
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                '02/03/2016 Lorwik: Si tiene el Carcaj solo tiramos el 10% de las flechas
                If Carcaj And ObjData(ItemIndex).OBJType = otFlechas Then
                    MiObj.Amount = Porcentaje(UserList(UserIndex).Invent.Object(i).Amount, 10)
                Else
                    MiObj.Amount = UserList(UserIndex).Invent.Object(i).Amount
                End If
                
                MiObj.ObjIndex = ItemIndex
                
                Debug.Print "Amount: " & MiObj.Amount
                
                'Pablo (ToxicWaste) 24/01/2007
                'Si es pirata y usa un Galeón entonces no explota los items. (en el agua)
                If UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, False, True
                Else
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
                End If
                
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then _
                    Call DropObj(UserIndex, i, MiObj.Amount, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
             End If
        End If
    Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
    ItemNewbie = ObjData(ItemIndex).Newbie = 1
End Function

Public Sub EquiparItemBonificador(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
'*******************************************************************************
'Lorwik
'22/03/2016
'Comprobamos si el Objeto aporta algun aumento de estadisticas y lo aplicamos.
'*******************************************************************************

    With UserList(UserIndex)
    
        If ObjData(ObjIndex).SumaVida > 0 Then _
            .flags.AumentodeVida = .flags.AumentodeVida + ObjData(ObjIndex).SumaVida
                
        If ObjData(ObjIndex).SumaMana > 0 Then _
            .flags.AumentodeMana = .flags.AumentodeMana + ObjData(ObjIndex).SumaMana
                    
        If ObjData(ObjIndex).SumaFuerza > 0 Then _
            .flags.AumentodeFuerza = .flags.AumentodeFuerza + ObjData(ObjIndex).SumaFuerza
                    
        If ObjData(ObjIndex).SumaAgilidad > 0 Then _
            .flags.AumentodeAgilidad = .flags.AumentodeAgilidad + ObjData(ObjIndex).SumaAgilidad
                
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
    End With
End Sub

Private Sub DesEquiparItemBonificador(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
'*******************************************************************************
'Lorwik
'22/03/2016
'Quitamos las bonificaciones de aumento de estadisticas.
'*******************************************************************************

Dim Resta As Integer

    With UserList(UserIndex)
        If ObjData(ObjIndex).SumaVida > 0 Then
            Resta = .Stats.MinHP - (VidaMaxima(UserIndex) - ObjData(ObjIndex).SumaVida)
            If Resta < 1 Then Resta = 0
                .flags.AumentodeVida = .flags.AumentodeVida - ObjData(ObjIndex).SumaVida
                .Stats.MinHP = .Stats.MinHP - Resta
                Call WriteUpdateUserStats(UserIndex)
            End If
            
            If ObjData(ObjIndex).SumaMana > 0 Then
                Resta = .Stats.MinMAN - (ManaMaxima(UserIndex) - ObjData(ObjIndex).SumaMana)
                If Resta < 1 Then Resta = 0
                .flags.AumentodeMana = .flags.AumentodeMana - ObjData(ObjIndex).SumaMana
                .Stats.MinMAN = .Stats.MinMAN - Resta
                Call WriteUpdateUserStats(UserIndex)
            End If

            If ObjData(ObjIndex).SumaFuerza > 0 Then .flags.AumentodeFuerza = .flags.AumentodeFuerza - ObjData(ObjIndex).SumaFuerza: Call WriteUpdateStrenght(UserIndex)
            If ObjData(ObjIndex).SumaAgilidad > 0 Then .flags.AumentodeAgilidad = .flags.AumentodeAgilidad - ObjData(ObjIndex).SumaAgilidad: Call WriteUpdateDexterity(UserIndex)
    End With
End Sub

