Attribute VB_Name = "Invusuario"
Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error GoTo fallo

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If ObjData(ObjIndex).OBJType <> OBJTYPE_LLAVES Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i

Exit Function
fallo:
Call LogError("TIENEOBJETOSROBABLES" & Err.number & " D: " & Err.Description)

End Function
Public Function ObjetosConMana(ByVal UserIndex As Integer) As Integer


On Error GoTo fallo

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    
    If ObjIndex > 0 Then
    
        If UserList(UserIndex).Invent.Object(i).Equipped > 0 Then
    
    If ObjData(ObjIndex).objetoespecial = 8 Then ObjetosConMana = ObjetosConMana + 100
    If ObjData(ObjIndex).objetoespecial = 9 Then ObjetosConMana = ObjetosConMana + 200
    If ObjData(ObjIndex).objetoespecial = 10 Then ObjetosConMana = ObjetosConMana + 300
      'pluto:7.0
    If ObjData(ObjIndex).objetoespecial = 17 Then ObjetosConMana = ObjetosConMana + 200
    If ObjData(ObjIndex).objetoespecial = 19 Then ObjetosConMana = ObjetosConMana + 55

        End If
        
    
    End If
Next i

Exit Function
fallo:
Call LogError("OBjetosconMana" & Err.number & " D: " & Err.Description)

End Function
Sub CambiarGemas(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim dar As Integer
Dim clase As String
Dim raza As String
Dim Genero As String

'Dim alli As Byte
clase = UCase$(UserList(UserIndex).clase)
raza = UCase$(UserList(UserIndex).raza)
Genero = UCase$(UserList(UserIndex).Genero)

If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then
'PLUTO:6.8 AÑADE CLASES PARA TUNICAS
   If clase = "MAGO" Or clase = "DRUIDA" Or clase = "BARDO" Then
    dar = 626
   Else
    dar = 592
  End If
Else 'raza
   If clase = "MAGO" Or clase = "DRUIDA" Or clase = "BARDO" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 619
            Case "MUJER"
            dar = 619
         End Select 'GENERO
  Else
    Select Case Genero
            Case "HOMBRE"
            dar = 590
            Case "MUJER"
            dar = 591
       End Select 'GENERO
   End If 'CLASE
End If 'RAZA

Dim MiObj As obj
MiObj.Amount = 1
MiObj.ObjIndex = dar
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
Call SendData(ToIndex, UserIndex, 0, "||6°Enhorabuena, te has ganado esta Armadura Dragón que no se cae.!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

Exit Sub
fallo:
Call LogError("CAMBIARGEMAS" & Err.number & " D: " & Err.Description)

End Sub
Sub CambiarGriaL(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim dar As Integer
Dim clase As String
Dim raza As String
Dim Genero As String

'Dim alli As Byte
clase = UCase$(UserList(UserIndex).clase)
raza = UCase$(UserList(UserIndex).raza)
Genero = UCase$(UserList(UserIndex).Genero)

If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then
             
        Select Case Genero
            Case "HOMBRE"
            dar = 943
            Case "MUJER"
            dar = 944
         End Select 'GENERO
        'pluto:6.0A
    If clase = "MAGO" Or clase = "DRUIDA" Or clase = "BARDO" Then
    dar = 1217
    End If
         
         
         
         
Else 'raza
  
Select Case Genero
            Case "HOMBRE"
            dar = 941
            Case "MUJER"
            dar = 942
       End Select 'GENERO
     'pluto:6.0A
     If clase = "MAGO" Or clase = "DRUIDA" Or clase = "BARDO" Then
    dar = 1216
    End If
 
 
End If 'RAZA

Dim MiObj As obj
MiObj.Amount = 1
MiObj.ObjIndex = dar
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
Call SendData(ToIndex, UserIndex, 0, "||6°Enhorabuena, te has ganado esta Armadura Legendaria que no se cae.!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

Exit Sub
fallo:
Call LogError("CAMBIARLEGENDARIAS" & Err.number & " D: " & Err.Description)

End Sub
Sub CambiarBola(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim dar As Integer
Dim clase As String
Dim raza As String
Dim Genero As String

'Dim alli As Byte
clase = UCase$(UserList(UserIndex).clase)
raza = UCase$(UserList(UserIndex).raza)
Genero = UCase$(UserList(UserIndex).Genero)

If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then
             
        Select Case Genero
            Case "HOMBRE"
            dar = 1012
            Case "MUJER"
            dar = 1012
         End Select 'GENERO
        'pluto:6.0A
    If clase = "MAGO" Or clase = "DRUIDA" Or clase = "BARDO" Then
    dar = 1291
    End If
         
         
         
         
Else 'raza
  
Select Case Genero
            Case "HOMBRE"
            dar = 1011
            Case "MUJER"
            dar = 1011
       End Select 'GENERO
     'pluto:6.0A
     If clase = "MAGO" Or clase = "DRUIDA" Or clase = "BARDO" Then
    dar = 1292
    End If
 
 
End If 'RAZA

Dim MiObj As obj
MiObj.Amount = 1
MiObj.ObjIndex = dar
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
Call SendData(ToIndex, UserIndex, 0, "||6°Enhorabuena, te has ganado esta Armadura del Caballero de la Muerte que no se cae.!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

Exit Sub
fallo:
Call LogError("CAMBIARLEGENDARIAS" & Err.number & " D: " & Err.Description)

End Sub

Sub CambiarTrofeo(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim dar As Integer
Dim clase As String
Dim raza As String
Dim Genero As String

'Dim alli As Byte
clase = UCase$(UserList(UserIndex).clase)
raza = UCase$(UserList(UserIndex).raza)
Genero = UCase$(UserList(UserIndex).Genero)

If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then
             
        Select Case Genero
            Case "HOMBRE"
            dar = 963
            Case "MUJER"
            dar = 963
         End Select 'GENERO
        'pluto:6.0A
    If clase = "MAGO" Or clase = "DRUIDA" Or clase = "BARDO" Then
    dar = 963
    End If
         
         
         
         
Else 'raza
  
Select Case Genero
            Case "HOMBRE"
            dar = 963
            Case "MUJER"
            dar = 963
       End Select 'GENERO
     'pluto:6.0A
     If clase = "MAGO" Or clase = "DRUIDA" Or clase = "BARDO" Then
    dar = 963
    End If
 
 
End If 'RAZA

Dim MiObj As obj
MiObj.Amount = 30
MiObj.ObjIndex = dar
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
Call SendData(ToIndex, UserIndex, 0, "||6°Enhorabuena, te has ganado un premio de ganador.!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

Exit Sub
fallo:
Call LogError("CAMBIARLEGENDARIAS" & Err.number & " D: " & Err.Description)

End Sub

Sub CambiarTrofeo2(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim dar As Integer
Dim clase As String
Dim raza As String
Dim Genero As String

'Dim alli As Byte
clase = UCase$(UserList(UserIndex).clase)
raza = UCase$(UserList(UserIndex).raza)
Genero = UCase$(UserList(UserIndex).Genero)

If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then
             
        Select Case Genero
            Case "HOMBRE"
            dar = 1245
            Case "MUJER"
            dar = 1245
         End Select 'GENERO
        'pluto:6.0A
    If clase = "MAGO" Or clase = "DRUIDA" Or clase = "BARDO" Then
    dar = 1245
    End If
         
         
         
         
Else 'raza
  
Select Case Genero
            Case "HOMBRE"
            dar = 1245
            Case "MUJER"
            dar = 1245
       End Select 'GENERO
     'pluto:6.0A
     If clase = "MAGO" Or clase = "DRUIDA" Or clase = "BARDO" Then
    dar = 1245
    End If
 
 
End If 'RAZA

Dim MiObj As obj
MiObj.Amount = 1
MiObj.ObjIndex = dar
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
Call SendData(ToIndex, UserIndex, 0, "||6°Enhorabuena, te has ganado 1 Trofeo de Primer puesto.!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

Exit Sub
fallo:
Call LogError("CAMBIARLEGENDARIAS" & Err.number & " D: " & Err.Description)

End Sub
Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador
If UserList(UserIndex).flags.Privilegios > 0 Then
    ClasePuedeUsarItem = True
    Exit Function
End If
'pluto:2.15
If UserList(UserIndex).Bebe > 0 And ObjIndex <> 460 Then
ClasePuedeUsarItem = False
Exit Function
End If
'----------


Dim flag As Boolean

If ObjData(ObjIndex).ClaseProhibida(1) <> "" Then
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(ObjIndex).ClaseProhibida(i) = UCase$(UserList(UserIndex).clase) Then
                ClasePuedeUsarItem = False
                Exit Function
        End If
    Next i
    
Else
End If
ClasePuedeUsarItem = True
Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, UserList(UserIndex).Invent.Object(j).Amount)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next


Exit Sub
fallo:
Call LogError("QUITARNEWBIEOBJ" & Err.number & " D: " & Err.Description)

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)

On Error GoTo fallo
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).ObjIndex = 0
        UserList(UserIndex).Invent.Object(j).Amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
        
Next

UserList(UserIndex).Invent.NroItems = 0

UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0

UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0

UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0
'[GAU]
UserList(UserIndex).Invent.BotaEqpObjIndex = 0
UserList(UserIndex).Invent.BotaEqpSlot = 0
'[GAU]
'pluto:2.4
UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
UserList(UserIndex).Invent.AnilloEqpSlot = 0

UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0

UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
UserList(UserIndex).Invent.HerramientaEqpSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0
Exit Sub
fallo:
Call LogError("LIMPIARINVENTARIO" & Err.number & " D: " & Err.Description)

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
On Error GoTo fallo
'PLUTO:6.2
If UserList(UserIndex).Pos.Map = 191 Or UserList(UserIndex).Pos.Map = 293 Or UserList(UserIndex).Pos.Map = MapaTorneo2 Then Exit Sub

If Cantidad > 100000 Then Exit Sub
If UserList(UserIndex).flags.Privilegios > 0 And UserList(UserIndex).flags.Privilegios < 3 Then Exit Sub

'SI EL NPC TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As obj
        'info debug
        Dim loops As Integer
        Do While (Cantidad > 0) And (UserList(UserIndex).Stats.GLD > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            
            If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
            
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            Call LogCasino("Usuario tira oro: " & UserList(UserIndex).Name & " IP: " & UserList(UserIndex).ip & " Nom: " & " MAPA: " & UserList(UserIndex).Pos.Map)
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
    
End If

Exit Sub

Exit Sub
fallo:
Call LogError("TIRARORO" & Err.number & " D: " & Err.Description)


End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
On Error GoTo fallo
Dim MiObj As obj
'Desequipa
If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)

'Quita un objeto
UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - Cantidad
'pluto.2.3
UserList(UserIndex).Stats.Peso = UserList(UserIndex).Stats.Peso - (ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Peso * Cantidad)
'pluto:2.4.5
If UserList(UserIndex).Stats.Peso < 0.001 Then UserList(UserIndex).Stats.Peso = 0
Call SendUserStatsPeso(UserIndex)

'¿Quedan mas?
If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
    UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
    UserList(UserIndex).Invent.Object(Slot).Amount = 0
End If

 Exit Sub
fallo:
Call LogError("QUITARUSERINVENTARIO " & Err.number & " D: " & Err.Description)
  
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo fallo
Dim NullObj As UserOBJ
Dim loopc As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For loopc = 1 To MAX_INVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(loopc).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, loopc, UserList(UserIndex).Invent.Object(loopc))
        Else
            
            Call ChangeUserInv(UserIndex, loopc, NullObj)
            
        End If

    Next loopc

End If
Exit Sub
fallo:
Call LogError("UPDATEUSERINV" & Err.number & " D: " & Err.Description)

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo fallo

If UserList(UserIndex).flags.Privilegios > 0 Then GoTo sipuede

'PLUTO:6.2
If UserList(UserIndex).Pos.Map = 191 Or UserList(UserIndex).Pos.Map = Prision.Map Or UserList(UserIndex).Pos.Map = 293 Or UserList(UserIndex).Pos.Map = MapaTorneo2 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes soltar Objetos en este Mapa." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

Dim obj As obj

'pluto:2.17
If (ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Real = 1 Or ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Caos = 1) And UserList(UserIndex).Pos.Map <> 49 And UserList(UserIndex).Invent.Object(Slot).ObjIndex <> 1018 And UserList(UserIndex).Invent.Object(Slot).ObjIndex <> 1019 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes soltar la Ropa de Armadas" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

'pluto:6.7------------------------
If UserList(UserIndex).Invent.Object(Slot).ObjIndex = 1236 Or UserList(UserIndex).Invent.Object(Slot).ObjIndex = 1238 Or UserList(UserIndex).Invent.Object(Slot).ObjIndex = 1285 Or UserList(UserIndex).Invent.Object(Slot).ObjIndex = 1286 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes soltar la Perseus" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'--------------------------------
sipuede:
'pluto:2.14
If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = 42 And UserList(UserIndex).flags.Montura > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes soltar la Ropa mientrás Cabalgas." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
If num > 0 Then

  
  If num > UserList(UserIndex).Invent.Object(Slot).Amount Then num = UserList(UserIndex).Invent.Object(Slot).Amount
  
  'Check objeto en el suelo
            
        
    If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then

        If UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Demonio > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes desequipar estando transformado." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
        Call Desequipar(UserIndex, Slot)
    End If
obj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
obj.Amount = num

      
      
      
If MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex = 0 Then

                
        'If UserList(UserIndex).Flags.Privilegios > 0 And UserList(UserIndex).Flags.Privilegios < 3 Then
            'If ObjData(Obj.ObjIndex).Real = 0 And ObjData(Obj.ObjIndex).Caos = 0 _
            'And ObjData(Obj.ObjIndex).nocaer = 0 Or ObjData(Obj.ObjIndex).ObjType = 40 Then Exit Sub
       ' End If
'pluto:2.9.0
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType <> 60 Then
    Call MakeObj(ToMap, 0, Map, obj, Map, X, Y)
    If UserList(UserIndex).flags.Muerto = 0 Then UserList(UserIndex).ObjetosTirados = UserList(UserIndex).ObjetosTirados + 1
    If Alarma = 1 Then Call SendData(ToAdmins, UserIndex, 0, "||Tira Objeto: " & UserList(UserIndex).Name & " " & ObjData(obj.ObjIndex).Name & "´" & FontTypeNames.FONTTYPE_COMERCIO)
    End If
    
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = 60 And UserList(UserIndex).flags.Montura > 0 Then Exit Sub
If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = 60 Then
Dim xx As Integer
xx = obj.ObjIndex - 887
End If
    Call QuitarUserInvItem(UserIndex, Slot, num)
        Call UpdateUserInv(False, UserIndex, Slot)
        
        If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name)
  Else
'pluto:2.6.0
If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = 60 Then Exit Sub
'pluto:6.0A
If ObjData(MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex).OBJType = 6 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes soltar objetos en una puerta." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

Call SendData(ToIndex, UserIndex, 0, "M8")
Call TirarItemAlPiso(UserList(UserIndex).Pos, obj)
Call QuitarUserInvItem(UserIndex, Slot, num)
Call UpdateUserInv(False, UserIndex, Slot)
If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name)
'pluto:2.9.0
If UserList(UserIndex).flags.Muerto = 0 Then UserList(UserIndex).ObjetosTirados = UserList(UserIndex).ObjetosTirados + 1
If Alarma = 1 Then Call SendData(ToAdmins, UserIndex, 0, "||Tira Objeto: " & UserList(UserIndex).Name & " " & ObjData(obj.ObjIndex).Name & "´" & FontTypeNames.FONTTYPE_COMERCIO)

End If
    
End If
Exit Sub
fallo:
Call LogError("DROPOBJETO " & Err.number & " D: " & Err.Description)

End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo fallo
MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount - num

If MapData(Map, X, Y).OBJInfo.Amount <= 0 Then
    MapData(Map, X, Y).OBJInfo.ObjIndex = 0
    MapData(Map, X, Y).OBJInfo.Amount = 0
    'pluto:2.3----------
    'If sndRoute = 2 Then
    'Call SendToAreaByPos(Map, X, Y, "BO" & X & "," & Y)
    'Else
    Call SendData(sndRoute, sndIndex, sndMap, "BO" & X & "," & Y)
    'End If
    '--------------------

End If
Exit Sub
fallo:
Call LogError("ERASE OBJETO " & Err.number & " D: " & Err.Description)

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, obj As obj, Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo fallo
'Crea un Objeto
If obj.ObjIndex = 0 Then Exit Sub
'pluto:2.15
If ObjData(obj.ObjIndex).OBJType = 77 Then
Dim roda As Byte
roda = RandomNumber(1, 6)
Call SendData(sndRoute, sndIndex, sndMap, "HU" & ObjData(obj.ObjIndex).GrhIndex & "," & X & "," & Y & "," & roda)
Exit Sub
End If
'------------------------
MapData(Map, X, Y).OBJInfo = obj
Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(obj.ObjIndex).GrhIndex & "," & X & "," & Y)
Exit Sub
fallo:
Call LogError("MAKEOBJ " & Err.number & " D: " & Err.Description)

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As obj) As Boolean
On Error GoTo fallo

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
'           Call SendData(ToIndex, UserIndex, 0, "||No podes cargar mas objetos." & FONTTYPENAMES.FONTTYPE_fight)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If

'nati:Compruebo que solo se haga esta acción en > QUEST!
If UserList(UserIndex).Mision.estado > 0 Then
If UserList(UserIndex).raza = "Enano" Or UserList(UserIndex).raza = "Gnomo" Or UserList(UserIndex).raza = "Goblin" Then
If ComprobarObjetivos(UserIndex) = True Then
If Not MiObj.ObjIndex2 = 0 Then
MiObj.ObjIndex = MiObj.ObjIndex2
MiObj.Amount = MiObj.Amount2
End If
End If
End If
End If
'nati:Compruebo que solo se haga esta acción en > QUEST!
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True

'pluto.2.3
UserList(UserIndex).Stats.Peso = UserList(UserIndex).Stats.Peso + (ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Peso * MiObj.Amount)
Call SendUserStatsPeso(UserIndex)
       
Call UpdateUserInv(False, UserIndex, Slot)
'Debug.Print UserList(UserIndex).Invent.Object(11).Amount
Exit Function
fallo:
Call LogError("METERITEMINVENTARIO " & UserList(UserIndex).Name & " Obj: " & MiObj.ObjIndex & " C: " & MiObj.Amount & " " & Err.number & " D: " & Err.Description)

End Function


Sub GetObj(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim obj As ObjData
Dim MiObj As obj

'¿Hay algun obj?
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex > 0 Then
  UserList(UserIndex).ObjetosTirados = 0
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        obj = ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex

'pluto:2.4.5
If (ObjData(MiObj.ObjIndex).Peso * MiObj.Amount) + UserList(UserIndex).Stats.Peso > UserList(UserIndex).Stats.PesoMax Then
Dim pd, vc As Integer
pd = UserList(UserIndex).Stats.PesoMax - UserList(UserIndex).Stats.Peso
'pluto:6.5

If pd < 0 Then GoTo lala

vc = Int(pd / ObjData(MiObj.ObjIndex).Peso)
lala:
Call SendData(ToIndex, UserIndex, 0, "||Demasiada Carga." & "´" & FontTypeNames.FONTTYPE_info)
If vc < 1 Then Exit Sub
MiObj.Amount = vc
Else
vc = MiObj.Amount
End If
        
   
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call SendData(ToIndex, UserIndex, 0, "P5")
        Else
            'Quitamos el objeto
           
            Call EraseObj(ToMap, 0, UserList(UserIndex).Pos.Map, vc, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & vc & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
        End If
        
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "M9")
End If
Exit Sub
fallo:
Call LogError("GETOBJ " & Err.number & " D: " & Err.Description & "->" & UserList(UserIndex).Name & " Obj:" & MiObj.ObjIndex & " Cant:" & MiObj.Amount)

End Sub
Sub GetObjFantasma(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
On Error GoTo fallo
Dim obj As ObjData
Dim MiObj As obj

'¿Hay algun obj?
If MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex > 0 Then
  UserList(UserIndex).ObjetosTirados = 0
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex).Agarrable <> 1 Then
        
    Dim Slot As Byte
        
       
        obj = ObjData(MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex

'pluto:2.4.5
If (ObjData(MiObj.ObjIndex).Peso * MiObj.Amount) + UserList(UserIndex).Stats.Peso > UserList(UserIndex).Stats.PesoMax Then
Dim pd, vc As Integer
pd = UserList(UserIndex).Stats.PesoMax - UserList(UserIndex).Stats.Peso
vc = Int(pd / ObjData(MiObj.ObjIndex).Peso)
Call SendData(ToIndex, UserIndex, 0, "||Demasiada Carga." & "´" & FontTypeNames.FONTTYPE_info)
If vc < 1 Then Exit Sub
MiObj.Amount = vc
Else
vc = MiObj.Amount
End If
        
   
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call SendData(ToIndex, UserIndex, 0, "P5")
        Else
            'Quitamos el objeto
           
            Call EraseObj(ToMap, 0, UserList(UserIndex).Pos.Map, vc, UserList(UserIndex).Pos.Map, X, Y)
            If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & vc & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
        End If
        
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "M9")
End If
Exit Sub
fallo:
Call LogError("GETOBJ " & Err.number & " D: " & Err.Description & "->" & UserList(UserIndex).Name & " Obj:" & MiObj.ObjIndex & " Cant:" & MiObj.Amount)

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo fallo
'PLUTO:2.4.2
If UserList(UserIndex).Pos.Map = 191 Then Exit Sub


'Desequipa el item slot del inventario
Dim obj As ObjData
If UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Demonio > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes desequipar estando transformado." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
If Slot = 0 Then Exit Sub
If UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

    Select Case obj.OBJType
        Case OBJTYPE_WEAPON
        'objeto especial
        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 2 Then
            UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 5
        End If
        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 3 Then
            UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 2
        End If
        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 4 Then
            UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 3
        End If
        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 8 Then
            UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 100
        End If
        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 9 Then
            UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 200
        End If
        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 10 Then
            UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 300
        End If
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        '[GAU] Agregamo UserList(UserIndex).Char.Botas
        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)

    Case OBJTYPE_FLECHAS
    
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
    Case OBJTYPE_HERRAMIENTAS
    
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
        UserList(UserIndex).Invent.HerramientaEqpSlot = 0
'pluto:2.4
Case OBJTYPE_Anillo

If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).SubTipo = 1 Then
Call SendData(ToIndex, UserIndex, 0, "E3")
  UserList(UserIndex).flags.Oculto = 0
  UserList(UserIndex).Counters.Invisibilidad = 0
  UserList(UserIndex).flags.Invisible = 0
  Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 17, UserList(UserIndex).Char.CharIndex & ",0")
End If
'pluto:2.4
If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).SubTipo = 5 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax - 500
Call SendUserStatsPeso(UserIndex)
End If


UserList(UserIndex).Invent.Object(Slot).Equipped = 0
UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
UserList(UserIndex).Invent.AnilloEqpSlot = 0

    
    
    Case OBJTYPE_ARMOUR
        
    Select Case obj.SubTipo
        Case OBJTYPE_ARMADURA
     
Select Case ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).objetoespecial
Case 2
UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 5
Case 3
 UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 2
Case 4
 UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 3
Case 5
UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 5
Case 6
 UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 2
 Case 7
UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 3
Case 8
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 100
Case 9
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 200
Case 10
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 300
'pluto:6.5----------
Case 14
 UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 5
 UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 2
'------------------
'pluto:7.0--------------------
Case 16
UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 1
UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 1
Case 17
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 200
UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 2
Case 18
UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 1
Case 19
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 55

'-------------------------------------
End Select
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 0
            UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
            UserList(UserIndex).Invent.ArmourEqpSlot = 0
                
            If UserList(UserIndex).flags.Montura <> 1 Then Call DarCuerpoDesnudo(UserIndex)
            '[GAU] Agregamo UserList(UserIndex).Char.Botas
            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
                
        Case OBJTYPE_CASCO
            'objeto especial
        Select Case ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial
            Case 5
             UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 5
            Case 6
             UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 2
            Case 7
            UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 3
            Case 8
            UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 100
            Case 9
            UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 200
            Case 10
             UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN - 300
            Case 18
             UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 1
        End Select
           
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 0
            UserList(UserIndex).Invent.CascoEqpObjIndex = 0
            UserList(UserIndex).Invent.CascoEqpSlot = 0
            UserList(UserIndex).Char.CascoAnim = NingunCasco
             '[GAU] Agregamo UserList(UserIndex).Char.Botas
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
            '[GAU] AGREGAR TODO ESTO!!!
            Case OBJTYPE_BOTA
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.BotaEqpObjIndex = 0
                UserList(UserIndex).Invent.BotaEqpSlot = 0
                UserList(UserIndex).Char.Botas = NingunBota
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
            '[GAU] HASTA AK
            Case OBJTYPE_ESCUDO
                
                'objeto especial
                If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 5 Then
                    UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 5
                End If
                If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 6 Then
                    UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 2
                End If
                If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 7 Then
                    UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 3
                End If
                If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 2 Then
                    UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 5
                End If
                If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 3 Then
                    UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 2
                End If
                If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 4 Then
                    UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 3
                End If
                
                'pluto:6.5
                 If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 12 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 1
                 UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 3
                 End If
                  If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 13 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 2
                 UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 2
                 End If
                   If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 14 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 5
                 UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 2
                 End If
                   If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).objetoespecial = 15 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) - 3
                 UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 2
                 End If
                '-----------------
                
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
                UserList(UserIndex).Invent.EscudoEqpSlot = 0
                UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                '[GAU] Agregamo UserList(UserIndex).Char.Botas
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
                End Select
            End Select
If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
'pluto:evita quitar dope arqueros
 'if obj.OBJType = OBJTYPE_FLECHAS And (UCase$(UserList(UserIndex).clase) = "ARQUERO" Or UCase$(UserList(UserIndex).clase) = "CAZADOR") Then GoTo alli9
'anula efecto pociones
 
 'pluto:6.5 objetos que modifican atributos anulamos efecto pociones
 If obj.objetoespecial > 1 Then
 Dim loopX As Integer
  For loopX = 1 To NUMATRIBUTOS
UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
 Next
 End If
 
 
alli9:
Call SendUserStatsMana(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)
Exit Sub
fallo:
Call LogError("DESEQUIPAR " & Err.number & " D: " & Err.Description & " Nombre: " & UserList(UserIndex).Name & " Obj: " & obj.Name & " Slot: " & Slot)

End Sub
Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo errhandler
If UserList(UserIndex).flags.Privilegios > 0 Then
SexoPuedeUsarItem = True
Exit Function
End If
If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(UserIndex).Genero) <> "HOMBRE"
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(UserIndex).Genero) <> "MUJER"
Else
    SexoPuedeUsarItem = True
End If

Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function
Function SkillsPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo fallo
SkillsPuedeUsarItem = False

If UserList(UserIndex).flags.Privilegios > 0 Then
SkillsPuedeUsarItem = True
Exit Function
End If

If ObjData(ObjIndex).proyectil > 0 And UserList(UserIndex).Stats.UserSkills(RequeProyec) >= ObjData(ObjIndex).SkArco Then SkillsPuedeUsarItem = True
If ObjData(ObjIndex).proyectil = 0 And UserList(UserIndex).Stats.UserSkills(RequeArma) >= ObjData(ObjIndex).SkArma Then SkillsPuedeUsarItem = True
 
Exit Function
fallo:
Call LogError("skillspuedeusaritem" & Err.number & " D: " & Err.Description)

End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo fallo


If UserList(UserIndex).flags.Privilegios > 0 Then
FaccionPuedeUsarItem = True
Exit Function
End If

If ObjData(ObjIndex).Real > 0 Then
    If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
    FaccionPuedeUsarItem = False
    Else
    FaccionPuedeUsarItem = True
    End If
ElseIf ObjData(ObjIndex).Caos > 0 Then
    If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
    FaccionPuedeUsarItem = False
    Else
    FaccionPuedeUsarItem = True
    End If
Else
   FaccionPuedeUsarItem = True
End If


Exit Function
fallo:
Call LogError("FACCIONPUEDEUSARITEM" & Err.number & " D: " & Err.Description)

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo errhandler
'PLUTO:2.4.2
If UserList(UserIndex).Pos.Map = 191 Then Exit Sub


If UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Demonio > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes equipar estando transformado." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'Equipa un item del inventario
Dim obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
obj = ObjData(ObjIndex)
If UserList(UserIndex).flags.Privilegios > 0 Then
GoTo sipuede
End If

If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los newbies pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If
 'pluto:2.10
If obj.ObjetoClan <> "" Then
If UCase$(UserList(UserIndex).GuildInfo.GuildName) <> UCase$(obj.ObjetoClan) Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes equipar Ropa de ese Clan" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
End If
 
 'comprueba si es elfo
 If obj.razaelfa = 1 And UserList(UserIndex).raza <> "Elfo" And UserList(UserIndex).raza <> "Elfo Oscuro" Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los Elfos pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If
 'comprueba si es vampiro
 If obj.razavampiro = 1 And UserList(UserIndex).raza <> "Vampiro" Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los Vampiros pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If
 'comprueba si es humano
 If obj.razahumana = 1 And UserList(UserIndex).raza <> "Humano" Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los Humanos pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If
 'comprueba si es orco
 If obj.razaorca = 1 And UserList(UserIndex).raza <> "Orco" Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los Orcos pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If
 'comprueba si es enano
 'pluto:7.0 añado goblin
 If obj.RazaEnana = 1 And UserList(UserIndex).raza <> "Enano" And UserList(UserIndex).raza <> "Gnomo" And UserList(UserIndex).raza <> "Goblin" Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los Enanos y Gnomos pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If
 
 If obj.Caos > 1 And UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
 Call SendData(ToIndex, UserIndex, 0, "||Sólo miembros de las Fuerzas del Caos pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
 Exit Sub
 End If
 
 If obj.Real > 0 And UserList(UserIndex).Faccion.ArmadaReal <> 1 Then
Call SendData(ToIndex, UserIndex, 0, "||Sólo los miembros de la Armada Real pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
 Exit Sub
End If
'solo la legion
'If obj.Real = 2 And UserList(UserIndex).Faccion.ArmadaReal <> 2 Then
'Call SendData(ToIndex, UserIndex, 0, "||Sólo los miembros de la Legión pueden usar este objeto." & FONTTYPENAMES.FONTTYPE_INFO)
' Exit Sub
'End If

 
 
 'eliminamos pociones
 'Dim loopX As Integer
    'For loopX = 1 To NUMATRIBUTOS
    'UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
   ' Next

sipuede:
 'pluto:2.17-------------------
If UserList(UserIndex).Invent.EscudoEqpObjIndex = 0 Then GoTo n
'If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).SubTipo = 6 Or ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).SubTipo = 7 Then
If (obj.SubTipo = 6 Or obj.SubTipo = 7) And ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).SubTipo = 2 And ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).OBJType = 3 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes usar Armas de dos Manos con Escudo." & "´" & FontTypeNames.FONTTYPE_info)
 Exit Sub
End If
n:
 
 If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then GoTo n1
If obj.SubTipo = 2 And obj.OBJType = 3 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo = 6 Or ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo = 7) Then
 'If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo = 6 Or ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo = 7 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes usar Escudo con Armas de dos Manos." & "´" & FontTypeNames.FONTTYPE_info)
 Exit Sub
End If
'End If
n1:
'-----------------------------------------------
Select Case obj.OBJType
    Case OBJTYPE_WEAPON
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) And _
          SkillsPuedeUsarItem(UserIndex, ObjIndex) Then
                'Si esta equipado lo quITA
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    'Animacion por defecto
                    UserList(UserIndex).Char.WeaponAnim = NingunArma
                    '[GAU] Agregamo UserList(UserIndex).Char.Botas
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.WeaponEqpSlot = Slot
                'añade objeto especial

Select Case ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).objetoespecial
Case 2
UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 5
Case 3
UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 2
Case 4
UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 3
Case 8
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 100
Call SendUserStatsMana(UserIndex)
Case 9
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 200
Call SendUserStatsMana(UserIndex)
Case 10
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 300
Call SendUserStatsMana(UserIndex)
End Select
                 'Sonido
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SACARARMA)
        
                UserList(UserIndex).Char.WeaponAnim = obj.WeaponAnim
                '[GAU] Agregamo UserList(UserIndex).Char.Botas
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
                Else
            Call SendData(ToIndex, UserIndex, 0, "J4")
            'Call SendData(ToIndex, UserIndex, 0, "||No puede usar este objeto." & FONTTYPENAMES.FONTTYPE_INFO)
       End If
    Case OBJTYPE_HERRAMIENTAS
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) And _
          SkillsPuedeUsarItem(UserIndex, ObjIndex) Then
               
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.HerramientaEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.HerramientaEqpSlot = Slot
                
       Else
            'Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPENAMES.FONTTYPE_INFO)
       Call SendData(ToIndex, UserIndex, 0, "J4")
       End If
    
     'pluto:2.4
    Case OBJTYPE_Anillo
    If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
         FaccionPuedeUsarItem(UserIndex, ObjIndex) And _
          SkillsPuedeUsarItem(UserIndex, ObjIndex) Then
             
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.AnilloEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.AnilloEqpSlot = Slot

'pluto:2.4
If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).SubTipo = 1 Then
'pluto:2.11
If UserList(UserIndex).flags.Angel = 0 And UserList(UserIndex).flags.Demonio = 0 And UserList(UserIndex).flags.Morph = 0 And MapInfo(UserList(UserIndex).Pos.Map).Pk = True Then
UserList(UserIndex).flags.Invisible = 1
Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 17, UserList(UserIndex).Char.CharIndex & ",1")
End If
End If
'pluto:2.4
If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).SubTipo = 5 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + 500
Call SendUserStatsPeso(UserIndex)
End If

If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).SubTipo = 2 Then
If UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Angel = 1 Or UserList(UserIndex).flags.Demonio = 1 Or UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
   UserList(UserIndex).flags.Morph = UserList(UserIndex).Char.Body
   UserList(UserIndex).Counters.Morph = IntervaloMorphPJ
    Dim abody As Integer
    Dim al As Integer
    al = RandomNumber(1, 12)

Select Case al
Case 1
abody = 5
Case 2
abody = 6
Case 3
abody = 9
Case 4
abody = 10
Case 5
abody = 13
Case 6
abody = 42
Case 7
abody = 51
Case 8
abody = 59
Case 9
abody = 68
Case 10
abody = 71
Case 11
abody = 73
Case 12
abody = 88
End Select
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, val(abody), val(0), UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(43).FXgrh & "," & Hechizos(43).loops)

End If
       
       Else
            Call SendData(ToIndex, UserIndex, 0, "J4")
            'Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPENAMES.FONTTYPE_INFO)
       End If
    
    Case OBJTYPE_FLECHAS
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) And _
          SkillsPuedeUsarItem(UserIndex, ObjIndex) Then
             
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call SendData(ToIndex, UserIndex, 0, "J4")
            'Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPENAMES.FONTTYPE_INFO)
       End If
    
    Case OBJTYPE_ARMOUR
         
         If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
         
         Select Case obj.SubTipo
         
            Case OBJTYPE_ARMADURA
                'pluto:2.3
         If UserList(UserIndex).flags.Montura = 1 Then Exit Sub

                
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
                   SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
                   CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
                   FaccionPuedeUsarItem(UserIndex, ObjIndex) And _
                   SkillsPuedeUsarItem(UserIndex, ObjIndex) Then
                   
                   'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        Call DarCuerpoDesnudo(UserIndex)
                        '[GAU] Agregamo UserList(UserIndex).Char.Botas
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)

                    Exit Sub
                    End If
            
                    'Quita el anterior
                    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
                    End If
            
                    'Lo equipa
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.ArmourEqpSlot = Slot
                        
                    UserList(UserIndex).Char.Body = obj.Ropaje
                     'pluto:2-3-04
                     If UserList(UserIndex).Remort = 1 Then
                     If UserList(UserIndex).Char.Body = 196 Then UserList(UserIndex).Char.Body = 262
                     If UserList(UserIndex).Char.Body = 197 Then UserList(UserIndex).Char.Body = 263
                     End If
                    
                    UserList(UserIndex).flags.Desnudo = 0
                     'objeto especial
 Select Case ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).objetoespecial
Case 2
UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 5
Case 3
 UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 2
Case 4
 UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 3
Case 5
UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 5
Case 6
 UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 2
 Case 7
UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 3
Case 8
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 100
Call SendUserStatsMana(UserIndex)
Case 9
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 200
Call SendUserStatsMana(UserIndex)
Case 10
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 300
Call SendUserStatsMana(UserIndex)
'pluto:6.5----------------------
Case 14
UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 5
 UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 2
'-------------------------------
'pluto:7.0--------------------
Case 16
UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 1
UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 1
Case 17
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 200
UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 2
Case 18
UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 1
Case 19
UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 55

'-------------------------------------
End Select
                    ' If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).objetoespecial = 5 Then
                'UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 5
               ' End If
                      '   If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).objetoespecial = 6 Then
                'UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 2
                'End If
                    ' If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).objetoespecial = 7 Then
                'UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 3
               ' End If
                '    If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).objetoespecial = 8 Then
               ' UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 100
                'End If
               ' If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).objetoespecial = 9 Then
                '  UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 200
                 'End If
               ' If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).objetoespecial = 10 Then
                    'UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 300
           ' End If
            
                    
                    '[GAU] Agregamo UserList(UserIndex).Char.Botas
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)

                    Else
                    Call SendData(ToIndex, UserIndex, 0, "J4")
                    'Call SendData(ToIndex, UserIndex, 0, "||Tu clase,genero o raza no puede usar este objeto." & FONTTYPENAMES.FONTTYPE_INFO)
                End If
            Case OBJTYPE_CASCO
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                    'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        UserList(UserIndex).Char.CascoAnim = NingunCasco
                         '[GAU] Agregamo UserList(UserIndex).Char.Botas
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
                    Exit Sub
                    End If
            
                    'Quita el anterior
                    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
                    End If
            
                    'Lo equipa
                    
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.CascoEqpSlot = Slot
                    
                    UserList(UserIndex).Char.CascoAnim = obj.CascoAnim
                    
                    
                    'objeto especial
                        If ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).objetoespecial = 5 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 5
                End If
                         If ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).objetoespecial = 6 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 2
                End If
                     If ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).objetoespecial = 7 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 3
                End If
                    If ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).objetoespecial = 8 Then
                UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 100
                Call SendUserStatsMana(UserIndex)
                End If
                If ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).objetoespecial = 9 Then
                  UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 200
                 Call SendUserStatsMana(UserIndex)
                 End If
                If ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).objetoespecial = 10 Then
                    UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 300
                Call SendUserStatsMana(UserIndex)
                End If
                'pluto:7.0
                If ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).objetoespecial = 18 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 1
                End If
                    '[GAU] Agregamo UserList(UserIndex).Char.Botas
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
                    Else
                    Call SendData(ToIndex, UserIndex, 0, "J4")
                    'Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPENAMES.FONTTYPE_INFO)
                End If
            '[GAU] Agregar todo ESTO!!!!!
            Case OBJTYPE_BOTA
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                    'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        UserList(UserIndex).Char.Botas = NingunBota
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If UserList(UserIndex).Invent.BotaEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.BotaEqpSlot)
                    End If
            
                    'Lo equipa
                    
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.BotaEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.BotaEqpSlot = Slot
                    
                    UserList(UserIndex).Char.Botas = obj.Botas
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "J4")
                    'Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPENAMES.FONTTYPE_INFO)
                End If
                '[GAU] HASTA AK!!!!
            
            Case OBJTYPE_ESCUDO
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
       
                    'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                         '[GAU] Agregamo UserList(UserIndex).Char.Botas
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)

                    Exit Sub
                    End If
            
                    'Quita el anterior
                    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                    End If
            
                    'Lo equipa
                    
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                    UserList(UserIndex).Invent.EscudoEqpSlot = Slot
                    
                    UserList(UserIndex).Char.ShieldAnim = obj.ShieldAnim
                    'quitar esto
                    'UserList(UserIndex).Char.ShieldAnim = 33
                    'objeto especial
                            If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).objetoespecial = 5 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 5
                End If
                         If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).objetoespecial = 6 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 2
                End If
                     If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).objetoespecial = 7 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 3
                End If
                   If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).objetoespecial = 2 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 5
                End If
                If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).objetoespecial = 3 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 2
                End If
                If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).objetoespecial = 4 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 3
                End If
                
                'pluto:6.5---------
                If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).objetoespecial = 12 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 1
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 3
                End If
                    If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).objetoespecial = 13 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 2
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 2
                End If
                   If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).objetoespecial = 14 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 5
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 2
                End If
                         If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).objetoespecial = 15 Then
                UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 3
                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 2
                End If
                '----------------
                
                
                   '[GAU] Agregamo UserList(UserIndex).Char.Botas
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
 
                      Else
                    Call SendData(ToIndex, UserIndex, 0, "J4")
                    'Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPENAMES.FONTTYPE_INFO)
                End If
        End Select
End Select

'Actualiza
'Call UpdateUserInv(True, userindex, 0)
'pluto:2.4
Call UpdateUserInv(False, UserIndex, Slot)
'pluto:6.5
If ObjData(ObjIndex).objetoespecial > 1 Then
Dim loopX As Integer
    For loopX = 1 To NUMATRIBUTOS
    UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
    Next
End If


Exit Sub
errhandler:
Call LogError("EquiparInvItem Slot:" & Slot)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, itemIndex As Integer) As Boolean
On Error GoTo errhandler
'pluto:6.3 añado papa noel(1016)
If UserList(UserIndex).flags.Privilegios > 0 Or itemIndex = 1016 Then
CheckRazaUsaRopa = True
Exit Function
End If
'pluto.7.0 añade ciclope
'Verifica si la raza puede usar la ropa
If UserList(UserIndex).raza = "Humano" Or _
   UserList(UserIndex).raza = "Elfo" Or _
   UserList(UserIndex).raza = "Vampiro" Or _
   UserList(UserIndex).raza = "Orco" Or _
   UserList(UserIndex).raza = "Ciclope" Or _
   UserList(UserIndex).raza = "Elfo Oscuro" Then
        CheckRazaUsaRopa = (ObjData(itemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(itemIndex).RazaEnana = 1)
End If


Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & itemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo fallo
'Usa un item del inventario
Dim obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As obj
     Dim C As Integer
        Dim va1 As Integer
        Dim va2 As Integer
        Dim va3 As Integer
        Dim Cachis As Byte
obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

If UserList(UserIndex).flags.Privilegios > 0 Then GoTo sipuede

If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||Solo los newbies pueden usar estos objetos." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
End If
'comprueba si es elfo
 If obj.razaelfa = 1 And UserList(UserIndex).raza <> "Elfo" And UserList(UserIndex).raza <> "Elfo Oscuro" Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los Elfos pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If
 'comprueba si es vampiro
 If obj.razavampiro = 1 And UserList(UserIndex).raza <> "Vampiro" Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los Vampiros pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If
'comprueba si es enano
 If obj.RazaEnana = 1 And UserList(UserIndex).raza <> "Enano" Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los Enanos pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If
'comprueba si es humano
 If obj.razahumana = 1 And UserList(UserIndex).raza <> "Humano" Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los Humanos pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If
 'comprueba si es orco
 If obj.razaorca = 1 And UserList(UserIndex).raza <> "Orco" Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los Orcos pueden usar este objeto." & "´" & FontTypeNames.FONTTYPE_info)
     Exit Sub
End If

sipuede:

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).flags.TargetObjInvIndex = ObjIndex
UserList(UserIndex).flags.TargetObjInvSlot = Slot

Select Case obj.OBJType

'pluto:6.8-------Puntos Clan------------------------------------
Case 72
If UserList(UserIndex).Stats.PClan >= 0 Then
Call SendData(ToIndex, UserIndex, 0, "||No tienes Puntos Clan en Negativo." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
If UserList(UserIndex).Stats.GLD < (UserList(UserIndex).Stats.ELV * 500) Then
Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente Oro." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

 'Sonido
SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW45"
UserList(UserIndex).Stats.PClan = UserList(UserIndex).Stats.PClan + 1
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - (UserList(UserIndex).Stats.ELV * 500)

Call QuitarUserInvItem(UserIndex, Slot, 1)
Call SendData(ToIndex, UserIndex, 0, "||Has Ganado un Punto de Clan!! " & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Has Gastado " & (UserList(UserIndex).Stats.ELV * 500) & " Monedas de Oro." & "´" & FontTypeNames.FONTTYPE_info)
Call SendUserStatsOro(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)
Exit Sub







'pluto:6.5-------elixir de vida------------------------------------
Case 63
If UserList(UserIndex).flags.Elixir >= 3 Then
Call SendData(ToIndex, UserIndex, 0, "||No te hace ningún efecto." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
 'Sonido
SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER
UserList(UserIndex).flags.Elixir = UserList(UserIndex).flags.Elixir + 1

Call QuitarUserInvItem(UserIndex, Slot, 1)
Call SendData(ToIndex, UserIndex, 0, "||Obtendrás una bonificación de " & UserList(UserIndex).flags.Elixir & " Puntos de vida al pasar al siguiente Nivel." & "´" & FontTypeNames.FONTTYPE_info)
Call UpdateUserInv(False, UserIndex, Slot)
Exit Sub

Case 62
 'Sonido
SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER
UserList(UserIndex).flags.Elixir = 10

Call QuitarUserInvItem(UserIndex, Slot, 1)
Call SendData(ToIndex, UserIndex, 0, "||Obtendrás el Máximo de Puntos de Vida cuando pases al siguiente Nivel." & "´" & FontTypeNames.FONTTYPE_info)
Call UpdateUserInv(False, UserIndex, Slot)
Exit Sub

Case 67 'bolsitas vida
 'Sonido
 Dim Bolsita As Long
If obj.GrhIndex = 23583 Then Bolsita = 25000
If obj.GrhIndex = 23584 Then Bolsita = 50000

SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_DINERO
Call AddtoVar(UserList(UserIndex).Stats.GLD, Bolsita, MAXORO)

Call QuitarUserInvItem(UserIndex, Slot, 1)
Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & Bolsita & " Monedas de Oro." & "´" & FontTypeNames.FONTTYPE_info)
Call UpdateUserInv(False, UserIndex, Slot)
Call SendUserStatsOro(UserIndex)
Exit Sub

Case 68 'poción protección pluto:6.5
UserList(UserIndex).Counters.Protec = 500
UserList(UserIndex).flags.Protec = 10
Call SendData(ToIndex, UserIndex, 0, "S1")
Call SendData(ToIndex, UserIndex, 0, "||Circulo de Protección Mágica" & "´" & FontTypeNames.FONTTYPE_info)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 102 & "," & 1)
Call QuitarUserInvItem(UserIndex, Slot, 1)
Call UpdateUserInv(False, UserIndex, Slot)
Exit Sub
'------------------------------------------------------------------------

'pluto:2.11------------------------------------
Case 50
  If UserList(UserIndex).flags.Angel = 0 And UserList(UserIndex).flags.Demonio = 0 And UserList(UserIndex).flags.Morph = 0 And MapInfo(UserList(UserIndex).Pos.Map).Pk = True Then
    
    If UserList(UserIndex).flags.Muerto = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "L3")
    Exit Sub
    End If
If UserList(UserIndex).flags.Invisible = 1 Then
UserList(UserIndex).flags.Invisible = 0
UserList(UserIndex).flags.Oculto = 0
UserList(UserIndex).Counters.Invisibilidad = 0
Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 17, UserList(UserIndex).Char.CharIndex & ",0")
Else
UserList(UserIndex).flags.Invisible = 1
Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 17, UserList(UserIndex).Char.CharIndex & ",1")
End If
End If
'-----------------------------------------------
    Case OBJTYPE_USEONCE
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If

        'Usa el item
        Call AddtoVar(UserList(UserIndex).Stats.MinHam, obj.MinHam, UserList(UserIndex).Stats.MaxHam)
        UserList(UserIndex).flags.Hambre = 0
        Call EnviarHambreYsed(UserIndex)
        'pluto:6.2------ Sube Energía Newbies con Comida
        If EsNewbie(UserIndex) Then
        Call AddtoVar(UserList(UserIndex).Stats.MinSta, obj.MinHam, UserList(UserIndex).Stats.MaxSta)
        End If
        '---------------
        'Sonido
        SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_COMIDA
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
  
  'libros pluto:2.17
 Case 12
          
   If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
     Dim a As Byte
    a = RandomNumber(1, 10)

'malditos--------------------------------------
If a = 3 Then
        UserList(UserIndex).Stats.MinHP = 1
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Libro Maldito!!" & "´" & FontTypeNames.FONTTYPE_info)
       
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        SendUserStatsVida (UserIndex)
        'pluto:2.22
Call senduserstatsbox(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)
        Exit Sub
End If
If a = 4 And Not Criminal(UserIndex) Then
Call WarpUserChar(UserIndex, 170, Nix.X + RandomNumber(1, 5), Nix.Y, True)
 Call QuitarUserInvItem(UserIndex, Slot, 1)
Call SendData(ToIndex, UserIndex, 0, "||¡¡Libro Maldito!!" & "´" & FontTypeNames.FONTTYPE_info)
'pluto:2.22
Call senduserstatsbox(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)
Exit Sub
End If
If a = 4 And Criminal(UserIndex) Then
Call WarpUserChar(UserIndex, Banderbill.Map, Banderbill.X, Banderbill.Y - RandomNumber(1, 5), True)
 Call QuitarUserInvItem(UserIndex, Slot, 1)
Call SendData(ToIndex, UserIndex, 0, "||¡¡Libro Maldito!!" & "´" & FontTypeNames.FONTTYPE_info)
'pluto:2.22
Call senduserstatsbox(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)
Exit Sub
End If
'--------------------------------

If obj.GrhIndex = 538 Then
    
    'pluto:6.0A-----------------------------------------------------
    Dim AdiHp As Byte
    Dim AdihpR As Integer
    'Usa el item libro vida
If UserList(UserIndex).Remort = 1 Then
     Select Case UserList(UserIndex).clase
    Case "Guerrero"
    AdihpR = 800
    Case "Cazador"
    AdihpR = 650
    Case "Arquero"
    AdihpR = 500
    Case "Ladron"
    AdihpR = 625
    Case "Pirata"
    AdihpR = 625
    Case "Paladin"
   AdihpR = 650
    Case "Mago"
    AdihpR = 475
    Case "Clerigo"
    AdihpR = 600
    Case "Asesino"
    AdihpR = 625
    Case "Bardo"
    AdihpR = 625
    Case "Druida"
    AdihpR = 525
    Case Else
    AdihpR = 600
    End Select

End If 'remort
    'para remorts tope vida libros
    'pluto:6.6
If UserList(UserIndex).Remort = 1 Then
    If UserList(UserIndex).Stats.MaxHP >= AdihpR Then
    UserList(UserIndex).Stats.MaxHP = AdihpR
    Call SendData(ToIndex, UserIndex, 0, "||No puedes usar más Libros de Vida, tienes el máximo para tu clase." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
Else 'no remort
    
    If UserList(UserIndex).Stats.LibrosUsados >= (((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdiHp) * 2) Then
    Call SendData(ToIndex, UserIndex, 0, "||No puedes usar más Libros de Vida" & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
End If 'remort
  'añadimos el punto respetando topes de vida
  If UserList(UserIndex).Remort = 1 Then
  Call AddtoVar(UserList(UserIndex).Stats.MaxHP, 1, AdihpR)
  Call SendData(ToIndex, UserIndex, 0, "||¡¡Ganas 1 punto de Vida!!" & "´" & FontTypeNames.FONTTYPE_info)
  Call SendData(ToIndex, UserIndex, 0, "||¡¡Puedes usar Libros mientrás no superes los " & AdihpR & " de vida." & "´" & FontTypeNames.FONTTYPE_info)
  Else
  Call AddtoVar(UserList(UserIndex).Stats.MaxHP, 1, STAT_MAXHP)
  Call SendData(ToIndex, UserIndex, 0, "||¡¡Ganas 1 punto de Vida!!" & "´" & FontTypeNames.FONTTYPE_info)
  Call SendData(ToIndex, UserIndex, 0, "||¡¡Sólo puedes usar " & UserList(UserIndex).Stats.LibrosUsados - (((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdiHp) * 2) & " Libros más !!" & "´" & FontTypeNames.FONTTYPE_info)
  End If
  UserList(UserIndex).Stats.LibrosUsados = UserList(UserIndex).Stats.LibrosUsados + 1
'---------------------------fin pluto:6.0A-----------------------------------------


   ' UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + 1
      

'Sonido
    SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_resu
    'Quitamos del inv el item
    Call QuitarUserInvItem(UserIndex, Slot, 1)
    SendUserStatsVida (UserIndex)
    'End If
End If '538
'-------------------------------
  
  
  If obj.GrhIndex = 539 Then
    'Usa el item
    UserList(UserIndex).Stats.Puntos = UserList(UserIndex).Stats.Puntos + 20
    Call SendData(ToIndex, UserIndex, 0, "||¡¡Ganas 20 DraGPuntos!!" & "´" & FontTypeNames.FONTTYPE_info)
'Sonido
    SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_resu
 'Quitamos del inv el item
    Call QuitarUserInvItem(UserIndex, Slot, 1)
End If '539

 If obj.GrhIndex = 18530 Then
    'Usa el item
    UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + 2
    Call SendData(ToIndex, UserIndex, 0, "||¡¡Ganas 2 Puntos de Habilidad para Asignar!!" & "´" & FontTypeNames.FONTTYPE_info)
'Sonido
    SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_resu
    'Quitamos del inv el item
    Call QuitarUserInvItem(UserIndex, Slot, 1)
    SendUserStatsVida (UserIndex)
End If '18530
   '-------------------------------
  'amuleto resucitar
  Case OBJTYPE_resu
        
        If UserList(UserIndex).flags.Muerto <> 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas vivo!! Solo podes usar este items cuando estas muerto." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
'pluto:6.0A
If UserList(UserIndex).flags.Navegando > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas Navegando!! Solo podes usar este items cuando este en tierra." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
'pluto:6.0A
If MapInfo(UserList(UserIndex).Pos.Map).Resucitar = 1 Then Exit Sub

        'Usa el item
        Call RevivirUsuario(UserIndex)
       Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido resucitado!!" & "´" & FontTypeNames.FONTTYPE_info)

        'Sonido
        SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_resu
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
 'regalo
 Case OBJTYPE_regalo

        Dim rega As Integer
        Dim rega2 As Integer
        Dim rega3 As Integer
        Dim Rr As Byte
        Dim Reox As Byte

        
        
        
        rega2 = RandomNumber(1, 400)
        If rega2 + UserList(UserIndex).Stats.UserSkills(suerte) < 380 Then Rr = 1
        If rega2 + UserList(UserIndex).Stats.UserSkills(suerte) > 379 Then Rr = 2
        If rega2 + UserList(UserIndex).Stats.UserSkills(suerte) > 489 Then Rr = 3
     'pluto:6.5
     If Rr = 0 Then Rr = 1
       Select Case Rr
       Case 1
         rega3 = RandomNumber(1, Reo1)
        rega = ObjRegalo1(rega3)
      Case 2
         rega3 = RandomNumber(1, Reo2)
        rega = ObjRegalo2(rega3)
       Case 3
        rega3 = RandomNumber(1, Reo3)
        rega = ObjRegalo3(rega3)
       End Select
      

        'pluto:6.5
        If ObjData(rega).Pregalo = 0 Then rega = 158

        MiObj.ObjIndex = rega
        If ObjData(rega).Cregalos = 0 Then ObjData(rega).Cregalos = 1
        MiObj.Amount = ObjData(rega).Cregalos
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
        'pluto:2.14
        If UserList(UserIndex).flags.Privilegios > 0 Then
        Call LogGM(UserList(UserIndex).Name, "Regalo/Cofre: " & ObjData(rega).Name)
        End If
        
      'pluto:2.4 sonidos regalos y cofres
        If ObjIndex = 866 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás abierto un regalo!!" & "´" & FontTypeNames.FONTTYPE_info)
        'Sonido
        SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 118
        Else
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás abierto un Cofre!!" & "´" & FontTypeNames.FONTTYPE_info)
        'Sonido
        SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 120
       End If
       
'baston sube mana
Case OBJTYPE_WEAPON
   'pluto:2.22
        Dim Manita As Integer
  
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "L3")
        Exit Sub
    End If
If obj.objetoespecial > 0 Then

    Select Case obj.objetoespecial
    Case 51
    If Not IntervaloPermiteTomar(UserIndex) Then Exit Sub
    If UCase$(UserList(UserIndex).clase) = "CLERIGO" Or UCase$(UserList(UserIndex).clase) = "MAGO" Then
      Manita = 50 'Int(Porcentaje(UserList(UserIndex).Stats.MaxHP, 10))
  
        Call AddtoVar(UserList(UserIndex).Stats.MinHP, Manita, UserList(UserIndex).Stats.MaxHP)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 47)
    'pluto:2.14
    SendUserStatsVida (UserIndex)
    
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Sólo los Clérigos o Magos pueden usar este objeto. " & "´" & FontTypeNames.FONTTYPE_info)
    End If
    Case 52
        'If UCase$(UserList(UserIndex).clase) = "CLERIGO" Or UCase$(UserList(UserIndex).clase) = "MAGO" Then
If Not UserList(UserIndex).Invent.WeaponEqpObjIndex = 840 Then
Call SendData(ToIndex, UserIndex, 0, "||No tienes equipado el objeto!!. " & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
        Call AddtoVar(UserList(UserIndex).Stats.MinHam, 50, UserList(UserIndex).Stats.MaxHam)
        Call AddtoVar(UserList(UserIndex).Stats.MinAGU, 50, UserList(UserIndex).Stats.MaxAGU)
        UserList(UserIndex).flags.Sed = 0
        UserList(UserIndex).flags.Hambre = 0
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 47)
        Call EnviarHambreYsed(UserIndex)
    'Else
    'Call SendData(ToIndex, UserIndex, 0, "||Sólo los Clérigos pueden usar este objeto. " & "´" & FontTypeNames.FONTTYPE_info)
'End If

    Case 50
    'nati:solo si esta equipado
        If UCase$(UserList(UserIndex).clase) = "MAGO" Then
If Not UserList(UserIndex).Invent.WeaponEqpObjIndex = 842 Then
Call SendData(ToIndex, UserIndex, 0, "||No tienes equipado el objeto!!. " & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
        Manita = Int(Porcentaje(UserList(UserIndex).Stats.MaxMAN, 10))
        
        Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Manita, UserList(UserIndex).Stats.MaxMAN)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 47)
    'pluto:2.14
    SendUserStatsMana (UserIndex)
    Else
    Call SendData(ToIndex, UserIndex, 0, "||Sólo los Magos pueden usar este objeto. " & "´" & FontTypeNames.FONTTYPE_info)
End If
    Case 55
             'pluto:2.17
             
        If UserList(UserIndex).raza = "Elfo Oscuro" Then
        If Len(UserList(UserIndex).Padre) = 0 Then
            Exit Sub
        End If
        End If
        
          If UserList(UserIndex).flags.DuracionEfecto = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "S1")
                End If
        
        UserList(UserIndex).flags.TomoPocion = True
        UserList(UserIndex).flags.TipoPocion = 1
        UserList(UserIndex).flags.DuracionEfecto = 1000
     
'Usa el item
        Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), 5, UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 13)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 47)
    End Select
End If
    
    If ObjData(ObjIndex).proyectil = 1 Then
        Call SendData2(ToIndex, UserIndex, 0, 31, Proyectiles)
    Else
        If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(UserIndex).flags.TargetObj)
        '¿El target-objeto es leña?
        If TargObj.OBJType = OBJTYPE_LEÑA Then
            If UserList(UserIndex).Invent.Object(Slot).ObjIndex = DAGA Then
                Call TratarDeHacerFogata(UserList(UserIndex).flags.TargetObjMap _
                     , UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY, UserIndex)
            End If
        End If
    End If
'amuleto quitarparalisis
  Case OBJTYPE_para
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
If UserList(UserIndex).flags.Paralizado = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡No estás Paralizado!! " & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        Else
        'Usa el item
       UserList(UserIndex).flags.Paralizado = 1
        UserList(UserIndex).flags.Paralizado = 0
        Call SendData2(ToIndex, UserIndex, 0, 68)
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado la paralisis." & "´" & FontTypeNames.FONTTYPE_info)
               
        'Sonido
        SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_para
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
End If



'amuleto sanacion
  Case OBJTYPE_sana
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
'PLUTO:6.0A
If UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP Then Exit Sub
'--------

        'Usa el item
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
       Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sanado completamente!!" & "´" & FontTypeNames.FONTTYPE_info)

        'Sonido
        SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_sana
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)

'amuleto teleport: 'pluto:2.14 +c a los telep
 Case OBJTYPE_tele
        'pluto:6.0a
        If UserList(UserIndex).Counters.Pena > 0 Or UserList(UserIndex).Pos.Map = 191 Then Exit Sub
        'pluto:2.15
        If UserList(UserIndex).flags.Paralizado > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes estando paralizado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        
        C = RandomNumber(1, 5)
        
        If C = 1 Then
        va1 = Nix.Map
        va2 = Nix.X + C
        va3 = Nix.Y
        End If
        
        If C = 2 Then
        va1 = Banderbill.Map
        va2 = Banderbill.X
        va3 = Banderbill.Y - C
        End If
        
        If C = 3 Then
        va1 = Ullathorpe.Map
        va2 = Ullathorpe.X + C
        va3 = Ullathorpe.Y
        End If
        If C = 4 Then
        va1 = Lindos.Map
        va2 = Lindos.X
        va3 = Lindos.Y
        End If
        
        If C = 5 Then
        va1 = 170
        va2 = 34
        va3 = 34 + C
        End If
        
        'Usa el item
         
         Call WarpUserChar(UserIndex, va1, va2, va3, True)
  'PLUTO:6.0a
   Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 100 & "," & 1)
   'Sonido
    SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_tele
  
        
     
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
            
    Case OBJTYPE_GUITA
    
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        
        'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(Slot).Amount
        Call AddtoVar(UserList(UserIndex).Stats.GLD, UserList(UserIndex).Invent.Object(Slot).Amount, MAXORO)

        UserList(UserIndex).Stats.Peso = UserList(UserIndex).Stats.Peso - (UserList(UserIndex).Invent.Object(Slot).Amount * ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Peso)
        'pluto:2.4.5
        If UserList(UserIndex).Stats.Peso < 0.001 Then UserList(UserIndex).Stats.Peso = 0

        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
         Call SendUserStatsPeso(UserIndex)
    
    Case OBJTYPE_POCIONES
    
        'pluto:2.23
      If Not IntervaloPermiteTomar(UserIndex) Then Exit Sub
        '---------------------
        'If UserList(UserIndex).flags.PuedeAtacar = 0 Then
          '  Call SendData(ToIndex, UserIndex, 0, "||¡¡Debes esperar unos momentos para tomar otra pocion!!" & FONTTYPENAMES.FONTTYPE_INFO)
            'Exit Sub
        'End If
        'pluto:2.10
        'If UserList(UserIndex).flags.PuedeTomar = 0 Then
            'Call SendData(ToIndex, UserIndex, 0, "||¡¡Debes esperar unos momentos para tomar otra pocion!!" & FONTTYPENAMES.FONTTYPE_INFO)
            'Exit Sub
        'End If

        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        
        UserList(UserIndex).flags.TomoPocion = True
        UserList(UserIndex).flags.TipoPocion = obj.TipoPocion
        'pluto:2.10
        UserList(UserIndex).flags.PuedeTomar = 0
        
        Select Case UserList(UserIndex).flags.TipoPocion
        
            Case 1 'Modif la agilidad
                 'pluto:7.0
                If UserList(UserIndex).raza = "Elfo Oscuro" Then
                If Len(UserList(UserIndex).Padre) = 0 Then
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                Call UpdateUserInv(False, UserIndex, Slot)
                Exit Sub
                End If
                End If
                
                If UserList(UserIndex).flags.DuracionEfecto = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "S1")
                End If
                UserList(UserIndex).flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), RandomNumber(obj.MinModificador, obj.MaxModificador), UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 13)
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                    
            Case 2 'Modif la fuerza
                'pluto:7.0
                If UserList(UserIndex).raza = "Enano" Then
                If Len(UserList(UserIndex).Padre) = 0 Then
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                Call UpdateUserInv(False, UserIndex, Slot)
                Exit Sub
                End If
                End If
                
                If UserList(UserIndex).flags.DuracionEfecto = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "S1")
                End If
                UserList(UserIndex).flags.DuracionEfecto = obj.DuracionEfecto
    'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Fuerza), RandomNumber(obj.MinModificador, obj.MaxModificador), UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 13)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
               
            Case 3 'Pocion roja, restaura HP
                'pluto:6.0A mas potencia pociones en los sin mana
                If UserList(UserIndex).Stats.MaxMAN = 0 Then C = C + 10
                'pluto:7.0 pociones en humanos, nati: cambio de 10 a 5.
                'nati(18.06.11): veo algo ilógico que a una clase sin maná una poción roja pueda recuperarle lo de arriba + 5 extra.
                If UserList(UserIndex).raza = "Humano" And Not UserList(UserIndex).Stats.MaxMAN = 0 Then C = C + 5
                   
                AddtoVar UserList(UserIndex).Stats.MinHP, obj.MaxModificador + C, UserList(UserIndex).Stats.MaxHP

                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
            
            Case 4 'Pocion azul, restaura MANA
                
                'Usa el item
                If ObjData(ObjIndex).MaxModificador < 270 Then
                    'pluto:7.0 pociones en humanos
                    If UserList(UserIndex).raza = "Humano" Then
                    Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(UserList(UserIndex).Stats.MaxMAN, 7), UserList(UserIndex).Stats.MaxMAN)
                    Else
                    Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(UserList(UserIndex).Stats.MaxMAN, 5), UserList(UserIndex).Stats.MaxMAN)
                    End If
                'pluto: Pociones mejoradas
                Else
                    'pluto:7.0 pociones en humanos
                    If UserList(UserIndex).raza = "Humano" Then
                    Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(UserList(UserIndex).Stats.MaxMAN, 22), UserList(UserIndex).Stats.MaxMAN)
                    Else
                    Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(UserList(UserIndex).Stats.MaxMAN, 20), UserList(UserIndex).Stats.MaxMAN)
                    End If
                End If
                
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                
            Case 5 ' Pocion violeta
                If UserList(UserIndex).flags.Envenenado > 0 Then
                    UserList(UserIndex).flags.Envenenado = 0
                    Call SendData(ToIndex, UserIndex, 0, "||Te has curado del envenenamiento." & "´" & FontTypeNames.FONTTYPE_info)
                End If
                
                'Añadimos esto para pocion paralisis
                
                'If UserList(UserIndex).Flags.Paralizado = 1 Then
                    'UserList(UserIndex).Flags.Paralizado = 0
                    'Call SendData(ToIndex, UserIndex, 0, "PARADOK")
                    'Call SendData(ToIndex, UserIndex, 0, "||Te has quitado la paralisis." & FONTTYPENAMES.FONTTYPE_INFO)
                'End If
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
            Case 6 ' Ron para Pirata
            If (UCase$(UserList(UserIndex).clase) = "PIRATA") Then
    'Fuerza
                If UserList(UserIndex).flags.DuracionEfecto = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "S1")
                End If
                UserList(UserIndex).flags.DuracionEfecto = 6000
    'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Fuerza), RandomNumber(1, 5), UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 13)
    
    'Agilidad
                If UserList(UserIndex).flags.DuracionEfecto = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "S1")
                End If
                UserList(UserIndex).flags.DuracionEfecto = 6000
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), RandomNumber(1, 10), UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 13)
    'Aumenta la energia
    
UserList(UserIndex).Counters.Ron = 500
UserList(UserIndex).flags.Ron = 10
Call SendData(ToIndex, UserIndex, 0, "S1")
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 3 & "," & 1)
Call QuitarUserInvItem(UserIndex, Slot, 1)
Call UpdateUserInv(False, UserIndex, Slot)
Exit Sub
    
            End If
Call QuitarUserInvItem(UserIndex, Slot, 1)
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
       End Select
     Case OBJTYPE_BEBIDA
    
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        AddtoVar UserList(UserIndex).Stats.MinAGU, obj.MinSed, UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call EnviarHambreYsed(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
        
    
    Case OBJTYPE_LLAVES
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(UserIndex).flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = OBJTYPE_PUERTAS Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.Clave = obj.Clave Then
         
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerrada
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                        Call SendData(ToIndex, UserIndex, 0, "||Has abierto la puerta." & "´" & FontTypeNames.FONTTYPE_info)
                        Exit Sub
                     Else
                        Call SendData(ToIndex, UserIndex, 0, "||La llave no sirve." & "´" & FontTypeNames.FONTTYPE_info)
                        Exit Sub
                     End If
                  Else
                     If TargObj.Clave = obj.Clave Then
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerradaLlave
                        Call SendData(ToIndex, UserIndex, 0, "||Has cerrado con llave la puerta." & "´" & FontTypeNames.FONTTYPE_info)
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                        Exit Sub
                     Else
                        Call SendData(ToIndex, UserIndex, 0, "||La llave no sirve." & "´" & FontTypeNames.FONTTYPE_info)
                        Exit Sub
                     End If
                  End If
            Else
                  Call SendData(ToIndex, UserIndex, 0, "||No esta cerrada." & "´" & FontTypeNames.FONTTYPE_info)
                  Exit Sub
            End If
            
        End If
    
        Case OBJTYPE_BOTELLAVACIA
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "L3")
                Exit Sub
            End If
            If Not HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then
                Call SendData(ToIndex, UserIndex, 0, "||No hay agua allí." & "´" & FontTypeNames.FONTTYPE_info)
                Exit Sub
            End If
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexAbierta
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
            '    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            Call SendData(ToIndex, UserIndex, 0, "||Inventario Lleno." & "´" & FontTypeNames.FONTTYPE_info)

            End If
            
            
        Case OBJTYPE_BOTELLALLENA
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "L3")
                Exit Sub
            End If
            
            AddtoVar UserList(UserIndex).Stats.MinAGU, obj.MinSed, UserList(UserIndex).Stats.MaxAGU
            UserList(UserIndex).flags.Sed = 0
            Call EnviarHambreYsed(UserIndex)
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexCerrada
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            'If Not MeterItemEnInventario(UserIndex, MiObj) Then
             '   Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            'End If
            'pluto:2.17
            If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex > 0 Then
            If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex).OBJType = 15 Then
            Call EraseObj(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 1, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            Call SubirSkill(UserIndex, Supervivencia)
            End If
            End If
            
        Case OBJTYPE_HERRAMIENTAS
            
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "L3")
                Exit Sub
            End If
            If Not UserList(UserIndex).Stats.MinSta > 0 Then
               Call SendData(ToIndex, UserIndex, 0, "L7")
               Exit Sub
            End If
            
            If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Antes de usar la herramienta deberias equipartela." & "´" & FontTypeNames.FONTTYPE_info)
                Exit Sub
            End If
            
            Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlProleta, MAXREP)
            
            Select Case ObjIndex
                Case OBJTYPE_CAÑA
                    Call SendData2(ToIndex, UserIndex, 0, 31, Pesca)
                Case 543
                    Call SendData2(ToIndex, UserIndex, 0, 31, Pesca)
                Case HACHA_LEÑADOR
                    Call SendData2(ToIndex, UserIndex, 0, 31, Talar)
                Case PIQUETE_MINERO
                    Call SendData2(ToIndex, UserIndex, 0, 31, Mineria)
                Case MARTILLO_HERRERO
                    If (UCase$(UserList(UserIndex).clase) <> "HERRERO") Then
                        Call SendData(ToIndex, UserIndex, 0, "||Sólo los Herreros pueden usar estos objetos." & "´" & FontTypeNames.FONTTYPE_info)
                        Exit Sub
                    End If
                    Call SendData2(ToIndex, UserIndex, 0, 31, Herreria)
                Case SERRUCHO_CARPINTERO
                    If (UCase$(UserList(UserIndex).clase) <> "CARPINTERO") Then
                        Call SendData(ToIndex, UserIndex, 0, "||Sólo los Carpinteros pueden usar estos objetos." & "´" & FontTypeNames.FONTTYPE_info)
                        Exit Sub
                    End If
                    Call EnivarObjConstruibles(UserIndex)
                    Call SendData2(ToIndex, UserIndex, 0, 13)
                '[MerLiNz:6]
                Case SERRUCHOMAGICO_ermitano
                    If (UCase$(UserList(UserIndex).clase) <> "ERMITAÑO") Then
                        Call SendData(ToIndex, UserIndex, 0, "||Solo los ermitaños pueden usar estos objetos. " & "´" & FontTypeNames.FONTTYPE_info)
                    Else
                        Call EnviarObjMagicosConstruibles(UserIndex)
                        Call SendData2(ToIndex, UserIndex, 0, 13)
                    End If
                '[\END]
            End Select
        
        Case OBJTYPE_PERGAMINOS
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "L3")
                Exit Sub
            End If
            'pluto:6.0A
            If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) = False Then
             Call SendData(ToIndex, UserIndex, 0, "||El " & UserList(UserIndex).clase & " no puede usar este hechizo." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
            End If
            
            If UserList(UserIndex).flags.Hambre = 0 And _
               UserList(UserIndex).flags.Sed = 0 Then
                Call AgregarHechizo(UserIndex, Slot)
                
            Else
               Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado hambriento y sediento." & "´" & FontTypeNames.FONTTYPE_info)
            End If
       
       Case OBJTYPE_MINERALES
           If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "L3")
                Exit Sub
           End If
           Call SendData2(ToIndex, UserIndex, 0, 31, FundirMetal)
       
       Case OBJTYPE_INSTRUMENTOS
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "L3")
                Exit Sub
            End If
          'pluto:2.12
          If UserList(UserIndex).flags.Privilegios > 0 Then
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & obj.Snd1)
        End If
            
    If UCase$(UserList(UserIndex).clase) = "BARDO" Then
     If UserList(UserIndex).flags.DuracionEfecto = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "S1")
                End If
    UserList(UserIndex).flags.DuracionEfecto = 2000
    'Usa el item
    UserList(UserIndex).flags.TomoPocion = True
    Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), RandomNumber(1, 5), MAXATRIBUTOS)
    Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Fuerza), RandomNumber(1, 5), MAXATRIBUTOS)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & obj.Snd1)
       'pluto:2.12
Else
Call SendData(ToIndex, UserIndex, 0, "||Sólo para Bardos." & "´" & FontTypeNames.FONTTYPE_info)
End If
   
       
       Case OBJTYPE_BARCOS
           UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
           UserList(UserIndex).Invent.BarcoSlot = Slot

'pluto:2.3
Dim Pos As WorldPos
'pluto:2.4 añado el or para poder quitarlo en tierra
If HayAguaCerca(UserList(UserIndex).Pos) Or UserList(UserIndex).flags.Navegando = 1 Then
  Call DoNavega(UserIndex, obj)

 Else
Call SendData(ToIndex, UserIndex, 0, "||No puedes usar el barco en tierra." & "´" & FontTypeNames.FONTTYPE_info)
End If

'pluto:2.3
Case OBJTYPE_Montura

Call UsaMontura(UserIndex, obj)
If UserList(UserIndex).flags.Montura = 1 Then
UserList(UserIndex).flags.ClaseMontura = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo
Else
UserList(UserIndex).flags.ClaseMontura = 0
End If

End Select

'Actualiza
Call senduserstatsbox(UserIndex)
'Call UpdateUserInv(True, userindex, 0)
'pluto:2.4
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub
fallo:
Call LogError("USEINVITEM " & obj.Name & "->" & UserList(UserIndex).Name & " D: " & Err.Description)


End Sub

'[MerLiNz:6]
Sub EnviarObjMagicosConstruibles(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim i As Integer, cad$
Dim n As Byte
n = 0
For i = 1 To UBound(Objermitano)
    If (ObjData(Objermitano(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(Carpinteria) / ModCarpinteria(UserList(UserIndex).clase)) _
    And (ObjData(Objermitano(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(UserIndex).clase)) Then
        'cad$ = cad$ & ObjData(Objermitano(i)).Name & " (" & ObjData(Objermitano(i)).Madera & ":M) " & "(" & ObjData(Objermitano(i)).LingO & ":LO)" & "(" & ObjData(Objermitano(i)).LingP & ":LP)" & "(" & ObjData(Objermitano(i)).Gemas & ":G)" & "(" & ObjData(Objermitano(i)).Diamantes & ":D)" & "," & Objermitano(i) & ","
      n = n + 1
    cad$ = cad$ & Objermitano(i) & ","
    End If
Next i

Call SendData2(ToIndex, UserIndex, 0, 40, n + 1 & "," & cad$)
'[\END]
Exit Sub
fallo:
Call LogError("ENVIAROBJMAGICOSCONTRUIBLES " & Err.number & " D: " & Err.Description)


End Sub


Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim i As Integer, cad$
Dim n As Byte
For i = 1 To UBound(ArmasHerrero)
    If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) \ ModHerreriA(UserList(UserIndex).clase) Then
        'añado type=32 para municiones
        If ObjData(ArmasHerrero(i)).OBJType = OBJTYPE_WEAPON Or ObjData(ArmasHerrero(i)).OBJType = 32 Or ObjData(ArmasHerrero(i)).OBJType = 18 Then
            'cad$ = cad$ & ObjData(ArmasHerrero(i)).name & " (" & ObjData(ArmasHerrero(i)).MinHIT & "/" & ObjData(ArmasHerrero(i)).MaxHIT & ")" & "," & ArmasHerrero(i) & ","
            n = n + 1
            cad$ = cad$ & ArmasHerrero(i) & ","

        'Else
        'cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & "," & ArmasHerrero(i) & ","
        End If
    End If
Next i

Call SendData2(ToIndex, UserIndex, 0, 37, n + 1 & "," & cad$)

Exit Sub
fallo:
Call LogError("ENVIARARMASCONSTRUIBLES " & Err.number & " D: " & Err.Description)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim i As Integer, cad$
Dim n As Byte
n = 0
For i = 1 To UBound(ObjCarpintero)
    If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(Carpinteria) / ModCarpinteria(UserList(UserIndex).clase) Then
        n = n + 1
        cad$ = cad$ & ObjCarpintero(i) & ","
        End If
Next i

Call SendData2(ToIndex, UserIndex, 0, 39, n + 1 & "," & cad$)
Exit Sub
fallo:
Call LogError("ENVIAROBJCONSTRUIBLES " & Err.number & " D: " & Err.Description)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim i As Integer, cad$
Dim n As Byte
n = 0
For i = 1 To UBound(ArmadurasHerrero)
    If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(UserIndex).clase) Then
            n = n + 1
            cad$ = cad$ & ArmadurasHerrero(i) & ","
    End If
Next i

Call SendData2(ToIndex, UserIndex, 0, 38, n + 1 & "," & cad$)
Exit Sub
fallo:
Call LogError("ENVIARARMADURASCONSTRUIBLES " & Err.number & " D: " & Err.Description)

End Sub

                   

Sub TirarTodo(ByVal UserIndex As Integer)
On Error GoTo fallo
'PLUTO:6.7 AÑADO MAPA TORNEO TODOSVSTODOS Y SALAS CLAN
If UserList(UserIndex).Pos.Map = 191 Or UserList(UserIndex).Pos.Map = 293 Or UserList(UserIndex).Pos.Map = MapaTorneo2 Or UCase$(MapInfo(UserList(UserIndex).Pos.Map).Terreno) = "CLANATACA" Then Exit Sub

If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub
Call TirarTodosLosItems(UserIndex)
If UCase$(UserList(UserIndex).clase) <> "BANDIDO" Then Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)
Exit Sub
fallo:
Call LogError("TIRAR TODO " & Err.number & " D: " & Err.Description)

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean
On Error GoTo fallo
'pluto:2.18
If index = 1018 Or index = 1019 Then ItemSeCae = True: Exit Function

ItemSeCae = ObjData(index).Real <> 1 And _
            ObjData(index).nocaer <> 1 And _
            ObjData(index).Caos <> 1 And _
            ObjData(index).OBJType <> OBJTYPE_LLAVES And _
            ObjData(index).OBJType <> OBJTYPE_BARCOS

Exit Function
fallo:
Call LogError("ITEMSECAE " & Err.number & " D: " & Err.Description)

End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
On Error GoTo fallo
'PLUTO:2.4.2
If UserList(UserIndex).Pos.Map = 191 Or UserList(UserIndex).Pos.Map = 293 Or UserList(UserIndex).Pos.Map = MapaTorneo2 Then Exit Sub

'Call LogTarea("Sub TirarTodosLosItems")

Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As obj
Dim itemIndex As Integer

'Sacrificio amuleto
    If TieneObjetos(881, 1, UserIndex) = True Then
        Call QuitarObjetos(881, 1, UserIndex)
        'particula de que quemo sacri!!
        Exit Sub
    End If

For i = 1 To MAX_INVENTORY_SLOTS

  itemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
  If itemIndex > 0 Then
         If ItemSeCae(itemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                Tilelibre UserList(UserIndex).Pos, NuevaPos
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
  If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
         End If
         
  End If
  
Next i
Exit Sub
fallo:
Call LogError("TIRAR TODOS LOS ITEMS " & Err.number & " D: " & Err.Description)

End Sub


Function ItemNewbie(ByVal itemIndex As Integer) As Boolean
On Error GoTo fallo
ItemNewbie = ObjData(itemIndex).Newbie = 1
Exit Function
fallo:
Call LogError("ITEMNEWBIE " & Err.number & " D: " & Err.Description)

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
On Error GoTo fallo
'PLUTO:2.4.2
If UserList(UserIndex).Pos.Map = 191 Or UserList(UserIndex).Pos.Map = 293 Or UserList(UserIndex).Pos.Map = MapaTorneo2 Then Exit Sub

Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As obj
Dim itemIndex As Integer
'pluto:2-3-04
If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
  itemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
  If itemIndex > 0 Then
         If ItemSeCae(itemIndex) And Not ItemNewbie(itemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                Tilelibre UserList(UserIndex).Pos, NuevaPos
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
         End If
         
  End If
Next i
Exit Sub
fallo:
Call LogError("TIRARTODOSLOSITEMSNEWBIES " & Err.number & " D: " & Err.Description)

End Sub





