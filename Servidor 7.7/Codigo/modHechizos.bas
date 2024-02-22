Attribute VB_Name = "modHechizos"
Option Explicit

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)
On Error GoTo fallo
If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
If UserList(UserIndex).flags.Invisible = 1 And Npclist(NpcIndex).flags.Magiainvisible = 0 Then Exit Sub
If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub

'pluto:2.20
'If Hechizos(Spell).Noesquivar = 1 Then GoTo noesq

'pluto:6.0A skill
'Dim oo As Byte
'oo = RandomNumber(1, 100)
'Call SubirSkill(UserIndex, EvitaMagia)
'If oo < CInt((UserList(UserIndex).Stats.UserSkills(EvitaMagia) / 10) + 2) Then
'Call SendData(ToIndex, UserIndex, 0, "|| Has Resistido una Magia !!" & FONTTYPENAMES.FONTTYPE_fight)
'Exit Sub
'End If
'--------------------
noesq:

'pluto:6.0A
If Npclist(NpcIndex).Raid > 0 Then
Dim oo As Byte
oo = RandomNumber(1, 100)
If oo > 95 Then Spell = 69
End If


Npclist(NpcIndex).CanAttack = 0
Dim daño As Integer
'pluto:6.0A----------------------------------------------
If Npclist(NpcIndex).Anima = 1 Then
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 94, Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Char.Heading)
End If
'--------------------------------------------------------
If Hechizos(Spell).SubeHP = 1 Then

    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    'If UserList(userindex).raza = "Enano" Then daño = daño - CInt(daño / 5)
    'If UserList(userindex).raza = "Humano" Then daño = daño - CInt(daño / 10)

    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

'pluto:2.4
Call AddtoVar(UserList(UserIndex).Stats.MinHP, Porcentaje(UserList(UserIndex).Stats.MaxHP, 15), UserList(UserIndex).Stats.MaxHP)
Call SendData(ToIndex, UserIndex, 0, "V1")
Call senduserstatsbox(UserIndex)
    
ElseIf Hechizos(Spell).SubeHP = 2 Then
    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
 'pluto:7.0 extra monturas subido arriba
If UserList(UserIndex).flags.Montura = 1 Then
'Dim kk As Integer
'Dim oo As Integer
Dim nivk As Integer
oo = UserList(UserIndex).flags.ClaseMontura
'kk = 0
'If oo = 1 Then kk = 2
'If oo = 5 Then kk = 3
nivk = UserList(UserIndex).Montura.Nivel(oo)
daño = daño - CInt(Porcentaje(daño, UserList(UserIndex).Montura.DefMagico(oo))) - 1

'daño = daño - CInt(Porcentaje(daño, nivk * PMascotas(oo).ReduceMagia)) - 1
If daño < 1 Then daño = 1
End If
'------------fin pluto:2.4-------------------
'pluto:2.18
daño = daño - CInt(Porcentaje(daño, UserList(UserIndex).UserDefensaMagiasRaza))

' If UserList(UserIndex).raza = "Elfo" Then daño = daño - CInt(Porcentaje(daño, 8))
    'If UserList(UserIndex).raza = "Humano" Then daño = daño - CInt(Porcentaje(daño, 5))
    'If UserList(UserIndex).raza = "Gnomo" Then daño = daño - CInt(Porcentaje(daño, 15))
   ' If UserList(UserIndex).raza = "Elfo Oscuro" Then daño = daño - CInt(Porcentaje(daño, 5))

'pluto:6.0A Skills
'daño = daño - CInt(Porcentaje(daño, (CInt(UserList(UserIndex).Stats.UserSkills(DefMagia) / 10))))
'Call SubirSkill(UserIndex, DefMagia)
'-------------------

'pluto:2.16
If UserList(UserIndex).flags.Protec > 0 Then daño = daño - CInt(Porcentaje(daño, UserList(UserIndex).flags.Protec))

'pluto:7.0
 If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
daño = daño - ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Defmagica
If daño < 1 Then daño = 1
End If

'pluto:2.4.1
Dim obj As ObjData
If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).SubTipo = 4 Then daño = daño - CInt(daño / 5)
End If




    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    If UserList(UserIndex).flags.Privilegios = 0 Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
    
    Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida." & "´" & FontTypeNames.FONTTYPE_FIGHT)
       
    'pluto:7.0 10% quedar 1 vida en ciclopes
 If UserList(UserIndex).Stats.MinHP < 1 And UserList(UserIndex).raza = "Ciclope" Then
     Dim bup As Byte
     bup = RandomNumber(1, 10)
     If bup = 8 Then UserList(UserIndex).Stats.MinHP = 1

 End If
    
     Call SendUserStatsVida(UserIndex)

    'Muere
    If UserList(UserIndex).Stats.MinHP < 1 Then
        UserList(UserIndex).Stats.MinHP = 0
        'pluto:7.0 añado aviso de muerte
        Call SendData(ToIndex, UserIndex, 0, "6")
        
        Call UserDie(UserIndex)
    End If
    
End If
'pluto:2.4

If Hechizos(Spell).SubeMana = 1 Then
 'Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Hechizos(Spell).WAV)
'Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(UserList(UserIndex).Stats.MaxMAN, 15), UserList(UserIndex).Stats.MaxMAN)
Call SendData(ToIndex, UserIndex, 0, "V4")
Call senduserstatsbox(UserIndex)
End If
'-----fin pluto:2.4------------------
If Hechizos(Spell).Paraliza = 1 Then
     If UserList(UserIndex).flags.Paralizado = 0 Then
          UserList(UserIndex).flags.Paralizado = 1
          'pluto:7.0
          If UserList(UserIndex).raza = "Enano" Then
          UserList(UserIndex).Counters.Paralisis = CInt(IntervaloParalisisPJ - 50)
          Else
          UserList(UserIndex).Counters.Paralisis = IntervaloParalisisPJ
          End If
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
          Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
          Call SendData2(ToIndex, UserIndex, 0, 68)
      Call SendData2(ToIndex, UserIndex, 0, 15, UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
Dim rt As Integer
rt = RandomNumber(1, 100)
'If UserList(UserIndex).clase = "DRUIDA" And rt > 80 Then UserList(UserIndex).Counters.Paralisis = 0
'pluto:7.0
If UserList(UserIndex).raza = "Gnomo" And rt < 15 Then UserList(UserIndex).Counters.Paralisis = 0


'pluto:2.4.1
If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).SubTipo = 3 And rt > 80 Then
UserList(UserIndex).Counters.Paralisis = 0
Call SendData(ToIndex, UserIndex, 0, "||Anillo impide parálisis" & "´" & FontTypeNames.FONTTYPE_VENENO)
End If
End If
End If
End If
'ceguera
If Hechizos(Spell).Ceguera = 1 Then
     'pluto:2.10
     'nati: agrego el "Not UserList(UserIndex).flags.Morph = 214" para que no le afecte la ceguera al berserker
     If UserList(UserIndex).flags.Ceguera = 0 And UCase(UserList(UserIndex).clase) <> "BARDO" And UserList(UserIndex).flags.Angel = 0 And UserList(UserIndex).flags.Demonio = 0 And Not UserList(UserIndex).flags.Morph = 214 Then
          UserList(UserIndex).flags.Ceguera = 1
          UserList(UserIndex).Counters.Ceguera = Intervaloceguera
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)

        Call SendData2(ToIndex, UserIndex, 0, 2)
End If
End If
'estupidez
If Hechizos(Spell).Estupidez = 1 Then
     'pluto:2.11
     'nati: agrego el "Not UserList(UserIndex).flags.Morph = 214" para que lo ne afecte la estupidez
    If UserList(UserIndex).flags.Estupidez = 0 And UCase(UserList(UserIndex).clase) <> "BARDO" And UserList(UserIndex).flags.Angel = 0 And UserList(UserIndex).flags.Demonio = 0 And UserList(UserIndex).flags.Montura = 0 And Not UserList(UserIndex).flags.Morph = 214 Then
          UserList(UserIndex).flags.Estupidez = 1
          UserList(UserIndex).Counters.Estupidez = Intervaloceguera
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
          Call SendData2(ToIndex, UserIndex, 0, 3)
    End If
End If
'veneno
If Hechizos(Spell).Envenena > 1 Then
'[Tite]Añado la condicion de que no sea bardo el pj  y que no este muerto
    If UserList(UserIndex).flags.Envenenado = 0 And UserList(UserIndex).flags.Muerto = 0 And UCase(UserList(UserIndex).clase) <> "BARDO" Then
    'If UserList(UserIndex).flags.Envenenado = 0 Then
        UserList(UserIndex).flags.Envenenado = Hechizos(Spell).Envenena
        UserList(UserIndex).Counters.veneno = IntervaloVeneno
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
    End If
End If
'pluto:2.4
'fuerza npc
If Hechizos(Spell).SubeFuerza > 0 Then
If Not UserList(UserIndex).raza = "Enano" Then
'pluto:2.15
If UserList(UserIndex).flags.DuracionEfecto = 0 Then
Call SendData(ToIndex, UserIndex, 0, "S1")
End If

daño = RandomNumber(Hechizos(Spell).MinFuerza, Hechizos(Spell).MaxFuerza)
UserList(UserIndex).flags.DuracionEfecto = 1200
Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Fuerza), daño, UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) + 13)
UserList(UserIndex).flags.TomoPocion = True
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
Call SendData(ToIndex, UserIndex, 0, "V2")
'b = True
End If
End If
'agilidad npc
If Hechizos(Spell).SubeAgilidad > 0 Then
If Not UserList(UserIndex).raza = "Elfo Oscuro" Then
'pluto:2.15
If UserList(UserIndex).flags.DuracionEfecto = 0 Then
Call SendData(ToIndex, UserIndex, 0, "S1")
End If

daño = RandomNumber(Hechizos(Spell).MinAgilidad, Hechizos(Spell).MaxAgilidad)
UserList(UserIndex).flags.DuracionEfecto = 1200
Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), daño, UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 13)
UserList(UserIndex).flags.TomoPocion = True
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
Call SendData(ToIndex, UserIndex, 0, "V3")
'b = True
End If
End If

Exit Sub
fallo:
Call LogError("npclanzaspellsobreuser " & Npclist(NpcIndex).Name & "->" & Spell & " " & Err.number & " D: " & Err.Description)

End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

On Error GoTo fallo
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
fallo:
Call LogError("tienehechizo " & Err.number & " D: " & Err.Description)

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
On Error GoTo fallo
Dim hindex As Integer
Dim j As Integer
Dim pero As Byte
hindex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex


If Not TieneHechizo(hindex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
 pero = 4
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No tenes espacio para mas hechizos." & "´" & FontTypeNames.FONTTYPE_info)
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hindex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
         'pluto:2.17
 pero = 5
    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Dim n As Long
  pero = 6
    n = Porcentaje(ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Valor, 20)
    Call AddtoVar(UserList(UserIndex).Stats.GLD, n, MAXORO)
    Call SendData(ToIndex, UserIndex, 0, "||El Rey del Imperio te proporciona " & n & " Monedas de Oro para ayudarte en los gastos ocasionados por la compra de ese Hechizo. " & "´" & FontTypeNames.FONTTYPE_info)
    Call SendUserStatsOro(UserIndex)
    End If
    '-------
 pero = 7
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
   
    
    End If
Else
   pero = 8
    Call SendData(ToIndex, UserIndex, 0, "||Ya tienes ese hechizo." & "´" & FontTypeNames.FONTTYPE_info)
End If
Exit Sub
fallo:
Call LogError("agregarhechizo: " & UserList(UserIndex).Name & " " & hindex & " Señal: " & pero & " D: " & Err.Description)

End Sub
Sub AgregarHechizoangel(ByVal UserIndex As Integer, ByVal hindex As Integer)
On Error GoTo fallo
Dim j As Integer

If Not TieneHechizo(hindex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No tenes espacio para mas hechizos." & "´" & FontTypeNames.FONTTYPE_info)
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hindex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
       End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||Ya tenes ese hechizo." & "´" & FontTypeNames.FONTTYPE_info)
End If
Exit Sub
fallo:
Call LogError("agregarhechizoangel " & Err.number & " D: " & Err.Description)

End Sub
            
Sub DecirPalabrasMagicas(ByVal s As String, ByVal UserIndex As Integer)
On Error GoTo fallo

    Dim ind As String
    ind = UserList(UserIndex).Char.CharIndex
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||7°" & s & "°" & ind)
 
Exit Sub
fallo:
Call LogError("decirpalabrasmagicas " & Err.number & " D: " & Err.Description)


End Sub
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
On Error GoTo fallo


'pluto
If HechizoIndex = 0 Then Exit Function

If UserList(UserIndex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(UserIndex).flags.TargetMap
    wp2.X = UserList(UserIndex).flags.TargetX
    wp2.Y = UserList(UserIndex).flags.TargetY
    
  'pluto:2.14
If UserList(UserIndex).Pos.Map <> wp2.Map Then
Call SendData(ToIndex, UserIndex, 0, "||No seas tramposo " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
Call LogCasino(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de lanzar un spell desde otro mapa -> " & UserList(UserIndex).Pos.Map & " / " & wp2.Map)
Exit Function
End If
    
    'pluto:6.0A
If UserList(UserIndex).flags.Hambre > 0 Or UserList(UserIndex).flags.Sed > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||Demasiado hambriento o sediento para poder atacar!!" & "´" & FontTypeNames.FONTTYPE_info)
    Exit Function
End If
    
    
If Distancia(UserList(UserIndex).Pos, wp2) > 19 Then
            'UserList(UserIndex).Flags.AdministrativeBan = 1
            Call SendData(ToIndex, UserIndex, 0, "||No seas tramposo " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
            'Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de lanzar un spell desde otro mapa.")
            'Call CloseSocket(UserIndex)
            Exit Function
End If
'-------fin pluto:2.14--------------

    
    
    If UserList(UserIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(UserIndex).Stats.UserSkills(Magia) >= Hechizos(HechizoIndex).MinSkill Then
            PuedeLanzar = (UserList(UserIndex).Stats.MinSta > 0)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No tienes suficientes puntos en la habilidad APRENDIZAJE DE ARTES MÁGICAS para lanzar este hechizo." & "´" & FontTypeNames.FONTTYPE_info)
            PuedeLanzar = False
        End If
    Else
            Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente mana." & "´" & FontTypeNames.FONTTYPE_info)
            PuedeLanzar = False
    End If
Else
   Call SendData(ToIndex, UserIndex, 0, "L3")
   PuedeLanzar = False
End If

Dim H As Integer

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

 ' pluto:6.0A Restricciones por nivel
 If Hechizos(H).MinNivel > UserList(UserIndex).Stats.ELV Then
 Call SendData(ToIndex, UserIndex, 0, "||Necesitas nivel " & Hechizos(H).MinNivel & " para poder lanzar este hechizo." & "´" & FontTypeNames.FONTTYPE_info)
PuedeLanzar = False
End If

' solo angeles
If (H = 37 Or H = 38) And UserList(UserIndex).flags.Angel = 0 Then
Call SendData(ToIndex, UserIndex, 0, "||No eres Angel" & "´" & FontTypeNames.FONTTYPE_info)
PuedeLanzar = False
End If
' solo demonios
If (H = 53 Or H = 52) And UserList(UserIndex).flags.Demonio = 0 Then
Call SendData(ToIndex, UserIndex, 0, "||No eres Demonio" & "´" & FontTypeNames.FONTTYPE_info)
PuedeLanzar = False
End If

Exit Function
fallo:
Call LogError("puedelanzar " & Err.number & " D: " & Err.Description)


End Function

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)
On Error GoTo fallo
'Call LogTarea("HechizoInvocacion")
If UserList(UserIndex).NroMacotas >= MAXMASCOTAS Then Exit Sub
'pluto:2.17
'If MapInfo(UserList(UserIndex).Pos.Map).Terreno = "CONQUISTA" Then
'Call SendData(ToIndex, UserIndex, 0, "||No puedes en este Mapa!!." & FONTTYPENAMES.FONTTYPE_TALK)
'Exit Sub
'End If

If UserList(UserIndex).Pos.Map = 34 Then
    Call SendData(ToIndex, UserIndex, 0, "||No puedes invocar Mascotas en este Mapa." & "´" & FontTypeNames.FONTTYPE_TALK)
Exit Sub
End If

'pluto:6.0A
If MapInfo(UserList(UserIndex).Pos.Map).Mascotas = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||No puedes invocar Mascotas en este Mapa." & "´" & FontTypeNames.FONTTYPE_TALK)
Exit Sub
End If

Dim H As Integer, j As Integer, ind As Integer, index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(UserIndex).flags.TargetMap
TargetPos.X = UserList(UserIndex).flags.TargetX
TargetPos.Y = UserList(UserIndex).flags.TargetY

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
For j = 1 To Hechizos(H).Cant
    
    If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
  'pluto:2.4
If ind = MAXNPCS Then
Call SendData(ToIndex, UserIndex, 0, "||No hay sitio para tu mascota." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
        
        If ind < MAXNPCS Then
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
            
            index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(index) = ind
            UserList(UserIndex).MascotasType(index) = Npclist(ind).numero
            
            Npclist(ind).MaestroUser = UserIndex
            'pluto:mas duracion mascotas cutres
            
            If UCase$(Hechizos(H).Nombre) = "INVOCAR HADA" Or UCase$(Hechizos(H).Nombre) = "INVOCAR GENIO" Then IntervaloInvocacion = 1200
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(UserIndex)
b = True

Exit Sub
fallo:
Call LogError("hechizoinvocacion " & Err.number & " D: " & Err.Description)

End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)
On Error GoTo fallo
Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case uInvocacion '
       Call HechizoInvocacion(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    'pluto:7.0 menos mana elfos
    If UserList(UserIndex).raza <> "Elfo" Then
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    Else
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Porcentaje(Hechizos(uh).ManaRequerido, 85)
    End If
    'pluto:6.9
     If UserList(UserIndex).flags.Privilegios > 0 Then UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN

   'pluto:6.5----------------
Dim obj As ObjData
If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).SubTipo = 8 Then
Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(Hechizos(uh).ManaRequerido, 20), UserList(UserIndex).Stats.MaxMAN)
End If
End If
'----------------------------

    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    Call SendUserStatsMana(UserIndex)
End If

Exit Sub
fallo:
Call LogError("handlehechizoterreno " & Err.number & " D: " & Err.Description)

End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
    On Error GoTo fallo
    Dim b As Boolean
    Select Case Hechizos(uh).Tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoUsuario(UserIndex, b)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropUsuario(UserIndex, b)
    End Select
    
    If b Then
        Call SubirSkill(UserIndex, Magia)
'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
        
'pluto:7.0 menos mana elfos
        If UserList(UserIndex).raza <> "Elfo" Then
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
        Else
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Porcentaje(Hechizos(uh).ManaRequerido, 85)
        End If
        
'pluto:6.9
        If UserList(UserIndex).flags.Privilegios > 0 Then UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
        
'pluto:6.5----------------
        Dim obj As ObjData
        If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).SubTipo = 8 Then
                Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(Hechizos(uh).ManaRequerido, 20), UserList(UserIndex).Stats.MaxMAN)
            End If
        End If
'----------------------------
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        Call SendUserStatsMana(UserIndex)
        Call senduserstatsbox(UserList(UserIndex).flags.TargetUser)
        UserList(UserIndex).flags.TargetUser = 0
    End If
    Exit Sub
fallo:
    Call LogError("handlehechizousuario " & Err.number & " D: " & Err.Description)
    
End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)
'pluto:6.5--------------
'quitar esto
'GoTo je
'If Npclist(UserList(UserIndex).flags.TargetNpc).Raid > 0 And UserList(UserIndex).flags.Privilegios = 0 Then
 '   If UserList(UserIndex).flags.party = False Then
'Call SendData(ToIndex, UserIndex, 0, "||Debes estar en Party (Grupo) con 4 jugadores más para poder atacar este Monster DraG" & "´" & FontTypeNames.FONTTYPE_party)
'Exit Sub
 '   Else
  '      If partylist(UserList(UserIndex).flags.partyNum).numMiembros < 4 Then
'Call SendData(ToIndex, UserIndex, 0, "||Debes estar en Party (Grupo) con 4 jugadores más para poder atacar este Monster DraG" & "´" & FontTypeNames.FONTTYPE_party)
'Exit Sub
 '       End If
 '   End If
  '          If UserList(UserIndex).Stats.ELV > Npclist(UserList(UserIndex).flags.TargetNpc).Raid Then
  '          Call SendData(ToIndex, UserIndex, 0, "||Los Dioses no te dejan atacar este MonsterDraG, tienes demasiado nivel." & "´" & FontTypeNames.FONTTYPE_party)
   '         End If

'End If
'--------------------
'je:
On Error GoTo fallo
Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNpc, uh, b, UserIndex)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNpc, UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).flags.TargetNpc = 0
    
    'pluto:7.0 menos mana elfos
    If UserList(UserIndex).raza <> "Elfo" Then
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    Else
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Porcentaje(Hechizos(uh).ManaRequerido, 85)
    End If
    
         'pluto:6.9
     If UserList(UserIndex).flags.Privilegios > 0 Then UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
 
       'pluto:6.5----------------
Dim obj As ObjData
If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).SubTipo = 8 Then
Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(Hechizos(uh).ManaRequerido, 20), UserList(UserIndex).Stats.MaxMAN)
End If
End If
'----------------------------
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    Call SendUserStatsMana(UserIndex)
End If
Exit Sub
fallo:
Call LogError("handlehechizonpc " & Err.number & " D: " & Err.Description)

End Sub
Sub LanzarHechizo(index As Integer, UserIndex As Integer)
On Error GoTo fallo
Dim uh As Integer
Dim exito As Boolean

uh = UserList(UserIndex).Stats.UserHechizos(index)

If PuedeLanzar(UserIndex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case uUsuarios
            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Este hechizo actua solo sobre usuarios." & "´" & FontTypeNames.FONTTYPE_info)
            End If
        Case uNPC
            If UserList(UserIndex).flags.TargetNpc > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Este hechizo solo afecta a los npcs." & "´" & FontTypeNames.FONTTYPE_info)
            End If
        Case uUsuariosYnpc
            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
            ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Target invalido." & "´" & FontTypeNames.FONTTYPE_info)
            End If
        Case uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)
    End Select
    
End If
                
Exit Sub
fallo:
Call LogError("lanzarhechizo " & Err.number & " D: " & Err.Description)

End Sub
Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)


On Error GoTo fallo
Dim H As Integer, TU As Integer, abody As Integer, al As Integer
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
TU = UserList(UserIndex).flags.TargetUser


'pluto:2.17
'If Hechizos(H).Invisibilidad = 1 And UserList(UserIndex).Pos.Map = 252 Then Exit Sub

'pluto:6.0A
'If Hechizos(H).Noesquivar = 1 Then GoTo noes
'Dim oo As Byte
'oo = RandomNumber(1, 100)
'Call SubirSkill(TU, EvitaMagia)
'If oo < CInt((UserList(TU).Stats.UserSkills(EvitaMagia) / 10) + 2) And UserIndex <> TU Then
'Call SendData(ToIndex, UserIndex, 0, "|| Se ha Resistido a la Magia !!" & FONTTYPENAMES.FONTTYPE_fight)
'Call SendData(ToIndex, TU, 0, "|| Has Resistido una Magia !!" & FONTTYPENAMES.FONTTYPE_fight)
'b = True
'Exit Sub
'End If
'--------------------
noes:



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
'[MerLiNz:X]
If Hechizos(H).Morph = 1 And UserList(TU).flags.Morph = 0 And UserList(TU).flags.Angel = 0 And UserList(TU).flags.Demonio = 0 Then
   If UserList(TU).flags.Navegando = 1 Or UserList(TU).flags.Muerto > 0 Then Exit Sub
'pluto:2.14
If UserList(TU).flags.ClaseMontura > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes usar este hechizo contra una mascota." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'pluto:6.9
If MapInfo(UserList(TU).Pos.Map).Pk = False And TU <> UserIndex Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes usar este hechizo sobre otros personajes en zona segura." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

'[\END]
If UCase$(UserList(UserIndex).clase) <> "DRUIDA" And UCase$(UserList(UserIndex).clase) <> "MAGO" Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes usar este hechizo." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
If UCase$(UserList(UserIndex).clase) = "MAGO" And UserList(UserIndex).Stats.ELV < 30 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes usar este hechizo hasta level 30." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
   UserList(TU).flags.Morph = UserList(TU).Char.Body
   UserList(TU).Counters.Morph = IntervaloMorphPJ
   Call InfoHechizo(UserIndex)
    '[gau]
    Call ChangeUserChar(ToMap, 0, UserList(TU).Pos.Map, TU, val(abody), val(0), UserList(TU).Char.Heading, UserList(TU).Char.WeaponAnim, UserList(TU).Char.ShieldAnim, UserList(TU).Char.CascoAnim, UserList(UserIndex).Char.Botas)
    Call SendData2(ToPCArea, UserIndex, UserList(TU).Pos.Map, 22, UserList(TU).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
   b = True
End If

 'pluto:2.7.0 impide invis en demonios, angeles..
If Hechizos(H).Invisibilidad = 1 And UserList(TU).flags.Morph = 0 And UserList(TU).flags.Angel = 0 And UserList(TU).flags.Demonio = 0 Then
'pluto:6.0A-----
If MapInfo(UserList(TU).Pos.Map).Pk = False Then GoTo nopi
'---------------
UserList(TU).flags.Invisible = 1
   Call SendData2(ToMap, 0, UserList(TU).Pos.Map, 17, UserList(TU).Char.CharIndex & ",1")
   Call InfoHechizo(UserIndex)
   
'gollum
Dim ry88 As Integer
ry88 = RandomNumber(1, 1000)
'pluto:2-3-04
'If ry88 = 251 Then Tesoromomia = 0 SE PASO A UN TIMER'NICO
'If ry88 = 243 Then tesorocaballero = 0 SE PASO A UN TIMER
'pluto:6.0 añade sala invo 165
If ry88 = 250 And MapInfo(UserList(TU).Pos.Map).Pk = True And Not (UserList(TU).Pos.Map > 164 And UserList(TU).Pos.Map < 170) Then
Call SpawnNpc(594, UserList(TU).Pos, True, False)
Call SendData(ToAll, 0, 0, "TW" & 106)
Call SendData(ToAll, 0, 0, "||¡¡¡ Gollum, la más terrible de las criaturas apareció junto a " & UserList(UserIndex).Name & " en el Mapa " & UserList(UserIndex).Pos.Map & " !!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
End If
'fin gollum
b = True
End If

nopi:
If Hechizos(H).Envenena > 0 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        
If UCase$(UserList(TU).clase) <> "BARDO" And UserList(TU).flags.Angel = 0 And UserList(TU).flags.Demonio = 0 Then
UserList(TU).flags.Envenenado = Hechizos(H).Envenena + CInt(UserList(UserIndex).Stats.ELV / 5)
        Call InfoHechizo(UserIndex)
        b = True
Else
Call SendData(ToIndex, TU, 0, "|| " & UserList(UserIndex).Name & " te ha intentado envenenar, pero eres INMUNE!!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendData(ToIndex, UserIndex, 0, "|| " & UserList(TU).Name & " es INMUNE!!" & "´" & FontTypeNames.FONTTYPE_FIGHT)

End If
End If
'pluto:2.15
If Hechizos(H).Protec > 0 Then
        
    If UserIndex = TU Then
    UserList(TU).flags.Protec = Hechizos(H).Protec
    Call InfoHechizo(UserIndex)
    UserList(TU).Counters.Protec = 100 * Hechizos(H).Protec
    Call SendData(ToIndex, UserIndex, 0, "S1")

    b = True
    Else
    Call SendData(ToIndex, UserIndex, 0, "|| No puedes lanzar este hechizo sobre otros usuarios." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
        
End If
'--------------

If Hechizos(H).CuraVeneno = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Maldicion = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Paraliza = 1 Then

     If UserList(TU).flags.Paralizado = 0 And UserList(TU).flags.Muerto = 0 Then
            If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
            
            If UserIndex <> TU Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TU)
            End If
            
            UserList(TU).flags.Paralizado = 1
            
             'pluto:7.0
          If UserList(TU).raza = "Enano" Then
          UserList(TU).Counters.Paralisis = CInt(IntervaloParalisisPJ - 50)
          Else
          UserList(TU).Counters.Paralisis = IntervaloParalisisPJ
          End If
            
           
            Dim rt As Integer
            rt = RandomNumber(1, 100)
            'If UCase$(UserList(TU).clase) = "DRUIDA" And rt > 80 Then UserList(TU).Counters.Paralisis = 0
'pluto:7.0
If UserList(TU).raza = "Gnomo" And rt < 15 Then UserList(TU).Counters.Paralisis = 0

'pluto:2.4.1
Dim obj As ObjData
If UserList(TU).Invent.AnilloEqpObjIndex > 0 Then
If ObjData(UserList(TU).Invent.AnilloEqpObjIndex).SubTipo = 3 And rt > 80 Then
UserList(TU).Counters.Paralisis = 0
Call SendData(ToIndex, TU, 0, "||Anillo impide parálisis" & "´" & FontTypeNames.FONTTYPE_VENENO)
End If
End If
            Call SendData2(ToIndex, TU, 0, 68)
            Call SendData2(ToIndex, TU, 0, 15, UserList(TU).Pos.X & "," & UserList(TU).Pos.Y)
            Call InfoHechizo(UserIndex)
  b = True
    End If
End If

If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(TU).flags.Paralizado = 1 Then
                UserList(TU).flags.Paralizado = 0
                Call SendData2(ToIndex, TU, 0, 68)
                Call InfoHechizo(UserIndex)
                b = True
    End If
End If

If Hechizos(H).Revivir = 1 Then
   'pluto:6.0A
   If MapInfo(UserList(TU).Pos.Map).Resucitar = 1 Then Exit Sub
    
    If UserList(TU).flags.Muerto = 1 And UserList(TU).Char.Body <> 87 Then
        If Not Criminal(TU) Then
                If TU <> UserIndex Then
                    Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, 500, MAXREP)
                    Call SendData(ToIndex, UserIndex, 0, "||¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!." & "´" & FontTypeNames.FONTTYPE_info)
                End If
        End If
If UCase$(Hechizos(H).Nombre) = "PODER DIVINO" Then Call RevivirUsuarioangel(TU) Else Call RevivirUsuario(TU)
Call InfoHechizo(UserIndex)
    b = True
Else
Call SendData(ToIndex, UserIndex, 0, "||¡No puedes resucitar, no está muerto o está en modo barco." & "´" & FontTypeNames.FONTTYPE_info)
End If

End If

If Hechizos(H).Ceguera = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
If UCase$(UserList(TU).clase) <> "BARDO" And UserList(TU).flags.Angel = 0 And UserList(TU).flags.Demonio = 0 And UserList(TU).flags.Montura = 0 Then
UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = Intervaloceguera
        Call SendData2(ToIndex, TU, 0, 2)
        Call InfoHechizo(UserIndex)
        b = True
Else
Call SendData(ToIndex, TU, 0, "|| " & UserList(UserIndex).Name & " te ha intentado cegar, pero eres INMUNE!!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendData(ToIndex, UserIndex, 0, "|| " & UserList(TU).Name & " es INMUNE!!" & "´" & FontTypeNames.FONTTYPE_FIGHT)

End If
End If

If Hechizos(H).Estupidez = 1 Then
    If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
    If UserIndex <> TU Then
        Call UsuarioAtacadoPorUsuario(UserIndex, TU)
    End If
             'pluto:2.11 añade montura
    If UCase$(UserList(TU).clase) <> "BARDO" And UserList(TU).flags.Angel = 0 And UserList(TU).flags.Demonio = 0 And UserList(TU).flags.Montura = 0 Then
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Estupidez = Intervaloceguera
        Call SendData2(ToIndex, TU, 0, 3)
        Call InfoHechizo(UserIndex)
        b = True
    Else
        Call SendData(ToIndex, TU, 0, "|| " & UserList(UserIndex).Name & " te ha intentado volver estúpido, pero eres INMUNE!!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(ToIndex, UserIndex, 0, "|| " & UserList(TU).Name & " es INMUNE!!" & "´" & FontTypeNames.FONTTYPE_FIGHT)

    End If
End If
Exit Sub
fallo:
Call LogError("hechizoestadousuario " & Err.number & " D: " & Err.Description)

End Sub
Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hindex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)

On Error GoTo fallo

If Hechizos(hindex).Invisibilidad = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

If Hechizos(hindex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
    Call SendData(ToIndex, UserIndex, 0, "L5")
    Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If

If Hechizos(hindex).CuraVeneno = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hindex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "L5")
        Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 1
   b = True
End If

If Hechizos(hindex).RemoverMaldicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hindex).Bendicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If
'paralisis en area
If Hechizos(hindex).Paralizaarea = 1 Then
  If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 1
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "||El npc es inmune a este hechizo." & "´" & FontTypeNames.FONTTYPE_info)
   End If
Dim X As Integer
Dim Y As Integer
Dim H As Integer
'Dim P As Integer
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
'P = MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex
   For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
            If MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex > 0 Then
            
            If Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).flags.AfectaParalisis = 0 Then
            'Call InfoHechizo(UserIndex)
            Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).flags.Paralizado = 1
            Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).Contadores.Paralisis = IntervaloParalizado
            Call SendData2(ToPCArea, UserIndex, Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).Pos.Map, 22, Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)

            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "||El npc es inmune a este hechizo." & "´" & FontTypeNames.FONTTYPE_FIGHT)
   End If
            End If
            End If
            Next X
    Next Y
End If

If Hechizos(hindex).Paraliza = 1 Then
   If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 1
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "||El npc es inmune a este hechizo." & "´" & FontTypeNames.FONTTYPE_info)
   End If
End If

If Hechizos(hindex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
    
Else
Call SendData(ToIndex, UserIndex, 0, "||El npc no esta paralizado." & "´" & FontTypeNames.FONTTYPE_info)
End If
End If


 
Exit Sub
fallo:
Call LogError("hechizoestadonpc " & Err.number & " D: " & Err.Description)


End Sub

Sub HechizoPropNPC(ByVal hindex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)
On Error GoTo errhandler
Dim daño As Integer
Dim Loco As Integer
Dim nPos As WorldPos
Dim Critico As Integer
Dim Criti As Byte
 Dim Topito As Long
 Dim LogroOro As Boolean

'pluto:2.17
If Npclist(NpcIndex).NPCtype = 78 Then

nPos.Map = Npclist(NpcIndex).Pos.Map
nPos.X = Npclist(NpcIndex).Pos.X
nPos.Y = Npclist(NpcIndex).Pos.Y
'pluto:6.0A-----------------
If Hechizos(hindex).SubeHP = 1 And nPos.Y > UserList(UserIndex).Pos.Y Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes restaurar la puerta desde este lado." & "´" & FontTypeNames.FONTTYPE_GUILD)
Exit Sub
End If
'----------------------------
Select Case Npclist(NpcIndex).Stats.MinHP
Case 10000 To 15000
Npclist(NpcIndex).Char.Body = 360

Case 5000 To 9999
Npclist(NpcIndex).Char.Body = 361

Case 1 To 4999
Npclist(NpcIndex).Char.Body = 362

End Select

Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, 0, 1, 1)

End If
'--------------------------------------------


If Hechizos(hindex).SubeHP > 1 Then
'pluto:2.15
'If Npclist(NpcIndex).NPCtype = 79 Then

'If (MapInfo(Npclist(NpcIndex).Pos.Map).Dueño = 1 And UserList(UserIndex).Faccion.FuerzasCaos = 0) Or (MapInfo(Npclist(NpcIndex).Pos.Map).Dueño = 2 And UserList(UserIndex).Faccion.ArmadaReal = 0) Then
'Call SendData(ToIndex, UserIndex, 0, "||Tu armada te prohibe atacar este NPC." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If

'pluto:2.17
'If Conquistas = False Then
'Call SendData(ToIndex, UserIndex, 0, "||No se puede conquistar ciudades en estos momentos." & FONTTYPENAMES.FONTTYPE_INFO)
'Exit Sub
'End If

'End If '79
 '--------------
      
    'pluto:2.17
   If Npclist(NpcIndex).NPCtype = 33 Or Npclist(NpcIndex).NPCtype = 61 Or Npclist(NpcIndex).NPCtype = 77 Or Npclist(NpcIndex).NPCtype = 78 Then
        If MapInfo(Npclist(NpcIndex).Pos.Map).Zona = "CASTILLO" Then
            Dim castiact As String
            If Npclist(NpcIndex).Pos.Map = mapa_castillo1 Then castiact = castillo1
            If Npclist(NpcIndex).Pos.Map = mapa_castillo2 Then castiact = castillo2
            If Npclist(NpcIndex).Pos.Map = mapa_castillo3 Then castiact = castillo3
            If Npclist(NpcIndex).Pos.Map = mapa_castillo4 Then castiact = castillo4
            'pluto:2.18
            If Npclist(NpcIndex).Pos.Map = 268 Then castiact = castillo1
            If Npclist(NpcIndex).Pos.Map = 269 Then castiact = castillo2
            If Npclist(NpcIndex).Pos.Map = 270 Then castiact = castillo3
            If Npclist(NpcIndex).Pos.Map = 271 Then castiact = castillo4
            '------------------------------
            If Npclist(NpcIndex).Pos.Map = 185 Then castiact = fortaleza

            If UserList(UserIndex).GuildInfo.GuildName = "" Then
                Call SendData(ToIndex, UserIndex, 0, "||No tienes clan!!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
             Exit Sub
            End If
            If UserList(UserIndex).GuildInfo.GuildName = castiact Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar tu castillo ¬¬" & "´" & FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
'pluto:2.4.1

If UserList(UserIndex).Pos.Map = 185 And (UserList(UserIndex).GuildInfo.GuildName <> castillo1 Or UserList(UserIndex).GuildInfo.GuildName <> castillo2 Or UserList(UserIndex).GuildInfo.GuildName <> castillo3 Or UserList(UserIndex).GuildInfo.GuildName <> castillo4) Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar Fortaleza sin tener Conquistado los 4 Castillos." & "´" & FontTypeNames.FONTTYPE_FIGHT)
Exit Sub
End If

'pluto.6.0A
If UserList(UserIndex).GuildInfo.GuildName <> "" Then
If UserList(UserIndex).GuildRef.Nivel < 2 And Npclist(NpcIndex).NPCtype = 61 And UserList(UserIndex).Pos.Map = 185 Then
Call SendData(ToIndex, UserIndex, 0, "||Tu Clan no tiene suficiente Nivel." & "´" & FontTypeNames.FONTTYPE_FIGHT)
Exit Sub
End If
End If
'-----------------
            Set UserList(UserIndex).GuildRef = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
            If Not UserList(UserIndex).GuildRef Is Nothing Then
                If UserList(UserIndex).GuildRef.IsAllie(castiact) Then
                   Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar castillos de clanes aliados :P" & "´" & FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If
        End If
    End If
End If



'Salud
If Hechizos(hindex).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(hindex).MinHP, Hechizos(hindex).MaxHP)
    
    'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
     'pluto:6.0----------------------------------------
If UserList(UserIndex).Remort = 0 Then
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
Else
    If UserList(UserIndex).clase = "Mago" Or UserList(UserIndex).clase = "Druida" Then
    'Dim Topito As Long
    Topito = UserList(UserIndex).Stats.ELV * 3.65
    If UserList(UserIndex).Stats.ELV > 45 Then Topito = 45 * 3.65
    daño = daño + Porcentaje(daño, Topito)
    Else
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    End If
End If
    '-------------------------------------------------
    
    'pluto:2.17
    Dim lleno As Byte
    'If Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then lleno = 1 Else lleno = 0
    If Npclist(NpcIndex).Stats.MaxHP < Npclist(NpcIndex).Stats.MinHP + daño Then lleno = 1 Else lleno = 0
    Call InfoHechizo(UserIndex)
    Call AddtoVar(Npclist(NpcIndex).Stats.MinHP, daño, Npclist(NpcIndex).Stats.MaxHP)
    Call SendData(ToIndex, UserIndex, 0, "||Has curado " & daño & " puntos de salud a la criatura." & "´" & FontTypeNames.FONTTYPE_FIGHT)
     b = True
     'pluto:2.15
     If (Npclist(NpcIndex).NPCtype = 78 Or Npclist(NpcIndex).NPCtype = 77 Or Npclist(NpcIndex).NPCtype = 33 Or Npclist(NpcIndex).NPCtype = 61) And lleno = 1 Then
     
     If Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then
     
     Select Case Npclist(NpcIndex).Pos.Map
     Case 268
     Call SendData(ToAll, 0, 0, "C5")
     AtaNorte = 0
     Case 269
     Call SendData(ToAll, 0, 0, "C6")
     AtaSur = 0
     Case 270
     Call SendData(ToAll, 0, 0, "C7")
     AtaEste = 0
     Case 271
     Call SendData(ToAll, 0, 0, "C8")
     AtaOeste = 0
    Case 166
     Call SendData(ToAll, 0, 0, "C5")
     AtaNorte = 0
     Case 167
     Call SendData(ToAll, 0, 0, "C6")
     AtaSur = 0
     Case 168
     Call SendData(ToAll, 0, 0, "C7")
     AtaEste = 0
     Case 169
     Call SendData(ToAll, 0, 0, "C8")
     AtaOeste = 0
     Case 185
     Call SendData(ToAll, 0, 0, "V9")
     AtaForta = 0
     End Select
     
     
     End If
     
     End If
     
ElseIf Hechizos(hindex).SubeHP = 2 Then
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "L5")
        Exit Sub
    End If
'pluto:6.6--------
If Npclist(NpcIndex).MaestroUser = UserIndex Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar tus mascotas." & "´" & FontTypeNames.FONTTYPE_FIGHT)
Exit Sub
End If
'pluto:6.7
If Npclist(NpcIndex).MaestroUser > 0 And MapInfo(Npclist(NpcIndex).Pos.Map).Pk = False Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar mascotas en zona segura." & "´" & FontTypeNames.FONTTYPE_FIGHT)
Exit Sub
End If


'-----------------
'pluto:2.6.0
    If (EsMascotaCiudadano(NpcIndex, UserIndex) Or Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS) And Not Criminal(UserIndex) Then
        If UserList(UserIndex).Faccion.ArmadaReal > 0 Then Exit Sub
            If UserList(UserIndex).flags.Seguro = True Then
                Call SendData(ToIndex, UserIndex, 0, "||Debes desactivar el seguro." & "´" & FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If
    End If
  
  'pluto:2.11
'If Npclist(NpcIndex).Stats.Alineacion = 0 And UserList(UserIndex).Faccion.ArmadaReal > 0 Then
'Call SendData(ToIndex, UserIndex, 0, "||Tu armada te prohibe atacar este tipo de criaturas." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If
    
    'pluto:6.5----------------------
    If UserList(UserIndex).flags.Privilegios > 0 Then
    Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name
    End If
    '------------------------------
    daño = RandomNumber(Hechizos(hindex).MinHP, Hechizos(hindex).MaxHP)
    If UCase$(Hechizos(hindex).Nombre) = "RAYO GM" Then
    'pluto:2.14
    Call LogGM(UserList(UserIndex).Name, "RAYO GM: " & Npclist(NpcIndex).Name)

     daño = 800
    'quitar esto
    Npclist(NpcIndex).Stats.MinHP = 0
    End If
    
    'pluto:6.0----------------------------------------
If UserList(UserIndex).Remort = 0 Then
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
Else
    If UserList(UserIndex).clase = "Mago" Or UserList(UserIndex).clase = "Druida" Then
   ' Dim Topito As Long
    Topito = UserList(UserIndex).Stats.ELV * 3.65
    If UserList(UserIndex).Stats.ELV > 45 Then Topito = 45 * 3.65
    daño = daño + Porcentaje(daño, Topito)
    Else
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    End If
    
End If
    '-------------------------------------------------
    
'pluto:7.0 añado logro plata y oro-------------------------
'LogroOro = False
If Npclist(NpcIndex).LogroTipo > 0 Then
    Select Case UserList(UserIndex).Stats.PremioNPC(Npclist(NpcIndex).LogroTipo)
    Case 25 To 249
    daño = daño + Porcentaje(daño, 5)
    Case Is > 249
    daño = daño + Porcentaje(daño, 15)
    Case Is > 449
    LogroOro = True
    'If UserList(UserIndex).Stats.PremioNPC(Npclist(NpcIndex).LogroTipo) > 249 Then daño = daño + Porcentaje(daño, 10)
    'If UserList(UserIndex).Stats.PremioNPC(Npclist(NpcIndex).LogroTipo) > 449 Then LogroOro = True
    End Select
End If
'-----------------------------------------------------------
   

      
      'pluto:2.11
    If UserList(UserIndex).GranPoder > 0 Then daño = daño * 2
    'añadimos % de equipo
    'nati: cambio esto, ya no será por porcentaje.
   'daño = daño + CInt(Porcentaje(daño, DañoEquipoMagico(UserIndex)))
   daño = daño + DañoEquipoMagico(UserIndex)
    
     '¿arma equipada?
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        
    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo = 5 And Npclist(NpcIndex).NPCtype = 79 Then
            daño = daño * 5
            GoTo tuu
        'pluto:7.0 MENOS DAÑO SIN VARA
       ' ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo <> 13 Then
       ' daño = daño - CInt(Porcentaje(daño, 10))
        'End If
    
        'If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).SubTipo = 13 Then
            'daño = daño + CInt(Porcentaje(daño, ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Magia))
            'Else
            'daño = daño - CInt(Porcentaje(daño, 10))
    End If
  
End If

        
tuu:
'------------------------------
  'pluto:2.17 ettin y rey menos daño magias
'If (Npclist(NpcIndex).NPCtype = 77 Or Npclist(NpcIndex).NPCtype = 33) And daño > 0 Then daño = CInt(daño / 3)
      
        
        
        'pluto:2.3 quitar esto 1000 por un 0
    'quitar esto
    If UserList(UserIndex).flags.Privilegios > 0 Then daño = 0

    'pluto:2.4.5
If UserList(UserIndex).flags.Montura = 1 Then
Dim pl As Integer
Dim po As Integer
'Dim po As Byte
Dim nivk As Byte
Dim kk As Byte
po = UserList(UserIndex).flags.ClaseMontura
'If po = 1 Or po = 5 Then
'pluto:2.11
'If po = 1 Then kk = 2
'If po = 5 Then kk = 3
nivk = UserList(UserIndex).Montura.Nivel(po)
daño = daño + CInt(Porcentaje(daño, UserList(UserIndex).Montura.AtMagico(po))) + 1
'--------------

If UserList(UserIndex).Montura.AtMagico(po) > 0 Then pl = UserList(UserIndex).Montura.Golpe(po) Else pl = 0
'pluto:6.2
If UserList(UserIndex).Montura.Tipo(po) = 6 Then pl = UserList(UserIndex).Montura.Golpe(po)

Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - pl
'Call SendData(ToIndex, userindex, 0, "U2" & daño & "," & pl & "," & Npclist(NpcIndex).Char.CharIndex)
'Else
'Call SendData(ToIndex, UserIndex, 0, "U2" & daño)
End If
'End If
'-------

      
    
    Call InfoHechizo(UserIndex)
    b = True
    Call NpcAtacado(NpcIndex, UserIndex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    'pluto:2.8.0
    If Npclist(NpcIndex).NPCtype = 60 Then daño = daño - CInt(Porcentaje(daño, 50))
        
    'pluto:7.0 lo muevo detras para dar mas importancia a los modificadores
    daño = CInt(daño * ModMagia(UserList(UserIndex).clase))
    'nati: agrego la linea para que divida el golpe empleado al npc por magia. 10% = 1.1
    daño = CInt(daño / 1.1)
        
    daño = daño + Int(Porcentaje(daño, UserList(UserIndex).UserDañoMagiasRaza))

'[Tite] Pluto:6.0A Le aplico el skill daño magico
   ' daño = daño + CInt(Porcentaje(daño, (CInt(UserList(UserIndex).Stats.UserSkills(DañoMagia) / 10))))
   ' Call SubirSkill(UserIndex, DañoMagia)

'[\Tite]
    
    
    'pluto:7.0 Criticos de ciclopes
'If UserList(UserIndex).raza = "Ciclope" Then
 '   Dim probi As Integer
  '  probi = RandomNumber(1, 100) + CInt((UserList(UserIndex).Stats.UserSkills(suerte) / 40))
   ' If probi > 93 Then
    'Criti = 2
    'GoTo ciclo
    'End If
'End If
  
    
    
  'pluto:6.0A-----golpes criticos-------------
If Npclist(NpcIndex).GiveEXP < 37000 Or LogroOro = True Then
Dim cf As Integer

cf = 3500

'If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil > 0 Then cf = cf + 2000
'pluto:6.5--------------------
'Loco = RandomNumber(1, cf)
'If Loco < (UserList(UserIndex).Stats.UserSkills(suerte) * 5) Then Loco = (UserList(UserIndex).Stats.UserSkills(suerte) * 5)

'-------------------------
Critico = RandomNumber(1, cf) - (UserList(UserIndex).Stats.UserSkills(suerte) * 5)
If Critico < 60 Then Criti = 2
If Critico > 59 And Critico < 109 Then Criti = 3
If Critico > 108 And Critico < 118 Then Criti = 4
If Critico > 117 And Critico < 120 Then Criti = 5
 Else
 'pluto:6.5-----------
 'Loco = RandomNumber(1, cf + 7000)
 'If Loco < (UserList(UserIndex).Stats.UserSkills(suerte) * 10) Then Loco = (UserList(UserIndex).Stats.UserSkills(suerte) * 10)
'---------------------
Critico = RandomNumber(1, cf + 7000) - (UserList(UserIndex).Stats.UserSkills(suerte) * 10)
If Critico < 60 Then Criti = 2
If Critico > 59 And Critico < 109 Then Criti = 3
If Critico > 108 And Critico < 118 Then Criti = 4
End If
'------------------------------------------------
ciclo:
If UserList(UserIndex).flags.SegCritico = True Then Criti = 1

If Criti > 0 And Criti <> 5 Then daño = daño * Criti
'pluto:6.2 mortales no en piñatas y raids
If Criti = 5 And Npclist(NpcIndex).Raid = 0 And Npclist(NpcIndex).numero <> 664 Then Npclist(NpcIndex).Stats.MinHP = 0
    
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - _
    daño
If Npclist(NpcIndex).Stats.MinHP < 0 Then Npclist(NpcIndex).Stats.MinHP = 0

'pluto:2.10
'Call SendData(ToIndex, UserIndex, 0, "U2" & daño & "," & pl & "," & Npclist(NpcIndex).Char.CharIndex)
Call SendData(ToIndex, UserIndex, 0, "U2" & daño & "," & pl & "," & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Name & "," & Npclist(NpcIndex).Stats.MinHP & "," & Npclist(NpcIndex).Stats.MaxHP & "," & Criti)
'pluto:6.0A
If Npclist(NpcIndex).Raid > 0 Then

Dim nn As Byte
Dim MinPc As npc
MinPc = Npclist(NpcIndex)
Dim Porvida As Integer
Porvida = Int((Npclist(NpcIndex).Stats.MinHP * 100) / Npclist(NpcIndex).Stats.MaxHP)

Select Case Porvida

Case Is < 10
If RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 1 Then
    For nn = 1 To 5
    If RandomNumber(1, 100) > 20 Then Call SpawnNpc(MinPc.numero + 6, MinPc.Pos, True, False)
    Next
    RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 0
End If

Case Is < 20
If RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 2 Then
    For nn = 1 To 5
    If RandomNumber(1, 100) > 20 Then Call SpawnNpc(MinPc.numero + 6, MinPc.Pos, True, False)
    Next
    RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 1
End If
Case Is < 30
If RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 3 Then
    For nn = 1 To 5
    If RandomNumber(1, 100) > 20 Then Call SpawnNpc(MinPc.numero + 6, MinPc.Pos, True, False)
    Next
    RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 2
End If
Case Is < 40
If RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 4 Then
    For nn = 1 To 5
    If RandomNumber(1, 100) > 20 Then Call SpawnNpc(MinPc.numero + 6, MinPc.Pos, True, False)
    Next
    RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 3
End If
Case Is < 50
If RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 5 Then
    For nn = 1 To 5
    If RandomNumber(1, 100) > 20 Then Call SpawnNpc(MinPc.numero + 6, MinPc.Pos, True, False)
    Next
    RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 4
End If
Case Is < 60
If RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 6 Then
    For nn = 1 To 5
    If RandomNumber(1, 100) > 20 Then Call SpawnNpc(MinPc.numero + 6, MinPc.Pos, True, False)
    Next
    RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 5
End If
Case Is < 70
If RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 7 Then
    For nn = 1 To 5
    If RandomNumber(1, 100) > 20 Then Call SpawnNpc(MinPc.numero + 6, MinPc.Pos, True, False)
    Next
    RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 6
End If
Case Is < 80
If RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 8 Then
    For nn = 1 To 5
    If RandomNumber(1, 100) > 20 Then Call SpawnNpc(MinPc.numero + 6, MinPc.Pos, True, False)
    Next
    RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 7
End If
Case Is < 90
If RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 9 Then
    For nn = 1 To 5
    If RandomNumber(1, 100) > 20 Then Call SpawnNpc(MinPc.numero + 6, MinPc.Pos, True, False)
    Next
    RaidVivos(Npclist(NpcIndex).numero - 699).MiniRaids = 8
End If
 
End Select





'    If RandomNumber(1, 200) < Npclist(NpcIndex).Raid Then
'Dim recu As Integer
'recu = RandomNumber(1, Npclist(NpcIndex).Raid * 20)
'Call AddtoVar(Npclist(NpcIndex).Stats.MinHP, recu, Npclist(NpcIndex).Stats.MaxHP)
 '   Else
'recu = 0
 '   End If
'Call SendData(toParty, UserIndex, UserList(UserIndex).Pos.Map, "H4" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Stats.MinHP & "," & recu)
End If

'SendData ToIndex, UserIndex, 0, "||Causas " & daño & " de daño " & "(" & Npclist(NpcIndex).Stats.MinHP & "/" & Npclist(NpcIndex).Stats.MaxHP & ")" & FONTTYPENAMES.FONTTYPE_fight
   'pluto: npc en la casa
 If (Npclist(NpcIndex).Pos.Map = 171 Or Npclist(NpcIndex).Pos.Map = 177) And (Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP / 3) Then
 Dim ale
 ale = RandomNumber(1, 500)
Select Case ale
 'npc se quitaparalisis
 Case Is < 20
 If Npclist(NpcIndex).flags.Paralizado > 0 Then
 Npclist(NpcIndex).flags.Paralizado = 0
 Npclist(NpcIndex).Contadores.Paralisis = 0
 Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "TW" & 115)
 Call SendData(ToIndex, UserIndex, 0, "|| Los Espiritus de la casa han desparalizado al " & Npclist(NpcIndex).Name & "´" & FontTypeNames.FONTTYPE_TALK)
End If
 'Pluto:2.20 añado >0 // npc se cura
 Case 21 To 30
 If Npclist(NpcIndex).Stats.MinHP > 0 And Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP Then
 Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
 Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "TW" & 115)
 Call SendData2(ToPCArea, UserIndex, Npclist(NpcIndex).Pos.Map, 22, Npclist(NpcIndex).Char.CharIndex & "," & Hechizos(32).FXgrh & "," & Hechizos(32).loops)
 Call SendData(ToIndex, UserIndex, 0, "|| Los Espiritus de la casa han Sanado al " & Npclist(NpcIndex).Name & "´" & FontTypeNames.FONTTYPE_TALK)
 End If
 'npc saca npcs
 Case 31 To 40
 Call SpawnNpc(550, UserList(UserIndex).Pos, True, False)
 Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "TW" & 115)
 Call SendData(ToIndex, UserIndex, 0, "|| Los Espiritus de invocan una ayuda al " & Npclist(NpcIndex).Name & "´" & FontTypeNames.FONTTYPE_TALK)
 End Select
 End If
    
    
    If Npclist(NpcIndex).Stats.MinHP < 1 Then
      'Call MuereNpc(NpcIndex, UserIndex)
        'Npclist(NpcIndex).Stats.MinHP = 0
If Npclist(NpcIndex).Name = "Rey del Castillo" Or Npclist(NpcIndex).Name = "Defensor Fortaleza" Then Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
  Call MuereNpc(NpcIndex, UserIndex)
    End If

'End If
'ataque area

ElseIf Hechizos(hindex).SubeHP = 4 Then
    'pluto:6.5
    If Npclist(NpcIndex).Attackable = 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "L5")
        Exit Sub
    End If
   Dim X As Integer
Dim Y As Integer
Dim H As Integer
 
 
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
'p = MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex
   For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then

If MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex > 0 Then
    'pluto:2.19----------------------------------
    Dim Bc As Integer
    Bc = MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex
    'pluto:6.5
    If Npclist(Bc).flags.PoderEspecial2 > 0 Or Npclist(Bc).Raid > 0 Then GoTo alli
    '------------------------------------------------
    daño = RandomNumber(Hechizos(hindex).MinHP, Hechizos(hindex).MaxHP)
        'pluto:6.0----------------------------------------
If UserList(UserIndex).Remort = 0 Then
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
Else
    If UserList(UserIndex).clase = "Mago" Or UserList(UserIndex).clase = "Druida" Then
    'Dim Topito As Long
    Topito = UserList(UserIndex).Stats.ELV * 3.65
    If UserList(UserIndex).Stats.ELV > 45 Then Topito = 45 * 3.65
    daño = daño + Porcentaje(daño, Topito)
    Else
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    End If
End If
    '-------------------------------------------------    'pluto:2.3
    If UserList(UserIndex).flags.Privilegios > 0 Then daño = 0
  'pluto:6.0A quito rey daño magias
  'pluto:2.18 ettin menos daño magias
'If Npclist(Bc).NPCtype = 77 And daño > 0 Then daño = CInt(daño / 3)

    If Npclist(Bc).Attackable = 0 Then GoTo alli
    'pluto:2.18
    If Npclist(Bc).MaestroUser > 0 Then GoTo alli

    Call InfoHechizo(UserIndex)
    Call SendData2(ToPCArea, UserIndex, Npclist(Bc).Pos.Map, 22, Npclist(Bc).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)

    Call NpcAtacado(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex, UserIndex)
        If Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).flags.Snd2 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).flags.Snd2)
     b = True
    
   
    Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).Stats.MinHP = Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).Stats.MinHP - daño
    SendData ToIndex, UserIndex, 0, "||Le has causado " & daño & " puntos de daño a la criatura!" & "(" & Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).Stats.MinHP & "/" & Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).Stats.MaxHP & ")" & "´" & FontTypeNames.FONTTYPE_FIGHT
    If Npclist(Bc).Stats.MinHP < 1 Then
    Npclist(Bc).Stats.MinHP = 0
If Npclist(Bc).Name = "Rey del Castillo" Or Npclist(Bc).Name = "Defensor Fortaleza" Then Npclist(Bc).Stats.MinHP = Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex).Stats.MaxHP
  Call MuereNpc(Bc, UserIndex)
    End If
alli:
End If
End If
        
    Next X
    Next Y
  
  
    
   'ataque zona cercana usuario

ElseIf Hechizos(hindex).SubeHP = 3 Then
  'pluto:6.5
    If Npclist(NpcIndex).Attackable = 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
       Call SendData(ToIndex, UserIndex, 0, "L5")
       Exit Sub
    End If
    
    
    
  If Npclist(NpcIndex).Pos.X > UserList(UserIndex).Pos.X + 1 Or Npclist(NpcIndex).Pos.X < UserList(UserIndex).Pos.X - 1 Or Npclist(NpcIndex).Pos.Y > UserList(UserIndex).Pos.Y + 10 Or Npclist(NpcIndex).Pos.Y < UserList(UserIndex).Pos.Y - 10 Then
        Call SendData(ToIndex, UserIndex, 0, "L2")
        Exit Sub
    End If
  
  ' Dim X As Integer
'Dim Y As Integer
'Dim H As Integer
 
 
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
'p = MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex
   For Y = UserList(UserIndex).Pos.Y - 2 To UserList(UserIndex).Pos.Y + 2
            For X = UserList(UserIndex).Pos.X - 2 To UserList(UserIndex).Pos.X + 2
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
If MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex > 0 Then
 'pluto:2.19----------------------------------
    'Dim Bc As Integer
    Bc = MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex
    'pluto:6.5
    If Npclist(Bc).flags.PoderEspecial2 > 0 Or Npclist(Bc).Raid > 0 Then GoTo alli3
    '------------------------------------------------

    daño = RandomNumber(Hechizos(hindex).MinHP, Hechizos(hindex).MaxHP)
       'pluto:6.0----------------------------------------
If UserList(UserIndex).Remort = 0 Then
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
Else
    If UserList(UserIndex).clase = "Mago" Or UserList(UserIndex).clase = "Druida" Then
   ' Dim Topito As Long
    Topito = UserList(UserIndex).Stats.ELV * 3.65
    If UserList(UserIndex).Stats.ELV > 45 Then Topito = 45 * 3.65
    daño = daño + Porcentaje(daño, Topito)
    Else
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    End If
End If
    '-------------------------------------------------    'pluto:2.3
    If UserList(UserIndex).flags.Privilegios > 0 Then daño = 0
  'pluto:2.18 ettin y rey menos daño magias
'If (Npclist(Bc).NPCtype = 77 Or Npclist(Bc).NPCtype = 33) And daño > 0 Then daño = CInt(daño / 3)

    If Npclist(Bc).Attackable = 0 Then GoTo alli3
   'pluto:2.18
   If Npclist(Bc).MaestroUser > 0 Then GoTo alli3

   Call InfoHechizo(UserIndex)
    Call SendData2(ToPCArea, UserIndex, Npclist(Bc).Pos.Map, 22, Npclist(Bc).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    
        Call NpcAtacado(Bc, UserIndex)
    
    If Npclist(Bc).flags.Snd2 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(Bc).flags.Snd2)
     b = True
     
    Npclist(Bc).Stats.MinHP = Npclist(Bc).Stats.MinHP - _
    daño
    SendData ToIndex, UserIndex, 0, "||Le has causado " & daño & " puntos de daño a la criatura!" & "(" & Npclist(Bc).Stats.MinHP & "/" & Npclist(Bc).Stats.MaxHP & ")" & "´" & FontTypeNames.FONTTYPE_FIGHT
    If Npclist(Bc).Stats.MinHP < 1 Then
    Npclist(Bc).Stats.MinHP = 0
If Npclist(Bc).Name = "Rey del Castillo" Or Npclist(Bc).Name = "Defensor Fortaleza" Then Npclist(Bc).Stats.MinHP = Npclist(Bc).Stats.MaxHP
  Call MuereNpc(Bc, UserIndex)
    End If
alli3:
End If
End If
           
    Next X
    Next Y
    
    
  End If
 'pluto:2.5.0
 Exit Sub
errhandler:
Call LogError("Error en HechizoPropNPC: " & UserList(UserIndex).Name & " -> " & Npclist(NpcIndex).Name & " -> " & Hechizos(hindex).Nombre & " " & Err.Description)

End Sub
Sub InfoHechizo(ByVal UserIndex As Integer)

On Error GoTo fallo
    Dim H As Integer
    Dim HH As Byte
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    'Hechizos(H).FXgrh = 95
    'Hechizos(H).loops = 2
    
   ' Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, UserIndex)
   
   'pluto:6.0A------------------------------------------------------------
If UserList(UserIndex).flags.TargetUser > 0 Then
   
  Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 76, UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex & "," & H & "," & UserList(UserIndex).Char.CharIndex)

ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
   
  Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 76, Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & "," & H & "," & UserList(UserIndex).Char.CharIndex)
Else 'terreno
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 76, UserList(UserIndex).Char.CharIndex & "," & H & "," & UserList(UserIndex).Char.CharIndex)

End If
   '----------------------------------------------------------------------
  
   
   'Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||7°" & Hechizos(H).PalabrasMagicas & "°" & UserList(UserIndex).Char.CharIndex)

   ' Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(H).WAV)
    
   If UserList(UserIndex).flags.TargetUser > 0 Then
       Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
        Call SendData2(ToPCArea, UserIndex, Npclist(UserList(UserIndex).flags.TargetNpc).Pos.Map, 22, Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    End If
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            Call SendData(ToIndex, UserIndex, 0, "S5" & H & "," & UserList(UserList(UserIndex).flags.TargetUser).Name)
            Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "S6" & H & "," & UserList(UserIndex).Name)
            
            'Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).Name & FONTTYPENAMES.FONTTYPE_fight)
           'Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " " & Hechizos(H).TargetMsg & FONTTYPENAMES.FONTTYPE_fight)
        Else
            Call SendData(ToIndex, UserIndex, 0, "S4" & H)

            'Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).PropioMsg & FONTTYPENAMES.FONTTYPE_fight)
        End If
    ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "S7" & H)
    End If
    Exit Sub
fallo:
Call LogError("infohechizo " & Err.number & " D: " & Err.Description)

End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
On Error GoTo fallo
Dim HH As Integer
Dim H As Integer
Dim daño As Integer
Dim tempChr As Integer
Dim Topito As Long

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
tempChr = UserList(UserIndex).flags.TargetUser

'nati: Agrego esto para cuando te ataquen dejes de meditar.
If UserList(tempChr).flags.Meditando Then
    Call SendData(ToIndex, tempChr, 0, "G7")
    Call SendData2(ToIndex, tempChr, 0, 54)
    Call SendData2(ToIndex, tempChr, 0, 15, UserList(tempChr).Pos.X & "," & UserList(tempChr).Pos.Y)
    UserList(tempChr).flags.Meditando = False
    UserList(tempChr).Char.FX = 0
    UserList(tempChr).Char.loops = 0
    'pluto:bug meditar
    Call SendData2(ToMap, tempChr, UserList(tempChr).Pos.Map, 22, UserList(tempChr).Char.CharIndex & "," & 0 & "," & 0)
End If
'nati: Agrego esto para cuando te ataquen dejes de meditar.

'nati: Agrego esto para cuando te ataquen dejes de descansar.
If UserList(tempChr).flags.Descansar Then
Call SendData(ToIndex, tempChr, 0, "||Te levantas." & "´" & FontTypeNames.FONTTYPE_info)
UserList(tempChr).flags.Descansar = False
Call SendData2(ToIndex, tempChr, 0, 41)
End If
'nati: Agrego esto para cuando te ataquen dejes de descansar.

'pluto:6.0A
'If Hechizos(H).Noesquivar = 1 Then GoTo noss
'skill EVITA MAGIA
'Dim oo As Byte
'oo = RandomNumber(1, 100)
'Call SubirSkill(tempChr, EvitaMagia)
'If oo < CInt((UserList(tempChr).Stats.UserSkills(EvitaMagia) / 10) + 2) And UserList(tempChr).flags.Muerto = 0 Then
'Call SendData(ToIndex, UserIndex, 0, "|| Se ha Resistido a la Magia !!" & FONTTYPENAMES.FONTTYPE_fight)
'Call SendData(ToIndex, tempChr, 0, "|| Has Resistido una Magia !!" & FONTTYPENAMES.FONTTYPE_fight)
'b = True
'Exit Sub
'End If
'--------------------

noss:

'Hambre
If Hechizos(H).SubeHam = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHam, _
         daño, UserList(tempChr).Stats.MaxHam)
    'pluto:
    UserList(tempChr).flags.Hambre = 0
    UserList(tempChr).flags.Sed = 0
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de hambre." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de hambre." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
    If UserList(tempChr).Stats.MinHam < 1 Then UserList(tempChr).Stats.MinHam = 1
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has dejado con " & UserList(tempChr).Stats.MinHam & " puntos de hambre a " & UserList(tempChr).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha dejado con " & UserList(tempChr).Stats.MinHam & " puntos de hambre." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has dejado con " & UserList(tempChr).Stats.MinHam & " puntos de hambre." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    
End If

'Sed
If Hechizos(H).SubeSed = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    Call AddtoVar(UserList(tempChr).Stats.MinAGU, daño, _
         UserList(tempChr).Stats.MaxAGU)
         
    If UserIndex <> tempChr Then
      Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
      Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de sed." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Else
      Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de sed." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
     Call EnviarHambreYsed(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeSed = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de sed." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de sed." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
End If
'nati: agrego que si es ELFO DROW no pueda doparse
If Not UserList(tempChr).raza = "Elfo Oscuro" Then
' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    
    Call InfoHechizo(UserIndex)
'pluto:2.15
If UserList(tempChr).flags.DuracionEfecto = 0 Then
Call SendData(ToIndex, tempChr, 0, "S1")
End If
    
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), daño, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) + 13)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Agilidad) = UserList(tempChr).Stats.UserAtributos(Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Agilidad) = MINATRIBUTOS
    b = True
    
End If
ElseIf Len(UserList(tempChr).Padre) > 0 Then ' <--- AGREGANDO ESTO, LE ESTOY DICIENDO QUE SI ES BEBE, SI QUE SE PUEDA DOPAR.
' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    
    Call InfoHechizo(UserIndex)
'pluto:2.15
If UserList(tempChr).flags.DuracionEfecto = 0 Then
Call SendData(ToIndex, tempChr, 0, "S1")
End If
    
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), daño, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) + 13)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Agilidad) = UserList(tempChr).Stats.UserAtributos(Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Agilidad) = MINATRIBUTOS
    b = True
    
End If
End If
'agrego que si es ELFO DROW no pueda doparse
' <-------- Fuerza ---------->
'nati: agrego que si es ENANO no se dope.
If Not UserList(tempChr).raza = "Enano" Then
If Hechizos(H).SubeFuerza = 1 Then
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    'pluto:2.15
    If UserList(tempChr).flags.DuracionEfecto = 0 Then
    Call SendData(ToIndex, tempChr, 0, "S1")
    End If
    
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Fuerza), daño, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) + 13)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Fuerza) = UserList(tempChr).Stats.UserAtributos(Fuerza) - daño
    If UserList(tempChr).Stats.UserAtributos(Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Fuerza) = MINATRIBUTOS
    b = True
    
End If
ElseIf Len(UserList(tempChr).Padre) > 0 Then ' <--- AGREGANDO ESTO, LE ESTOY DICIENDO QUE SI TIENE PADRE (ENTONCES ES HIJO), SI QUE SE PUEDA DOPAR.
If Hechizos(H).SubeFuerza = 1 Then
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    'pluto:2.15
    If UserList(tempChr).flags.DuracionEfecto = 0 Then
    Call SendData(ToIndex, tempChr, 0, "S1")
    End If
    
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Fuerza), daño, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) + 13)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Fuerza) = UserList(tempChr).Stats.UserAtributos(Fuerza) - daño
    If UserList(tempChr).Stats.UserAtributos(Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Fuerza) = MINATRIBUTOS
    b = True
    
End If
End If
'Salud
If Hechizos(H).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
       'pluto:6.0----------------------------------------
If UserList(UserIndex).Remort = 0 Then
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
Else
    If UserList(UserIndex).clase = "Mago" Or UserList(UserIndex).clase = "Druida" Then
    'Dim Topito As Long
    Topito = UserList(UserIndex).Stats.ELV * 3.65
    If UserList(UserIndex).Stats.ELV > 45 Then Topito = 45 * 3.65
    daño = daño + Porcentaje(daño, Topito)
    Else
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    End If
End If
    '-------------------------------------------------
    Call InfoHechizo(UserIndex)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHP, daño, _
         UserList(tempChr).Stats.MaxHP)
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vida." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de vida." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
    b = True
    
ElseIf Hechizos(H).SubeHP = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    If UserIndex = tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "L6")
        Exit Sub
    End If
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    'PLUTO
    If UCase$(Hechizos(H).Nombre) = "RAYO GM" Then
    'pluto:2.14
        Call LogGM(UserList(UserIndex).Name, "RAYO GM: " & UserList(tempChr).Name)
        daño = 500
    End If

'------------------------------
'pluto:7.0 extra monturas subido para calculo sobre daño base
If UserList(UserIndex).flags.Montura = 1 Then

Dim oo As Integer

    oo = UserList(UserIndex).flags.ClaseMontura

    'pluto:7.0----------
    daño = daño + CInt(Porcentaje(daño, UserList(UserIndex).Montura.AtMagico(oo))) + 1
    '------------------
    If daño < 1 Then daño = 1
End If

If UserList(tempChr).flags.Montura = 1 Then
    oo = UserList(tempChr).flags.ClaseMontura
    'kk = 0
'If oo = 1 Then kk = 2
'If oo = 5 Then kk = 3
    'nivk = UserList(tempChr).Montura.Nivel(oo)
    daño = daño - CInt(Porcentaje(daño, UserList(tempChr).Montura.DefMagico(oo))) - 1
    If daño < 1 Then daño = 1
End If
'------------fin pluto:2.13-------------------





   ' daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    'pluto:6.0----------------------------------------
If UserList(UserIndex).Remort = 0 Then
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
Else
    If UserList(UserIndex).clase = "Mago" Or UserList(UserIndex).clase = "Druida" Then
    'Dim Topito As Long
    Topito = UserList(UserIndex).Stats.ELV * 3.65
    If UserList(UserIndex).Stats.ELV > 45 Then Topito = 45 * 3.65
    daño = daño + Porcentaje(daño, Topito)
    Else
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    End If
End If
    '-------------------------------------------------
    


    'pluto:6.0A Skills------------------------
    'daño = daño + CInt(Porcentaje(daño, (CInt(UserList(UserIndex).Stats.UserSkills(DañoMagia) / 10))))
    'daño = daño - CInt(Porcentaje(daño, (CInt(UserList(tempChr).Stats.UserSkills(DefMagia) / 10))))
    'Call SubirSkill(tempChr, DefMagia)
    'Call SubirSkill(UserIndex, DañoMagia)
'---------------------------------------------------------------
    If UserList(tempChr).flags.Angel > 0 Then daño = CInt(daño - (daño * 0.5))
    If UserList(UserIndex).flags.Demonio > 0 Then daño = CInt(daño + (daño * 0.5))
'pluto:2.11
    If UserList(UserIndex).GranPoder > 0 Then daño = CInt(daño + daño)
      'pluto:2.16
    If UserList(tempChr).flags.Protec > 0 Then daño = daño - CInt(Porcentaje(daño, UserList(tempChr).flags.Protec))
'pluto:2.4.1
    Dim obj As ObjData
    If UserList(tempChr).Invent.AnilloEqpObjIndex > 0 Then
        If ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).SubTipo = 4 Then daño = daño - CInt(daño / 6)
    End If
'pluto:7.0
 If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
daño = daño - ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).Defmagica
If daño < 1 Then daño = 1
End If
'nati: Cuestion de balance, si no lleva ropa le hara un 15% de daño extra.
If UserList(tempChr).Invent.ArmourEqpObjIndex = 0 Then
 daño = daño + CInt(Porcentaje(daño, 15))
 End If
'nati: Cuestion de balance, si no lleva ropa le hara un 15% de daño extra.

    
'pluto:6.0A---------------------
   'If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        'If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo = 13 Then
           ' daño = daño + CInt(Porcentaje(daño, ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Magia))
           'Else
              'daño = daño - CInt(Porcentaje(daño, 10))
           
    'End If
  'añadimos % de equipo
  'nati: cambio esto, ya no será por porcentaje.
   'daño = daño + CInt(Porcentaje(daño, DañoEquipoMagico(UserIndex)))
   daño = daño + DañoEquipoMagico(UserIndex)
   
'pluto:7.0 MENOS DAÑO SIN VARA
'If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 '   If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo <> 13 Then
  '  daño = daño - CInt(Porcentaje(daño, 10))
   ' End If
'End If

tuu:


'pluto:7.0 lo muevo detras para aumentar importancia de modificadores
    daño = CInt(daño * ModMagia(UserList(UserIndex).clase))
    daño = CInt(daño / ModMagia(UserList(tempChr).clase))
    daño = daño - CInt(Porcentaje(daño, UserList(tempChr).UserDefensaMagiasRaza))
    daño = daño + CInt(Porcentaje(daño, UserList(UserIndex).UserDañoMagiasRaza))
'------------------------------------------------------------------------------
'nati: agrego el +20% del Berseker en magias
If UserList(tempChr).raza = "Orco" And UserList(tempChr).Counters.Morph > 0 Then
daño = daño + CInt(Porcentaje(daño, 20))
End If
'nati: fin
'nati: agrego el -20% del Berseker
If UserList(UserIndex).raza = "Orco" And UserList(UserIndex).Counters.Morph > 0 Then
daño = daño + CInt(Porcentaje(daño, 20))
End If
'nati:fin berseker
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - daño
    
    Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vida." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    
    
    'pluto:7.0 10% quedar 1 vida en ciclopes
 If UserList(tempChr).Stats.MinHP < 1 And UserList(tempChr).raza = "Ciclope" Then
     Dim bup As Byte
     bup = RandomNumber(1, 10)
     If bup = 8 Then UserList(tempChr).Stats.MinHP = 1
 End If
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, UserIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, UserIndex)
        'Call UserDie(tempChr)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeHP = 4 Then
    'pj area
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    Dim X As Integer
    Dim Y As Integer
    Dim tmpIndex As Integer
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    If UserIndex = tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "L6")
        Exit Sub
    End If
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    '[MerLiNz:X]
    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
        For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex > 0 Then
                    If Criminal(UserIndex) = Criminal(MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex) Then GoTo nop
                    tmpIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex
                    If tmpIndex = UserIndex Then GoTo nop
                    If UserList(tmpIndex).flags.Privilegios > 0 Then GoTo nop
    'pluto:hoy
                    If UserList(tmpIndex).flags.Muerto > 0 Then GoTo nop
    
                    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
                    
                    
               'pluto:7.0 extra monturas subido arriba
                    If UserList(UserIndex).flags.Montura = 1 Then
                        oo = UserList(UserIndex).flags.ClaseMontura

                        'pluto:7.0---------
                        daño = daño + CInt(Porcentaje(daño, UserList(UserIndex).Montura.AtMagico(oo))) + 1
                        '--------------
                        If daño < 1 Then daño = 1
                    End If
                    If UserList(tempChr).flags.Montura = 1 Then
                        oo = UserList(tempChr).flags.ClaseMontura
'kk = 0
'If oo = 1 Then kk = 2
'If oo = 5 Then kk = 3
                       ' nivk = UserList(tempChr).Montura.Nivel(oo)
                        daño = daño - CInt(Porcentaje(daño, UserList(tempChr).Montura.DefMagico(oo))) - 1
                        If daño < 1 Then daño = 1
                    End If

'------------fin pluto:2.13-------------------
                    
                    
                    
                    
                    'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
  'pluto:6.0----------------------------------------
If UserList(UserIndex).Remort = 0 Then
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
Else
    If UserList(UserIndex).clase = "Mago" Or UserList(UserIndex).clase = "Druida" Then
    'Dim Topito As Long
    Topito = UserList(UserIndex).Stats.ELV * 3.65
    If UserList(UserIndex).Stats.ELV > 45 Then Topito = 45 * 3.65
    daño = daño + Porcentaje(daño, Topito)
    Else
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    End If
End If
    '-------------------------------------------------
 
 'pluto:2.18
                    daño = daño - CInt(Porcentaje(daño, UserList(tmpIndex).UserDefensaMagiasRaza))
  
   
   'If UserList(tmpIndex).raza = "Elfo" Then daño = daño - CInt(Porcentaje(daño, 8))
    'If UserList(tmpIndex).raza = "Humano" Then daño = daño - CInt(Porcentaje(daño, 5))
    'If UserList(tmpIndex).raza = "Gnomo" Then daño = daño - CInt(Porcentaje(daño, 15))
    'If UserList(tmpIndex).raza = "Elfo Oscuro" Then daño = daño - CInt(Porcentaje(daño, 5))
  'pluto:6.0A Skills---------------
                   ' daño = daño + CInt(Porcentaje(daño, (CInt(UserList(UserIndex).Stats.UserSkills(DañoMagia) / 10))))
                    'daño = daño - CInt(Porcentaje(daño, (CInt(UserList(tmpIndex).Stats.UserSkills(DefMagia) / 10))))
                    'Call SubirSkill(tmpIndex, DefMagia)
                    'Call SubirSkill(UserIndex, DañoMagia)
  '--------------------------------
                    If UserList(tmpIndex).flags.Angel > 0 Then daño = CInt(daño - (daño * 0.5))
                    If UserList(UserIndex).flags.Demonio > 0 Then daño = CInt(daño + (daño * 0.5))
'pluto:2.11
                    If UserList(UserIndex).GranPoder > 0 Then daño = CInt(daño + daño)
      'pluto:2.16
                    If UserList(tmpIndex).flags.Protec > 0 Then daño = daño - CInt(Porcentaje(daño, UserList(tmpIndex).flags.Protec))

'pluto:2.4.1

                    If UserList(tmpIndex).Invent.AnilloEqpObjIndex > 0 Then
                        If ObjData(UserList(tmpIndex).Invent.AnilloEqpObjIndex).SubTipo = 4 Then daño = daño - CInt(daño / 5)
                    End If
'pluto:7.0
 If UserList(tmpIndex).Invent.ArmourEqpObjIndex > 0 Then
daño = daño - ObjData(UserList(tmpIndex).Invent.ArmourEqpObjIndex).Defmagica
If daño < 1 Then daño = 1
End If

                    Call InfoHechizo(UserIndex)
                    Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)

                    UserList(tmpIndex).Stats.MinHP = UserList(tmpIndex).Stats.MinHP - daño
    
                    Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de vida a " & UserList(MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
                    Call SendData(ToIndex, MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vida." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    '[\END]
    'Muere
                    If UserList(MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex).Stats.MinHP < 1 Then
                        Call ContarMuerte(MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex, UserIndex)
                        UserList(MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex).Stats.MinHP = 0
                        Call ActStats(MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex, UserIndex)
        'Call UserDie(MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex)
                    End If
    
                    b = True
nop:
                End If
            End If

        Next X
    Next Y
    
'cercano usuario zona
ElseIf Hechizos(H).SubeHP = 3 Then
    'pj area
 
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
'[MerLiNz:X]
    HH = MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex
'[\END]
    If UserIndex = tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "L6")
        Exit Sub
    End If
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    For Y = UserList(UserIndex).Pos.Y - 2 To UserList(UserIndex).Pos.Y + 2
        For X = UserList(UserIndex).Pos.X - 2 To UserList(UserIndex).Pos.X + 2
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex > 0 Then
                    If Criminal(UserIndex) = Criminal(MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex) Then GoTo nop2
                    tmpIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex
                    If UserList(tmpIndex).flags.Privilegios > 0 Then GoTo nop2
    'pluto:hoy
                    If UserList(tmpIndex).flags.Muerto > 0 Then GoTo nop
    
                    If tmpIndex = UserIndex Then GoTo nop2
                    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
                    
                    'pluto:7.0 extra monturas subido arriba
                    If UserList(UserIndex).flags.Montura = 1 Then
                        oo = UserList(UserIndex).flags.ClaseMontura

                        daño = daño - CInt(Porcentaje(daño, UserList(UserIndex).Montura.DefMagico(oo))) - 1
                        If daño < 1 Then daño = 1
                    End If
'------------fin pluto:2.4-------------------
                    
                    
                    
                    
                    'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    'pluto:6.0----------------------------------------
If UserList(UserIndex).Remort = 0 Then
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
Else
    If UserList(UserIndex).clase = "Mago" Or UserList(UserIndex).clase = "Druida" Then
    'Dim Topito As Long
    Topito = UserList(UserIndex).Stats.ELV * 3.65
    If UserList(UserIndex).Stats.ELV > 45 Then Topito = 45 * 3.65
    daño = daño + Porcentaje(daño, Topito)
    Else
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    End If
End If
    '-------------------------------------------------
   'pluto:2.18
                    daño = daño - CInt(Porcentaje(daño, UserList(tmpIndex).UserDefensaMagiasRaza))

 'pluto:6.0A Skills
'daño = daño + CInt(Porcentaje(daño, (CInt(UserList(UserIndex).Stats.UserSkills(DañoMagia) / 10))))
'daño = daño - CInt(Porcentaje(daño, (CInt(UserList(tmpIndex).Stats.UserSkills(DefMagia) / 10))))
'----------------------------------------
                    If UserList(tmpIndex).flags.Angel > 0 Then daño = CInt(daño - (daño * 0.5))
                    If UserList(UserIndex).flags.Demonio > 0 Then daño = CInt(daño + (daño * 0.5))
'pluto:2.11
                    If UserList(UserIndex).GranPoder > 0 Then daño = CInt(daño + daño)
      'pluto:2.16
                    If UserList(tmpIndex).flags.Protec > 0 Then daño = daño - CInt(Porcentaje(daño, UserList(tmpIndex).flags.Protec))

'pluto:2.4.1

                    If UserList(tmpIndex).Invent.AnilloEqpObjIndex > 0 Then
                        If ObjData(UserList(tmpIndex).Invent.AnilloEqpObjIndex).SubTipo = 4 Then daño = daño - CInt(daño / 5)
                    End If
'pluto:7.0
 If UserList(tmpIndex).Invent.ArmourEqpObjIndex > 0 Then
daño = daño - ObjData(UserList(tmpIndex).Invent.ArmourEqpObjIndex).Defmagica
If daño < 1 Then daño = 1
End If

                    Call InfoHechizo(UserIndex)
                    Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(tmpIndex).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    
                    UserList(tmpIndex).Stats.MinHP = UserList(tmpIndex).Stats.MinHP - daño
    
                    Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de vida a " & UserList(tmpIndex).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
                    Call SendData(ToIndex, tmpIndex, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vida." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    
    'Muere
                    If UserList(tmpIndex).Stats.MinHP < 1 Then
                        Call ContarMuerte(tmpIndex, UserIndex)
                        UserList(tmpIndex).Stats.MinHP = 0
                        Call ActStats(tmpIndex, UserIndex)
        'Call UserDie(tmpIndex)
                    End If
    
                    b = True
nop2:
                End If
            End If
        Next X
    Next Y

End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinMAN, daño, UserList(tempChr).Stats.MaxMAN)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de mana." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de mana." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de mana." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de mana." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
    
   
    
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño
    
    
    
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinSta, daño, _
         UserList(tempChr).Stats.MaxSta)
    If UserIndex <> tempChr Then
         Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
         Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vitalidad." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & daño & " puntos de vitalidad." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then



    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vitalidad." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & daño & " puntos de vitalidad." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If

'Habilidades Pirata
If Hechizos(H).Nombre = "¡Al Abordaje!" Then

End If


Exit Sub
fallo:
Call LogError("hechizopropiousuario " & Err.number & " D: " & Err.Description)

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo fallo
'Call LogTarea("Sub UpdateUserHechizos")

Dim loopc As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(UserIndex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For loopc = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(loopc) > 0 Then
            Call ChangeUserHechizo(UserIndex, loopc, UserList(UserIndex).Stats.UserHechizos(loopc))
        Else
            Call ChangeUserHechizo(UserIndex, loopc, 0)
        End If

Next loopc

End If
Exit Sub
fallo:
Call LogError("updateuserhechizos " & Err.number & " D: " & Err.Description)

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)
On Error GoTo fallo
UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo
If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
    Call SendData2(ToIndex, UserIndex, 0, 34, Slot & "," & Hechizo)
Else
    Call SendData2(ToIndex, UserIndex, 0, 34, Slot & "," & "0")
End If

Exit Sub
fallo:
Call LogError("changeuserhechizo " & Err.number & " D: " & Err.Description)

End Sub

Sub HabilidadesPirata(ByVal UserIndex As Integer, ByVal Hechizo As Integer)



End Sub
