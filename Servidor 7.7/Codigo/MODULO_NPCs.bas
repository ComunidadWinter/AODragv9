Attribute VB_Name = "NPCs"
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
On Error GoTo fallo
Dim i As Integer
UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
For i = 1 To MAXMASCOTAS
  If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
     UserList(UserIndex).MascotasIndex(i) = 0
     UserList(UserIndex).MascotasType(i) = 0
     Exit For
  End If
Next i
Exit Sub
fallo:
Call LogError("quitarmascota " & Err.number & " D: " & Err.Description)

End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer, ByVal Mascota As Integer)
On Error GoTo fallo
Dim i As Integer

Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1

'For i = 1 To UBound(Npclist(Maestro).Criaturas)
'  If Npclist(Maestro).Criaturas(i).NpcIndex = Mascota Then
'     Npclist(Maestro).Criaturas(i).NpcIndex = 0
'     Npclist(Maestro).Criaturas(i).NpcName = ""
'     Exit For
'  End If
'Next i
Exit Sub
fallo:
Call LogError("quitarmascotanpc " & Err.number & " D: " & Err.Description & " maest: " & Maestro)

End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim n As Byte
'Call LogTarea("Sub MuereNpc")
Dim Expdemas As Byte
Dim MinPc As npc
MinPc = Npclist(NpcIndex)
'pluto:6.0
Dim PosPuerta As WorldPos
PosPuerta.Map = MinPc.Pos.Map
If MinPc.Pos.Map <> 185 Then
PosPuerta.X = 41
PosPuerta.Y = 76
Else
PosPuerta.X = 51
PosPuerta.Y = 50
 End If
 
 'pluto:6.0A
 'If MiNPC.MaestroNpc > 0 Then
' Npclist(MiNPC.MaestroNpc).Mascotas = Npclist(MiNPC.MaestroNpc).Mascotas - 1
 'End If
 
 
 
'Quitamos el npc
'castillo
'pluto:2.4.1 añado defensor fortaleza
'If MiNPC.NPCtype = 79 Then MiNPC.Stats.MinHP = MiNPC.Stats.MaxHP
  

  
If MinPc.NPCtype = 33 Or MinPc.NPCtype = 61 Then
    MinPc.Stats.MinHP = MinPc.Stats.MaxHP
Else
    Call QuitarNPC(NpcIndex)
End If
   
'pluto:2.18
If MinPc.NPCtype = 33 Or MinPc.NPCtype = 78 Or MinPc.NPCtype = 77 Or MinPc.NPCtype = 61 Then
    Select Case MinPc.Pos.Map
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
   
' End If
'comprobar sala invocacion
  
If MinPc.MaestroUser = 0 Then



 'pluto:7.0 misiones activas----------------------------
    
    If MinPc.Raid > 0 And UserList(UserIndex).flags.party = True And UserList(UserIndex).flags.partyNum > 0 Then
'es raid mision para todo grupo (todo esto es nuevo)
    Dim Npart As Byte
    Npart = UserList(UserIndex).flags.partyNum
    Dim nn As Byte
For nn = 1 To partylist(Npart).numMiembros
    If UserList(partylist(Npart).miembros(nn).ID).Mision.estado > 0 Then
    If UserList(partylist(Npart).miembros(nn).ID).Mision.NEnemigos > 0 Then
    For n = 1 To UserList(partylist(Npart).miembros(nn).ID).Mision.NEnemigos
        If UserList(partylist(Npart).miembros(nn).ID).Mision.Enemigo(n) = MinPc.numero Then
        UserList(partylist(Npart).miembros(nn).ID).Mision.NEnemigosConseguidos(n) = UserList(partylist(Npart).miembros(nn).ID).Mision.NEnemigosConseguidos(n) + 1
        If UserList(partylist(Npart).miembros(nn).ID).Mision.NEnemigosConseguidos(n) > UserList(partylist(Npart).miembros(nn).ID).Mision.EnemigoCantidad(n) Then UserList(partylist(Npart).miembros(nn).ID).Mision.NEnemigosConseguidos(n) = UserList(partylist(Npart).miembros(nn).ID).Mision.EnemigoCantidad(n)
        End If
    Next
    
    End If 'estado
    End If 'enemigos
Next nn
 
    Else 'no es raid o no hay party -> mision individual (todo igual que antes)
    
     If UserList(UserIndex).Mision.estado > 0 Then
 If UserList(UserIndex).Mision.NEnemigos > 0 Then
    For n = 1 To UserList(UserIndex).Mision.NEnemigos
        If UserList(UserIndex).Mision.Enemigo(n) = MinPc.numero Then
        UserList(UserIndex).Mision.NEnemigosConseguidos(n) = UserList(UserIndex).Mision.NEnemigosConseguidos(n) + 1
        If UserList(UserIndex).Mision.NEnemigosConseguidos(n) > UserList(UserIndex).Mision.EnemigoCantidad(n) Then UserList(UserIndex).Mision.NEnemigosConseguidos(n) = UserList(UserIndex).Mision.EnemigoCantidad(n)
        End If
    Next
 End If
 End If
     
    End If
 
  '----------------------------------------













    If MinPc.Pos.Map = mapi Then MapInfo(mapi).invocado = 0
    'comprobar castillo clanes
'----------------------------------------------------------------------------------------------
    'nati: Esto Obtiene el numero en el guildinfo.
Dim TotalClanes As String
Dim NumGuild As Integer
Dim RevisoGuild As String
Dim Conquistador As String
    'nati: Esto Obtiene el numero en el guildinfo.
'----------------------------------------------------------------------------------------------
    
'pluto:2.4.1 fortaleza
   ' If MinPc.Pos.Map = 185 And MinPc.NPCtype = 61 And UserList(UserIndex).GuildInfo.GuildName <> "" Then
    
'nati: NUEVO SISTEMA DE PUNTOS AL CONQUISTAR
    'nati: Esto Obtiene el dueño y numero antes de conquistar
Dim DueñoC As String
Dim NumGuildD As Integer
Dim RevisoGuildD As String
If MinPc.Pos.Map = 185 And MinPc.NPCtype = 61 And UserList(UserIndex).GuildInfo.GuildName <> "" Then
fortaleza = UserList(UserIndex).GuildInfo.GuildName
DueñoC = fortaleza
TotalClanes = Guilds.Count
For NumGuildD = 1 To TotalClanes
RevisoGuildD = Guilds(NumGuildD).GuildName
If RevisoGuildD = DueñoC Then
Exit For
End If
Next
miembros = Guilds(NumGuildD).Members.Count
'nati: aquí cargo el clan que conquista
Conquistador = UserList(UserIndex).GuildInfo.GuildName
For NumGuild = 1 To TotalClanes
If Conquistador = Guilds(NumGuild).GuildName Then
Exit For
End If
Next
PorcentajeC = 60
variablepuntos = 1
Dim X As Integer
Dim puntosX As Double
Dim PuntosGuild As Double
Dim SumaPuntosC As Double
For X = 1 To miembros
PorcentajeC = CInt(PorcentajeC) - variablepuntos
Next X
puntosX = 15 + CInt(PorcentajeC) / 2
PuntosGuild = Guilds(NumGuild).Reputation
SumaPuntosC = PuntosGuild + puntosX

Guilds(NumGuild).Reputation = Round(Guilds(NumGuild).Reputation + puntosX)
Guilds(NumGuildD).Reputation = Round(Guilds(NumGuildD).Reputation - puntosX)
If Guilds(NumGuildD).Reputation < 0 Then Guilds(NumGuildD).Reputation = 0
'Call WriteVar(App.Path & "\Guilds\GuildsInfo.inf", "GUILD" & NumGuild, "Rep", PuntosGuild + puntosX)
'nati: NUEVO SISTEMA DE PUNTOS AL CONQUISTAR

        'pluto:6.0A
        UserList(UserIndex).Stats.Fama = UserList(UserIndex).Stats.Fama + 10
        
        fortaleza = UserList(UserIndex).GuildInfo.GuildName
'pluto:2.4
        'UserList(UserIndex).Stats.PClan = UserList(UserIndex).Stats.PClan + puntosX
        date5 = Date
        hora5 = Time
        'Call BDDConquistanCastillo("fortaleza", fortaleza)
        Call SendData(ToAll, 0, 0, "|| El CLAN " & UCase$(UserList(UserIndex).GuildInfo.GuildName) & " HA CONQUISTADO LA FORTALEZA" & "´" & FontTypeNames.FONTTYPE_talk)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "fortaleza", UserList(UserIndex).GuildInfo.GuildName)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "date5", Date)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "hora5", Time)
        'fortaleza = UserList(UserIndex).GuildInfo.GuildName
        'pluto:6.9-------------------------------------------------
        'Select Case UCase$(UserList(UserIndex).GuildInfo.GuildName)
        'Case "PT AMO"
        'Call SendData(ToAll, 0, 0, "TW" & "ptamo")
        'Case "BLACKLIST"
        'Call SendData(ToAll, 0, 0, "TW" & "black")
        'Case "UNDEAD"
        'Call SendData(ToAll, 0, 0, "TW" & "undead")
        'Case Else
        Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
        'End Select
       '----------------------------------------------------------
        MinutosFortaleza = 0
    End If

'pluto:6.0A --------------muere rey-----------------------
    If MinPc.NPCtype = 33 Then
       UserList(UserIndex).Stats.Fama = UserList(UserIndex).Stats.Fama + 3
 
        
        If MapData(MinPc.Pos.Map, PosPuerta.X, PosPuerta.Y).UserIndex > 0 Then
            Call WarpUserChar(MapData(MinPc.Pos.Map, PosPuerta.X, PosPuerta.Y).UserIndex, MinPc.Pos.Map, 43, 83, True)
        End If
        If MapData(MinPc.Pos.Map, PosPuerta.X, PosPuerta.Y).NpcIndex > 0 Then Call QuitarNPC(MapData(MinPc.Pos.Map, PosPuerta.X, PosPuerta.Y).NpcIndex)

'pluto:2.18
        'MapData(Pospuerta.Map - 102, 58, 50).TileExit.Map = Pospuerta.Map
        'MapData(Pospuerta.Map - 102, 58, 50).TileExit.X = 43
        'MapData(Pospuerta.Map - 102, 58, 50).TileExit.Y = 26
'---------

        Call SpawnNpc(157, PosPuerta, False, False)
        MapData(MinPc.Pos.Map, PosPuerta.X + 1, PosPuerta.Y).Blocked = 1
        MapData(MinPc.Pos.Map, PosPuerta.X - 1, PosPuerta.Y).Blocked = 1
        Call Bloquear(ToMap, UserIndex, MinPc.Pos.Map, MinPc.Pos.Map, PosPuerta.X - 1, PosPuerta.Y, 1)
        Call Bloquear(ToMap, UserIndex, MinPc.Pos.Map, MinPc.Pos.Map, PosPuerta.X + 1, PosPuerta.Y, 1)
        Call SendData(ToAll, 0, 0, "C5")
    End If
    
    'pluto:6.0A
'si muere puerta, se coloca ettin
    If MinPc.NPCtype = 78 Then
        Dim PosEttin As WorldPos
        PosEttin.X = 51
        PosEttin.Y = 76
        PosEttin.Map = 185



        If MapData(MinPc.Pos.Map, PosEttin.X, PosEttin.Y).UserIndex > 0 Then
            Call WarpUserChar(MapData(MinPc.Pos.Map, PosEttin.X, PosEttin.Y).UserIndex, MinPc.Pos.Map, 51, 83, True)
        End If
        If MapData(MinPc.Pos.Map, PosEttin.X, PosEttin.Y).NpcIndex > 0 Then Call QuitarNPC(MapData(MinPc.Pos.Map, PosEttin.X, PosEttin.Y).NpcIndex)

'pluto:2.18
        'MapData(PosEttin.Map, 51, 50).TileExit.Map = PosEttin.Map
        'MapData(PosEttin.Map, 58, 50).TileExit.X = 40
        'MapData(PosEttin.Map, 58, 50).TileExit.Y = 53
    
'--------------
        'pluto:6.0A----------------------------------
        If MinPc.Pos.Map = 185 Then
        Call SpawnNpc(666, PosEttin, False, False)
        'Call SpawnNpc(157, Pospuerta, False, False)
'bloquea hueco pegado al ettin
        MapData(MinPc.Pos.Map, PosEttin.X + 1, PosEttin.Y - 1).Blocked = 1
        MapData(MinPc.Pos.Map, PosEttin.X - 1, PosEttin.Y - 1).Blocked = 1
        Call Bloquear(ToMap, UserIndex, MinPc.Pos.Map, MinPc.Pos.Map, PosEttin.X - 1, PosEttin.Y - 1, 1)
        Call Bloquear(ToMap, UserIndex, MinPc.Pos.Map, MinPc.Pos.Map, PosEttin.X + 1, PosEttin.Y - 1, 1)
        End If
        '---------------------------------------
'desbloquea puerta
        MapData(MinPc.Pos.Map, PosPuerta.X + 1, PosPuerta.Y).Blocked = 0
        MapData(MinPc.Pos.Map, PosPuerta.X - 1, PosPuerta.Y).Blocked = 0
        Call Bloquear(ToMap, UserIndex, MinPc.Pos.Map, MinPc.Pos.Map, PosPuerta.X - 1, PosPuerta.Y, 0)
        Call Bloquear(ToMap, UserIndex, MinPc.Pos.Map, MinPc.Pos.Map, PosPuerta.X + 1, PosPuerta.Y, 0)
        Call SendData(ToAll, 0, 0, "V9")
    End If
' si muere el ettin se coloca puerta
If MinPc.NPCtype = 77 Then
        PosEttin.X = 51
        PosEttin.Y = 76
        PosEttin.Map = 185
  If MapData(185, PosPuerta.X, PosPuerta.Y).UserIndex > 0 Then
            Call WarpUserChar(MapData(185, PosPuerta.X, PosPuerta.Y).UserIndex, 185, 50, 73, True)
        End If
        If MapData(185, PosPuerta.X, PosPuerta.Y).NpcIndex > 0 Then Call QuitarNPC(MapData(185, PosPuerta.X, PosPuerta.Y).NpcIndex)

'pluto:2.18
        'MapData(PosEttin.Map, 51, 50).TileExit.Map = PosEttin.Map
        'MapData(PosEttin.Map, 58, 50).TileExit.X = 40
        'MapData(PosEttin.Map, 58, 50).TileExit.Y = 53
    
'---------

        Call SpawnNpc(157, PosPuerta, False, False)
        'Call SpawnNpc(157, Pospuerta, False, False)
'bloquea hueco pegado a la puerta
        MapData(185, PosPuerta.X + 1, PosPuerta.Y).Blocked = 1
        MapData(185, PosPuerta.X - 1, PosPuerta.Y).Blocked = 1
        Call Bloquear(ToMap, UserIndex, 185, 185, PosPuerta.X - 1, PosPuerta.Y, 1)
        Call Bloquear(ToMap, UserIndex, 185, 185, PosPuerta.X + 1, PosPuerta.Y, 1)

'desbloquea ettin
        MapData(185, PosEttin.X + 1, PosEttin.Y - 1).Blocked = 0
        MapData(185, PosEttin.X - 1, PosEttin.Y - 1).Blocked = 0
        Call Bloquear(ToMap, UserIndex, 185, 185, PosEttin.X - 1, PosEttin.Y - 1, 0)
        Call Bloquear(ToMap, UserIndex, 185, 185, PosEttin.X + 1, PosEttin.Y - 1, 0)
        Call SendData(ToAll, 0, 0, "V9")
    End If


'---------------------------------------------


    If MinPc.Pos.Map = mapa_castillo1 And MinPc.NPCtype = 33 And UserList(UserIndex).GuildInfo.GuildName <> "" Then
        castillo1 = UserList(UserIndex).GuildInfo.GuildName
'nati: NUEVO SISTEMA DE PUNTOS AL CONQUISTAR
    'nati: Esto Obtiene el dueño y numero antes de conquistar
DueñoC = castillo1
TotalClanes = Guilds.Count
For NumGuildD = 1 To TotalClanes
RevisoGuildD = Guilds(NumGuildD).GuildName
If RevisoGuildD = DueñoC Then
Exit For
End If
Next
'nati: aquí cargo el clan que conquista
Conquistador = UserList(UserIndex).GuildInfo.GuildName
For NumGuild = 1 To TotalClanes
RevisoGuild = Guilds(NumGuild).GuildName
If Conquistador = RevisoGuild Then
Exit For
End If
Next
miembros = Guilds(NumGuild).Members.Count
PorcentajeC = 60
variablepuntos = 1
For X = 1 To miembros
PorcentajeC = CInt(PorcentajeC) - variablepuntos
Next X
puntosX = 10 + CInt(PorcentajeC) / 2
PuntosGuild = Guilds(NumGuild).Reputation
SumaPuntosC = PuntosGuild + (puntosX / 2)

Guilds(NumGuild).Reputation = Round(Guilds(NumGuild).Reputation + puntosX)
Guilds(NumGuildD).Reputation = Round(Guilds(NumGuildD).Reputation - puntosX)
If Guilds(NumGuildD).Reputation < 0 Then Guilds(NumGuildD).Reputation = 0
'Call WriteVar(App.Path & "\Guilds\GuildsInfo.inf", "GUILD" & NumGuild, "Rep", PuntosGuild + puntosX)
'nati: NUEVO SISTEMA DE PUNTOS AL CONQUISTAR
'pluto:2.4
        'UserList(UserIndex).Stats.PClan = UserList(UserIndex).Stats.PClan + puntosX
        date1 = Date
        hora1 = Time
        'Call BDDConquistanCastillo("norte", castillo1)
        Call SendData(ToAll, 0, 0, "|| El CLAN " & UCase$(UserList(UserIndex).GuildInfo.GuildName) & " HA CONQUISTADO EL CASTILLO NORTE" & "´" & FontTypeNames.FONTTYPE_talk)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "castillo1", UserList(UserIndex).GuildInfo.GuildName)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "date1", Date)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "hora1", Time)
        castillo1 = UserList(UserIndex).GuildInfo.GuildName
        
        'pluto:6.9-------------------------------------------------
        'Select Case UCase$(UserList(UserIndex).GuildInfo.GuildName)
        'Case "PT AMO"
        'Call SendData(ToAll, 0, 0, "TW" & "ptamo")
        'Case "BLACKLIST"
        'Call SendData(ToAll, 0, 0, "TW" & "black")
        'Case "UNDEAD"
        'Call SendData(ToAll, 0, 0, "TW" & "undead")
        'Case Else
        Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
        'End Select
       '----------------------------------------------------------
        MinutosCastilloNorte = 0
    End If
 
    If MinPc.Pos.Map = mapa_castillo2 And MinPc.NPCtype = 33 And UserList(UserIndex).GuildInfo.GuildName <> "" Then
        castillo2 = UserList(UserIndex).GuildInfo.GuildName
'nati: NUEVO SISTEMA DE PUNTOS AL CONQUISTAR
    'nati: Esto Obtiene el dueño y numero antes de conquistar
DueñoC = castillo2
TotalClanes = Guilds.Count
For NumGuildD = 1 To TotalClanes
RevisoGuildD = Guilds(NumGuildD).GuildName
If RevisoGuildD = DueñoC Then
Exit For
End If
Next
'nati: aquí cargo el clan que conquista
Conquistador = UserList(UserIndex).GuildInfo.GuildName
For NumGuild = 1 To TotalClanes
RevisoGuild = Guilds(NumGuild).GuildName
If Conquistador = RevisoGuild Then
Exit For
End If
Next
miembros = Guilds(NumGuild).Members.Count
PorcentajeC = 60
variablepuntos = 1
For X = 1 To miembros
PorcentajeC = CInt(PorcentajeC) - variablepuntos
Next X
puntosX = 10 + CInt(PorcentajeC) / 2
PuntosGuild = Guilds(NumGuild).Reputation
SumaPuntosC = PuntosGuild + (puntosX / 2)

Guilds(NumGuild).Reputation = Round(Guilds(NumGuild).Reputation + puntosX)
Guilds(NumGuildD).Reputation = Round(Guilds(NumGuildD).Reputation - puntosX)
If Guilds(NumGuildD).Reputation < 0 Then Guilds(NumGuildD).Reputation = 0
'Call WriteVar(App.Path & "\Guilds\GuildsInfo.inf", "GUILD" & NumGuild, "Rep", PuntosGuild + puntosX)
'nati: NUEVO SISTEMA DE PUNTOS AL CONQUISTAR
'pluto:2.4
        'UserList(UserIndex).Stats.PClan = UserList(UserIndex).Stats.PClan + puntosX
        date2 = Date
        hora2 = Time
        'Call BDDConquistanCastillo("sur", castillo2)
        Call SendData(ToAll, 0, 0, "|| El CLAN " & UCase$(UserList(UserIndex).GuildInfo.GuildName) & " HA CONQUISTADO EL CASTILLO SUR" & "´" & FontTypeNames.FONTTYPE_talk)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "castillo2", UserList(UserIndex).GuildInfo.GuildName)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "date2", Date)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "hora2", Time)
        castillo2 = UserList(UserIndex).GuildInfo.GuildName
        'pluto:6.9-------------------------------------------------
        'Select Case UCase$(UserList(UserIndex).GuildInfo.GuildName)
        'Case "PT AMO"
        'Call SendData(ToAll, 0, 0, "TW" & "ptamo")
        'Case "BLACKLIST"
        'Call SendData(ToAll, 0, 0, "TW" & "black")
        'Case "UNDEAD"
        'Call SendData(ToAll, 0, 0, "TW" & "undead")
        'Case Else
        Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
        'End Select
       '----------------------------------------------------------
        MinutosCastilloSur = 0
    End If
 
    If MinPc.Pos.Map = mapa_castillo3 And MinPc.NPCtype = 33 And UserList(UserIndex).GuildInfo.GuildName <> "" Then
        castillo3 = UserList(UserIndex).GuildInfo.GuildName
'nati: NUEVO SISTEMA DE PUNTOS AL CONQUISTAR
    'nati: Esto Obtiene el dueño y numero antes de conquistar
DueñoC = castillo3
TotalClanes = Guilds.Count
For NumGuildD = 1 To TotalClanes
RevisoGuildD = Guilds(NumGuildD).GuildName
If RevisoGuildD = DueñoC Then
Exit For
End If
Next
'nati: aquí cargo el clan que conquista
Conquistador = UserList(UserIndex).GuildInfo.GuildName
For NumGuild = 1 To TotalClanes
RevisoGuild = Guilds(NumGuild).GuildName
If Conquistador = RevisoGuild Then
Exit For
End If
Next
miembros = Guilds(NumGuild).Members.Count
PorcentajeC = 60
variablepuntos = 1
For X = 1 To miembros
PorcentajeC = CInt(PorcentajeC) - variablepuntos
Next X
puntosX = 10 + CInt(PorcentajeC) / 2
PuntosGuild = Guilds(NumGuild).Reputation
SumaPuntosC = PuntosGuild + (puntosX / 2)

Guilds(NumGuild).Reputation = Round(Guilds(NumGuild).Reputation + puntosX)
Guilds(NumGuildD).Reputation = Round(Guilds(NumGuildD).Reputation - puntosX)
If Guilds(NumGuildD).Reputation < 0 Then Guilds(NumGuildD).Reputation = 0
'Call WriteVar(App.Path & "\Guilds\GuildsInfo.inf", "GUILD" & NumGuild, "Rep", PuntosGuild + puntosX)
'nati: NUEVO SISTEMA DE PUNTOS AL CONQUISTAR
'pluto:2.4
        'UserList(UserIndex).Stats.PClan = UserList(UserIndex).Stats.PClan + puntosX
        date3 = Date
        hora3 = Time
        'Call BDDConquistanCastillo("este", castillo3)
        Call SendData(ToAll, 0, 0, "|| El CLAN " & UCase$(UserList(UserIndex).GuildInfo.GuildName) & " HA CONQUISTADO EL CASTILLO ESTE" & "´" & FontTypeNames.FONTTYPE_talk)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "castillo3", UserList(UserIndex).GuildInfo.GuildName)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "date3", Date)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "hora3", Time)
        castillo3 = UserList(UserIndex).GuildInfo.GuildName
         'pluto:6.9-------------------------------------------------
        'Select Case UCase$(UserList(UserIndex).GuildInfo.GuildName)
        'Case "PT AMO"
        'Call SendData(ToAll, 0, 0, "TW" & "ptamo")
        'Case "BLACKLIST"
        'Call SendData(ToAll, 0, 0, "TW" & "black")
        'Case "UNDEAD"
        'Call SendData(ToAll, 0, 0, "TW" & "undead")
        'Case Else
        Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
        'End Select
       '----------------------------------------------------------
        MinutosCastilloEste = 0
    End If
 
    If MinPc.Pos.Map = mapa_castillo4 And MinPc.NPCtype = 33 And UserList(UserIndex).GuildInfo.GuildName <> "" Then
        castillo4 = UserList(UserIndex).GuildInfo.GuildName
'nati: NUEVO SISTEMA DE PUNTOS AL CONQUISTAR
    'nati: Esto Obtiene el dueño y numero antes de conquistar
DueñoC = castillo4
TotalClanes = Guilds.Count
For NumGuildD = 1 To TotalClanes
RevisoGuildD = Guilds(NumGuildD).GuildName
If RevisoGuildD = DueñoC Then
Exit For
End If
Next
'nati: aquí cargo el clan que conquista
Conquistador = UserList(UserIndex).GuildInfo.GuildName
For NumGuild = 1 To TotalClanes
RevisoGuild = Guilds(NumGuild).GuildName
If Conquistador = RevisoGuild Then
Exit For
End If
Next
miembros = Guilds(NumGuild).Members.Count
PorcentajeC = 60
variablepuntos = 1
For X = 1 To miembros
PorcentajeC = CInt(PorcentajeC) - variablepuntos
Next X
puntosX = 10 + CInt(PorcentajeC) / 2
PuntosGuild = Guilds(NumGuild).Reputation
SumaPuntosC = PuntosGuild + (puntosX / 2)

Guilds(NumGuild).Reputation = Round(Guilds(NumGuild).Reputation + puntosX)
Guilds(NumGuildD).Reputation = Round(Guilds(NumGuildD).Reputation - puntosX)
If Guilds(NumGuildD).Reputation < 0 Then Guilds(NumGuildD).Reputation = 0
'Call WriteVar(App.Path & "\Guilds\GuildsInfo.inf", "GUILD" & NumGuild, "Rep", PuntosGuild + puntosX)
'nati: NUEVO SISTEMA DE PUNTOS AL CONQUISTAR
'pluto:2.4
        'UserList(UserIndex).Stats.PClan = UserList(UserIndex).Stats.PClan + puntosX
        date4 = Date
        hora4 = Time
        'Call BDDConquistanCastillo("oeste", castillo4)
        Call SendData(ToAll, 0, 0, "|| El CLAN " & UCase$(UserList(UserIndex).GuildInfo.GuildName) & " HA CONQUISTADO EL CASTILLO OESTE" & "´" & FontTypeNames.FONTTYPE_talk)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "castillo4", UserList(UserIndex).GuildInfo.GuildName)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "date4", Date)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "hora4", Time)
        castillo4 = UserList(UserIndex).GuildInfo.GuildName
         'pluto:6.9-------------------------------------------------
        'Select Case UCase$(UserList(UserIndex).GuildInfo.GuildName)
        'Case "PT AMO"
        'Call SendData(ToAll, 0, 0, "TW" & "ptamo")
        'Case "BLACKLIST"
        'Call SendData(ToAll, 0, 0, "TW" & "black")
        'Case "UNDEAD"
        'Call SendData(ToAll, 0, 0, "TW" & "undead")
        'Case Else
        Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
        'End Select
       '----------------------------------------------------------
        MinutosCastilloOeste = 0
    End If
   
    If UserIndex > 0 Then ' Lo mato un usuario?
        If MinPc.flags.Snd3 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MinPc.flags.Snd3)
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        
        'El user que lo mato tiene mascotas?
        If UserList(UserIndex).NroMacotas > 0 Then
            Dim t As Integer
            For t = 1 To MAXMASCOTAS
                If UserList(UserIndex).MascotasIndex(t) > 0 Then
                    If Npclist(UserList(UserIndex).MascotasIndex(t)).TargetNpc = NpcIndex Then
                        Call FollowAmo(UserList(UserIndex).MascotasIndex(t))
                    End If
                End If
            Next t
        End If
'pluto:6.8-------------------
If MinPc.numero = BichoDelDia Then MinPc.GiveEXP = MinPc.GiveEXP + Int(MinPc.GiveEXP / 2)
'----------------------------

'pluto:doble exp en casas encantadas
        If UserList(UserIndex).Pos.Map = 171 Or UserList(UserIndex).Pos.Map = 177 Then MinPc.GiveEXP = MinPc.GiveEXP * 2
'pluto:6.0A
UserList(UserIndex).Stats.Fama = UserList(UserIndex).Stats.Fama + Int(MinPc.GiveEXP / 200000)

'pluto:6.0a restamos exp que se lleva mascota /1000
        If UserList(UserIndex).flags.Montura > 0 And UserList(UserIndex).flags.party = False Then MinPc.GiveEXP = Int(MinPc.GiveEXP / 2)
'pluto:6.9 evento 2
        If DobleExp > 0 Then MinPc.GiveEXP = MinPc.GiveEXP * 2
        If EventoDia = 2 Then MinPc.GiveEXP = MinPc.GiveEXP * 2

'pluto:2.17 server secundario
        If DifOro > 1 Then
        'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + (MiNPC.GiveGLD * DifOro)
Call AddtoVar(UserList(UserIndex).Stats.GLD, (MinPc.GiveGLD * DifOro), MAXORO)

        Call SendUserStatsOro(UserIndex)
        End If
'-----------
  'secundario
'If ServerPrimario = 2 Then MinPc.GiveEXP = MinPc.GiveEXP * expdemas
       
        'METE LA EXP AL PJ QUE LO MATA
'[Tite] Party: Repartimos exp entre todos si es party, si no solo al pj
        If UserList(UserIndex).flags.party = False Then
                'pluto:6.3----------------------------
               If ServerPrimario = 1 Then
                    'Select Case UserList(UserIndex).Stats.ELV
                        'Case Is < 50
                        'Expdemas = 30
                        'Case 50 To 60
                        'Expdemas = 30
                        'Case Is > 60
                        'Expdemas = 30
                        Expdemas = 1 ' 1
                    'End Select
                    If UserList(UserIndex).Remort > 0 Then Expdemas = 1
                Else 'secundario
                 'pluto:6.5
                        'Select Case UserList(UserIndex).Stats.ELV
                        'Case Is < 30
                        'Expdemas = 10
                        'Case 30 To 40
                        'Expdemas = 5
                        'Case 41 To 50
                        'Expdemas = 3
                        'Case Is > 50
                        'Expdemas = 2
                        Expdemas = 1 ' 1
                        'End Select
                     If UserList(UserIndex).Remort > 0 Then Expdemas = 1

                End If
            '---------------------------------------
            
            Call AddtoVar(UserList(UserIndex).Stats.exp, MinPc.GiveEXP * Expdemas, MAXEXP)
        Else 'en party
            MinPc.GiveEXP = MinPc.GiveEXP * (1 + (0.05 * partylist(UserList(UserIndex).flags.partyNum).numMiembros))
            Call PartyReparteExp(MinPc, UserIndex)
        End If
          
    
'[\Tite]
'pluto:6.2

        Dim aa As Long
        If UserList(UserIndex).flags.Montura > 0 Then ' en x30 hay que darle a /10000 en normal /1000
        aa = Int(MinPc.GiveEXP * Expdemas / 1000) 'meto x10 a mascotas antes /1000 ahora /100
        Else
        aa = 0
        End If
'[Tite]Party: Mandamos mensaje de la exp recibida :)
If UserList(UserIndex).flags.party = False Then
            Call SendData(ToIndex, UserIndex, 0, "V6" & MinPc.GiveEXP * Expdemas & "," & aa)

 

'El user tiene montura (hay que repartir exp con ella)
        If UserList(UserIndex).flags.Montura > 0 And UserList(UserIndex).flags.ClaseMontura > 0 Then
'añade topelevel
            If PMascotas(UserList(UserIndex).flags.ClaseMontura).TopeLevel > UserList(UserIndex).Montura.Nivel(UserList(UserIndex).flags.ClaseMontura) Then
                'Comprobamos que no este bugueada
                If UserList(UserIndex).Montura.Elu(UserList(UserIndex).flags.ClaseMontura) = 0 Then
                    Call SendData(ToGM, 0, 0, "|| Matanpc Mascota Bugueada: " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_COMERCIO)
                    Call LogMascotas("BUG MataNpcMASCOTA Serie: " & UserList(UserIndex).Serie & " IP: " & UserList(UserIndex).ip & " Nom: " & UserList(UserIndex).Name)
                End If
'----------------
                'Le metemos la exp a la montura
                Call AddtoVar(UserList(UserIndex).Montura.exp(UserList(UserIndex).flags.ClaseMontura), Int(MinPc.GiveEXP * Expdemas / 1000), MAXEXP) 'meto x10 a mascotas antes /1000 ahora /100
                Call CheckMonturaLevel(UserIndex)
            End If
        End If 'topelevel
End If 'party
'[\Tite]
    
    Call AddtoVar(UserList(UserIndex).Stats.NPCsMuertos, 1, 32000)
'pluto:2.15
        Call SendUserMuertos(UserIndex)
        If MinPc.Stats.Alineacion = 0 Then
'pluto:2.11 --> Todos no solo guardias(no activado)
            If MinPc.numero = Guardias Then
                Call VolverCriminal(UserIndex)
            End If
            If Not EsDios(UserList(UserIndex).Name) Then Call AddtoVar(UserList(UserIndex).Reputacion.AsesinoRep, vlASESINO, MAXREP)
        ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
        ElseIf MinPc.Stats.Alineacion = 1 Then
            Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR, MAXREP)
        ElseIf MinPc.Stats.Alineacion = 2 Then
            Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, vlASESINO / 2, MAXREP)
        
        ElseIf MinPc.Stats.Alineacion = 4 Then
            Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR, MAXREP)
        End If
    
'pluto:hoy
        'If UserList(UserIndex).Mision.Enemigo = MinPc.Numero And UserList(UserIndex).Mision.estado = 1 Then
         '   UserList(UserIndex).Mision.estado = 0
          '  Call SendData(ToIndex, UserIndex, 0, "!!Quest Número " & UserList(UserIndex).Mision.Numero & " : " & " Muy bién, has cumplido una misión!!")
  'pluto:6.0A
        'UserList(UserIndex).Stats.Fama = UserList(UserIndex).Stats.Fama + 5
'pluto:2-3-04
         '   Call SendData(ToIndex, UserIndex, 0, "|| Has ganado " & val(Int(UserList(UserIndex).Mision.Numero / 10) + 1) & " DragPuntos." & "´" & FontTypeNames.FONTTYPE_info)
          '  UserList(UserIndex).Stats.Puntos = UserList(UserIndex).Stats.Puntos + Int(UserList(UserIndex).Mision.Numero / 10) + 1
'pluto:2.19----------------------------------------------
           ' Dim Mise As Long
          
            'Mise = 500& * UserList(UserIndex).Mision.Numero
            'Call SendData(ToIndex, UserIndex, 0, "|| Has ganado " & " Puntos de Experiencia." & "´" & FontTypeNames.FONTTYPE_info)
            'UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + (UserList(UserIndex).Mision.Numero * 500&)
            'SendUserStatsEXP (UserIndex)
            'CheckUserLevel (UserIndex)
'----------------------------------------------
        'End If
        'Controla el nivel del usuario
        Call CheckUserLevel(UserIndex)
    End If ' Userindex > 0
    
    If MinPc.MaestroUser = 0 Then
        'Tiramos el oro
        Call NPCTirarOro(MinPc)
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MinPc)
    End If
    'ReSpawn o no
'pluto:6.0A
    'If MiNPC.MaestroUser = 0 And MiNPC.NPCtype <> 60 And MiNPC.NPCtype <> 33 And MiNPC.Name <> "Momia Faraón" And MiNPC.NPCtype <> 61 And MiNPC.NPCtype <> 88 Then Call ReSpawnNpc(MiNPC)
   If MinPc.MaestroUser = 0 Then Call ReSpawnNpc(MinPc)

   'Call ReSpawnNpc(MiNPC)
End If 'maestrouser=0
'pluto:2.4
If MinPc.NPCtype = 60 And MinPc.MaestroUser > 0 Then
    Call QuitarObjetos(887 + UserList(MinPc.MaestroUser).flags.ClaseMontura, 1, MinPc.MaestroUser)
    Call LogMascotas("Muere objeto " & 887 + UserList(MinPc.MaestroUser).flags.ClaseMontura & " de " & UserList(MinPc.MaestroUser).Name)

    UserList(MinPc.MaestroUser).flags.Montura = 0
    UserList(MinPc.MaestroUser).flags.ClaseMontura = 0

End If
'pluto:2.4
If MinPc.NPCtype = 60 And MinPc.MaestroUser = 0 Then
    Dim CabalgaPos As WorldPos
    Dim ini As Integer
    Dim mapita As Integer
    'evitamos respawn otro mapa del jabato
    If MinPc.flags.Domable = 506 Then GoTo fin
a:
    mapita = RandomNumber(1, 277)
    CabalgaPos.X = RandomNumber(15, 80)
    CabalgaPos.Y = RandomNumber(15, 80)
    CabalgaPos.Map = mapita
    If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a
    ini = SpawnNpc(MinPc.numero, CabalgaPos, False, True)
    If ini = MAXNPCS Then GoTo a
    Call WriteVar(IniPath & "cabalgar.txt", MinPc.Name, "Mapa", val(mapita))
End If
'---fin pluto:2.4-----
fin:

'pluto:6.8 añadido evento 2
If MinPc.NPCtype = 88 And MinPc.Pos.Map = 250 And UserIndex > 0 Then
    If UserList(UserIndex).flags.Privilegios = 0 Then
        HeroeExp = UserList(UserIndex).Name
        DobleExp = 30
        Call SendData(ToAll, 0, 0, "!! La energia del Caballero Helado derrotado por " & UserList(UserIndex).Name & " recorre el mundo Aodrag otorgando a todos los habitantes poderes especiales durante unos minutos.")
        MsgEntra = "La energía del Caballero Helado derrotado por " & UserList(UserIndex).Name & " recorre el mundo Aodrag otorgando a todos los habitantes poderes especiales durante " & DobleExp & " minutos."
        Call LogCasino("Jugador:" & UserList(UserIndex).Name & " Caballero x2 " & "Ip: " & UserList(UserIndex).ip)
        Caballero = False
    End If
End If

If MinPc.NPCtype = 145 And MinPc.Pos.Map = 279 And UserIndex > 0 Then
    If UserList(UserIndex).flags.Privilegios = 0 Then
        HeroeExp = UserList(UserIndex).Name
        DobleExp = 60
        Call SendData(ToAll, 0, 0, "!! La Oscuridad del Caballero de la Muerte derrotado por " & UserList(UserIndex).Name & " recorre el mundo otorgando a todos los habitantes poderes especiales durante unos minutos.")
        MsgEntra = "La Oscuridad del Caballero de la Muerte derrotado por " & UserList(UserIndex).Name & " recorre el mundo otorgando a todos los habitantes poderes especiales durante " & DobleExp & " minutos."
        Call LogCasino("Jugador:" & UserList(UserIndex).Name & " Caballero x2 " & "Ip: " & UserList(UserIndex).ip)
        Caballero = False
    End If
End If
'If MiNPC.NPCtype = 79 Then
    'Call ConquistarCiudad(MiNPC.Pos.Map, UserIndex)
'End If

'End If
'----------pluto:6.5 matamos un RAID------------------
  'Dim nn As Byte
   If MinPc.Raid > 0 Then
   RaidVivos(MinPc.numero - 699).Activo = 0
   RaidVivos(MinPc.numero - 699).MiniRaids = 0
   'For nn = 2 To 6
   'If RandomNumber(1, 100) > 70 Then Call SpawnNpc(MiNPC.Numero + 7, MiNPC.Pos, True, False)
   'Next
   Call LogCasino("Jugador:" & UserList(UserIndex).Name & " mata " & MinPc.Name & "Ip: " & UserList(UserIndex).ip)
   End If
'-------------------------
   'pluto:6.0A saca gollum pequeños------
   If MinPc.numero = 594 Then
   'Dim n As Byte

   'n = RandomNumber(1, 5)
   For nn = 1 To 6
   If RandomNumber(1, 100) > 70 Then Call SpawnNpc(697, MinPc.Pos, True, False)
   Next
   End If
   '---------------------------
   
   'pluto:6.0A minotauro
If MinPc.numero = 692 Then
   If Minotauro = UserList(UserIndex).Name Then
   'aqui la recompensa
   UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + 200
   UserList(UserIndex).Stats.Puntos = UserList(UserIndex).Stats.Puntos + 1000
   UserList(UserIndex).Stats.Fama = UserList(UserIndex).Stats.Fama + 50000
    Call SendData(ToIndex, UserIndex, 0, "|| Has ganado 200 Puntos Habilidades Libres para Asignar" & "´" & FontTypeNames.FONTTYPE_info)
    Call SendData(ToIndex, UserIndex, 0, "|| Has ganado 1000 DraGPuntos." & "´" & FontTypeNames.FONTTYPE_info)
    Call SendData(ToIndex, UserIndex, 0, "|| Has ganado mucha Popularidad." & "´" & FontTypeNames.FONTTYPE_info)
    Call SendData(ToAll, UserIndex, 0, "|| El Minotauro ha sido asesinado por " & UserList(UserIndex).Name & " que fué el Personaje que lo liberó, los Dioses le han otorgado los Poderes del Minotauro!!" & "´" & FontTypeNames.FONTTYPE_PARTY)
    UserList(UserIndex).flags.Minotauro = 1
    Call SendUserStatsFama(UserIndex)
    Else
    Call SendData(ToAll, UserIndex, 0, "|| El Minotauro ha sido asesinado por " & UserList(UserIndex).Name & ", este personaje no lo liberó y los Dioses no le conceden los Poderes del Minotauro." & "´" & FontTypeNames.FONTTYPE_PARTY)
    End If
   
   Minotauro = ""
   EstadoMinotauro = 2
   MinutosMinotauro = 0
End If
   
   'PLUTO:6.8
    If EventoDia = 3 And MinPc.Stats.MaxHP > 999 And MinPc.NPCtype <> NPCTYPE_GUARDIAS And MinPc.NPCtype <> NPCTYPE_GUARDIAS2 Then
    Dim PROBA As Byte
    PROBA = RandomNumber(1, 15)
 
If PROBA = 8 And UserList(UserIndex).Pos.Map <> 165 Then
Call SpawnNpc(718, UserList(UserIndex).Pos, True, False)
End If
   End If
   
    'Delzak sistema premios (17-8-10)
       'If MinPc.Premio = 1 Then Call PremioMataNPC(MinPc.numero, UserIndex)
 'pluto:7.0-------------------------
  If MinPc.LogroTipo > 0 Then
 ' UserList(UserIndex).Stats.PremioNPC(MinPc.LogroTipo) = UserList(UserIndex).Stats.PremioNPC(MinPc.LogroTipo) + 1
 Call PremioMataNPC(MinPc.LogroTipo, UserIndex)
  End If
   '-------------------
   Exit Sub

errhandler:
    Call LogError("Error en MuereNpc: " & MinPc.Name & " matado por " & UserList(UserIndex).Name)
    
End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
'Clear the npc's flags
On Error GoTo fallo
Npclist(NpcIndex).flags.PoderEspecial1 = 0
Npclist(NpcIndex).flags.PoderEspecial2 = 0
Npclist(NpcIndex).flags.PoderEspecial3 = 0
Npclist(NpcIndex).flags.PoderEspecial4 = 0
Npclist(NpcIndex).flags.PoderEspecial5 = 0
Npclist(NpcIndex).flags.PoderEspecial6 = 0

Npclist(NpcIndex).flags.AfectaParalisis = 0
Npclist(NpcIndex).flags.Magiainvisible = 0
Npclist(NpcIndex).flags.AguaValida = 0
Npclist(NpcIndex).flags.AttackedBy = ""
Npclist(NpcIndex).flags.Attacking = 0
Npclist(NpcIndex).flags.BackUp = 0
Npclist(NpcIndex).flags.Bendicion = 0
Npclist(NpcIndex).flags.Domable = 0
Npclist(NpcIndex).flags.Envenenado = 0
Npclist(NpcIndex).flags.Faccion = 0
Npclist(NpcIndex).flags.Follow = False
Npclist(NpcIndex).flags.LanzaSpells = 0
Npclist(NpcIndex).flags.GolpeExacto = 0
Npclist(NpcIndex).flags.Invisible = 0
Npclist(NpcIndex).flags.Maldicion = 0
Npclist(NpcIndex).flags.OldHostil = 0
Npclist(NpcIndex).flags.OldMovement = 0
Npclist(NpcIndex).flags.Paralizado = 0
Npclist(NpcIndex).flags.Respawn = 0
Npclist(NpcIndex).flags.RespawnOrigPos = 0
Npclist(NpcIndex).flags.Snd1 = 0
Npclist(NpcIndex).flags.Snd2 = 0
Npclist(NpcIndex).flags.Snd3 = 0
Npclist(NpcIndex).flags.Snd4 = 0
Npclist(NpcIndex).flags.TierraInvalida = 0
Npclist(NpcIndex).flags.UseAINow = False
Exit Sub
fallo:
Call LogError("resetnpcflags " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)
On Error GoTo fallo
Npclist(NpcIndex).Contadores.Paralisis = 0
Npclist(NpcIndex).Contadores.TiempoExistencia = 0
Exit Sub
fallo:
Call LogError("resetnpccounters " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
On Error GoTo fallo
Npclist(NpcIndex).Char.Body = 0
Npclist(NpcIndex).Char.CascoAnim = 0
Npclist(NpcIndex).Char.CharIndex = 0
Npclist(NpcIndex).Char.FX = 0
Npclist(NpcIndex).Char.Head = 0
Npclist(NpcIndex).Char.Heading = 0

Npclist(NpcIndex).Char.loops = 0
Npclist(NpcIndex).Char.ShieldAnim = 0
Npclist(NpcIndex).Char.WeaponAnim = 0

Exit Sub
fallo:
Call LogError("resetnpcharinfo " & Err.number & " D: " & Err.Description)

End Sub


Sub ResetNpcCriatures(ByVal NpcIndex As Integer)

On Error GoTo fallo
Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroCriaturas
    Npclist(NpcIndex).Criaturas(j).NpcIndex = 0
    Npclist(NpcIndex).Criaturas(j).NpcName = ""
Next j

Npclist(NpcIndex).NroCriaturas = 0
Exit Sub
fallo:
Call LogError("resetnpcriatures " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
On Error GoTo fallo
Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroExpresiones: Npclist(NpcIndex).Expresiones(j) = "": Next j

Npclist(NpcIndex).NroExpresiones = 0
Exit Sub
fallo:
Call LogError("resetexpresiones " & Err.number & " D: " & Err.Description)

End Sub


Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
On Error GoTo fallo
Npclist(NpcIndex).Attackable = 0
Npclist(NpcIndex).CanAttack = 0
Npclist(NpcIndex).Comercia = 0
Npclist(NpcIndex).GiveEXP = 0
Npclist(NpcIndex).GiveGLD = 0
Npclist(NpcIndex).Hostile = 0
Npclist(NpcIndex).Inflacion = 0
Npclist(NpcIndex).InvReSpawn = 0
Npclist(NpcIndex).Level = 0

If Npclist(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)

If Npclist(NpcIndex).MaestroNpc > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc, NpcIndex)

Npclist(NpcIndex).MaestroUser = 0
Npclist(NpcIndex).MaestroNpc = 0

Npclist(NpcIndex).Mascotas = 0
Npclist(NpcIndex).Movement = 0
Npclist(NpcIndex).Name = "NPC SIN INICIAR"
'pluto:2.22------------------------
Npclist(NpcIndex).flags.NPCActive = False
'---------------------------------
Npclist(NpcIndex).NPCtype = 0
Npclist(NpcIndex).numero = 0
Npclist(NpcIndex).Anima = 0
'pluto:6.0A
Npclist(NpcIndex).Arquero = 0
Npclist(NpcIndex).Orig.Map = 0
Npclist(NpcIndex).Orig.X = 0
Npclist(NpcIndex).Orig.Y = 0
Npclist(NpcIndex).PoderAtaque = 0
Npclist(NpcIndex).PoderEvasion = 0
Npclist(NpcIndex).Pos.Map = 0
Npclist(NpcIndex).Pos.X = 0
Npclist(NpcIndex).Pos.Y = 0
Npclist(NpcIndex).SkillDomar = 0
Npclist(NpcIndex).Target = 0
Npclist(NpcIndex).TargetNpc = 0
Npclist(NpcIndex).TipoItems = 0
Npclist(NpcIndex).veneno = 0
Npclist(NpcIndex).Desc = ""

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroSpells
    Npclist(NpcIndex).Spells(j) = 0
Next j

Call ResetNpcCharInfo(NpcIndex)
Call ResetNpcCriatures(NpcIndex)
Call ResetExpresiones(NpcIndex)
Exit Sub
fallo:
Call LogError("resetnpcmaininfo " & Err.number & " D: " & Err.Description)

End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

On Error GoTo errhandler

Npclist(NpcIndex).flags.NPCActive = False


If InMapBounds(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then
    Call EraseNPCChar(ToMap, 0, Npclist(NpcIndex).Pos.Map, NpcIndex)
End If

'Nos aseguramos de que el inventario sea removido...
'asi los lobos no volveran a tirar armaduras ;))
'Debug.Print Npclist(NpcIndex).Name
Call ResetNpcInv(NpcIndex)
Call ResetNpcFlags(NpcIndex)
Call ResetNpcCounters(NpcIndex)

Call ResetNpcMainInfo(NpcIndex)

If NpcIndex = LastNPC Then
    Do Until Npclist(LastNPC).flags.NPCActive
        LastNPC = LastNPC - 1
        If LastNPC < 1 Then Exit Do
    Loop
End If
    
  
If NumNPCs <> 0 Then
    NumNPCs = NumNPCs - 1
End If

Exit Sub

errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(Pos As WorldPos, agua As Byte) As Boolean

On Error GoTo fallo
If LegalPosNPC(Pos.Map, Pos.X, Pos.Y, agua) Then
    TestSpawnTrigger = _
    MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 3 And _
    MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 2 And _
    MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 1
End If
Exit Function
fallo:
Call LogError("testspawntrigger " & Err.number & " D: " & Err.Description)

End Function

Sub CrearNPC(NroNPC As Integer, Mapa As Integer, OrigPos As WorldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC
On Error GoTo fallo
Dim Pos As WorldPos
Dim Newpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long


Dim Map As Integer
Dim X As Integer
Dim Y As Integer

'pluto:6.4
'If Mapa = 0 Then Exit Sub

nIndex = OpenNPC(NroNPC) 'Conseguimos un indice

If nIndex > MAXNPCS Then Exit Sub

'Necesita ser respawned en un lugar especifico
If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
    
    Map = OrigPos.Map
    X = OrigPos.X
    Y = OrigPos.Y
    Npclist(nIndex).Orig = OrigPos
    Npclist(nIndex).Pos = OrigPos
   
Else
    
    Pos.Map = Mapa 'mapa
    
    Do While Not PosicionValida
    
        Randomize (Timer) 'pluto:2.18 cambio a 85 +5 el 100+1
        'pluto:6.0A--------------
        If Pos.Map = 139 Or Pos.Map = 48 Or Pos.Map = 110 Then
        Pos.X = CInt(Rnd * 15 + 40)
        Pos.Y = CInt(Rnd * 15 + 40)
        Else
        Pos.X = CInt(Rnd * 85 + 5) 'Obtenemos posicion al azar en x
        Pos.Y = CInt(Rnd * 85 + 5) 'Obtenemos posicion al azar en y
        End If
        '--------------------------
          
        Call ClosestLegalPos(Pos, Newpos, Npclist(nIndex).flags.AguaValida) 'Nos devuelve la posicion valida mas cercana
 'pluto:2.18
 If Newpos.X = 0 Or Newpos.Y = 0 Then GoTo debuge
 
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If LegalPosNPC(Newpos.Map, Newpos.X, Newpos.Y, Npclist(nIndex).flags.AguaValida) And _
           Not HayPCarea(Newpos) And TestSpawnTrigger(Newpos, Npclist(nIndex).flags.AguaValida) Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).Pos.Map = Newpos.Map
            Npclist(nIndex).Pos.X = Newpos.X
            Npclist(nIndex).Pos.Y = Newpos.Y
            PosicionValida = True

        Else
            Newpos.X = 0
            Newpos.Y = 0
        
        End If
debuge:
        'for debug
        Iteraciones = Iteraciones + 1
        If Iteraciones > MAXSPAWNATTEMPS Then
                Call QuitarNPC(nIndex)
                Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & Mapa & " NroNpc:" & Npclist(NroNPC).Name)
                Exit Sub
        End If
    Loop
    
    'asignamos las nuevas coordenas
    Map = Newpos.Map
    X = Npclist(nIndex).Pos.X
    Y = Npclist(nIndex).Pos.Y
'pluto:6.0A
'If Npclist(NpcIndex).Pos.Map = 60 Then Npclist(NpcIndex).flags.Respawn = 0

End If

'Crea el NPC
Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)
Exit Sub
fallo:
Call LogError("crearnpc Map:" & Map & " X: " & X & " Y: " & Y & " Name: " & Npclist(nIndex).Name & " Tipo: " & Npclist(NroNPC).Name & "Err: " & Err.number & " D: " & Err.Description)

End Sub

Sub MakeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

On Error GoTo fallo
Dim CharIndex As Integer

If Npclist(NpcIndex).Char.CharIndex = 0 Then
    CharIndex = NextOpenCharIndex
    Npclist(NpcIndex).Char.CharIndex = CharIndex
    CharList(CharIndex) = NpcIndex
End If

MapData(Map, X, Y).NpcIndex = NpcIndex
'pluto:6.0A añado raid

Call SendData(sndRoute, sndIndex, sndMap, "JX" & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Char.CharIndex & "," & X & "," & Y & "," & Npclist(NpcIndex).Raid)

Exit Sub
fallo:
Call LogError("makenpchar " & Err.number & " D: " & Err.Description)

End Sub

Sub ChangeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, NpcIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, Ata As Byte)
On Error GoTo fallo
If NpcIndex > 0 Then
    'pluto:6.0A--------
    If Npclist(NpcIndex).Char.Heading = Heading And Ata = 0 Then Exit Sub
    '-------------------
    Npclist(NpcIndex).Char.Body = Body
    Npclist(NpcIndex).Char.Head = Head
    Npclist(NpcIndex).Char.Heading = Heading
    Call SendData(sndRoute, sndIndex, sndMap, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)
End If
Exit Sub
fallo:
Call LogError("changenpcchar " & Err.number & " D: " & Err.Description)

End Sub

Sub EraseNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ByVal NpcIndex As Integer)
On Error GoTo fallo
If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar < 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

'Actualizamos los cliente
Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "BP" & Npclist(NpcIndex).Char.CharIndex)

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
'NumChars = NumChars - 1
Exit Sub
fallo:
Call LogError("erasenpcchar Npcindex: " & NpcIndex & " Name:" & Npclist(NpcIndex).Name & " " & Err.number & " D: " & Err.Description)


End Sub
Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)

On Error GoTo errh
    Dim n3 As Byte
    Dim nPos As WorldPos
    Dim n As Byte
    Dim n2 As Byte
    Dim n4 As Byte
    'pluto:2.8.0
    If NpcIndex < 1 Then Exit Sub
    'nati:agrego NPCType = 33
    If (Npclist(NpcIndex).NPCtype = 77 Or Npclist(NpcIndex).NPCtype = 78 Or Npclist(NpcIndex).NPCtype = 33) Then Exit Sub
    
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(nHeading, nPos)
  'pluto:2.9.0
If Npclist(NpcIndex).NPCtype = 21 And Npclist(NpcIndex).Pos.Map = 192 Then

If (Npclist(NpcIndex).Pos.X < 48 Or Npclist(NpcIndex).Pos.X > 54) Then
n = RandomNumber(1, 10)
If n = 10 Then Call SendData(ToMap, NpcIndex, Npclist(NpcIndex).Pos.Map, "TW" & 146)
If Npclist(NpcIndex).Pos.X < 39 And nHeading = 4 Then Exit Sub
If Npclist(NpcIndex).Pos.Y < 35 And nHeading = 1 Then Exit Sub
If Npclist(NpcIndex).Pos.X > 63 And nHeading = 2 Then Exit Sub
If Npclist(NpcIndex).Pos.Y > 66 And nHeading = 3 Then Exit Sub
End If

If Npclist(NpcIndex).NPCtype = 21 And Npclist(NpcIndex).Pos.Map = 192 And (Npclist(NpcIndex).Pos.X > 47 And Npclist(NpcIndex).Pos.X < 55) And Vezz = 0 Then
If Npclist(NpcIndex).Pos.Y = 33 Then GolesLocal = GolesLocal + 1: Call SendData2(ToMap, 0, Npclist(NpcIndex).Pos.Map, 92, GolesLocal & "," & GolesVisitante & "," & 1): Vezz = 1: Call SendData(ToMap, NpcIndex, Npclist(NpcIndex).Pos.Map, "TW" & 105)

If Npclist(NpcIndex).Pos.Y = 68 Then GolesVisitante = GolesVisitante + 1: Call SendData2(ToMap, 0, Npclist(NpcIndex).Pos.Map, 92, GolesLocal & "," & GolesVisitante & "," & 1): Vezz = 1: Call SendData(ToMap, NpcIndex, Npclist(NpcIndex).Pos.Map, "TW" & 105)

End If

End If 'map 192 y npctype21
  'pluto:2.14
  If Npclist(NpcIndex).flags.PoderEspecial1 > 0 Then
n2 = RandomNumber(1, 10)
If n2 > 4 Then GoTo ffu8
If Npclist(NpcIndex).Char.Body < 330 Then GoTo ffu3
If n2 = 1 Then Npclist(NpcIndex).Char.Body = 331
If n2 = 2 Then Npclist(NpcIndex).Char.Body = 330
ffu3:
n4 = RandomNumber(1, 100)
If n4 < 98 Then GoTo ffu8
If Npclist(NpcIndex).Char.Body < 330 Then Npclist(NpcIndex).Char.Body = 331: GoTo ffu
  n3 = RandomNumber(1, 100)
If n3 > 20 Then Npclist(NpcIndex).Char.Body = 10
If n3 > 40 Then Npclist(NpcIndex).Char.Body = 13
If n3 > 60 Then Npclist(NpcIndex).Char.Body = 9
If n3 > 80 Then Npclist(NpcIndex).Char.Body = 51

ffu:
Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, nHeading, 1)
Call SendData2(ToMap, 0, nPos.Map, 22, Npclist(NpcIndex).Char.CharIndex & "," & FXWARP & "," & 0)
ffu8:
End If 'especial




ffu2:
 '-------------------------------------
    'Es mascota ????
    If Npclist(NpcIndex).MaestroUser > 0 Then
            
       'pluto:2.4
If Npclist(NpcIndex).NPCtype = 60 Then
Dim User As Integer
User = Npclist(NpcIndex).MaestroUser

Dim tt As Integer
Dim nn As Integer
Dim kk As Integer
tt = RandomNumber(1, 100)
If tt > 20 Then

'pluto:2.17 index npc mascota
If Npclist(NpcIndex).numero < 621 Then
nn = Npclist(NpcIndex).numero - 615
Else
nn = Npclist(NpcIndex).numero - 663
End If

'sube mana
If (nn = 1 Or nn = 2 Or nn = 11 Or nn = 8) And (UserList(User).Stats.MinMAN < UserList(User).Stats.MaxMAN) Then
kk = 55
If UserList(User).Montura.Nivel(nn) > 4 Then kk = 56
End If 'nn=1

'dopa fuerza
If (nn = 4 Or nn = 9 Or nn = 7) And UserList(User).Stats.UserAtributos(1) < 35 Then
kk = 20
If UserList(User).Montura.Nivel(nn) > 4 Then kk = 22
End If 'nn=2

'dopa cele
If (nn = 3 Or nn = 10 Or nn = 12) And UserList(User).Stats.UserAtributos(2) < 35 Then
kk = 18
If UserList(User).Montura.Nivel(nn) > 4 Then kk = 40
End If 'nn=3

'dopa curar
If nn = 5 And UserList(User).Stats.MinHP < UserList(User).Stats.MaxHP Then
kk = 5
If UserList(User).Montura.Nivel(nn) > 4 Then kk = 42
End If 'nn=5


If kk = 0 Then Exit Sub
Call NpcLanzaSpellSobreUser(NpcIndex, User, kk)
Exit Sub
End If ' tt
End If ' type=60
  '------------------fin pluto:2.4-----------
            
            
            ' es una posicion legal
            If LegalPos(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then
            
                If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                'pluto:2.23-----------------------
                Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y & ",0")
                'Call SendToNpcArea(NpcIndex, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y & ",0")
                '-------------------------------
                'Update map and user pos
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
                Npclist(NpcIndex).Pos = nPos
                Npclist(NpcIndex).Char.Heading = nHeading
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            'pluto:2.23------------------------------
            'Call CheckUpdateNeededNpc(NpcIndex, nHeading)
            '-----------------------------------------
            
            End If
          'pluto:6.0A
          Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, nHeading, 0)

             
    Else ' No es mascota
            ' Controlamos que la posicion sea legal, los npc que
            ' no son mascotas tienen mas restricciones de movimiento.
            If LegalPosNPC(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida) Then
                
                If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
            'pluto:2.23-----------------------
                Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y & ",0")
                'Call SendToNpcArea(NpcIndex, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y & ",0")
            '---------------------------------
           
                              'Update map and user pos
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
                Npclist(NpcIndex).Pos = nPos
                Npclist(NpcIndex).Char.Heading = nHeading
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            'pluto:2.23------------------------------
            'Call CheckUpdateNeededNpc(NpcIndex, nHeading)
            '-----------------------------------------
            Else
                If Npclist(NpcIndex).Movement = NPC_PATHFINDING Then
                    'Someone has blocked the npc's way, we must to seek a new path!
                    Npclist(NpcIndex).PFINFO.PathLenght = 0
                End If
            
            End If
    End If
'pluto:santa claus

If Npclist(NpcIndex).NPCtype = 13 Then
Dim ale As Integer
ale = RandomNumber(1, 250)
If ale = 100 Then
Call SendData(ToMap, NpcIndex, Npclist(NpcIndex).Pos.Map, "TW" & 119)
Call NPC_TIRAR_ITEMS(Npclist(NpcIndex))
End If
End If
Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)


End Sub
Sub RespawnRaids(n As Byte)
Dim Raid As WorldPos
Dim a As Byte
Dim na As String
Dim ini As Integer
a = RandomNumber(1, 100)

If a > 20 Then Exit Sub
Select Case n
Case 1
Raid.X = 38
Raid.Y = 66
Raid.Map = 9
na = "Bursol"

Case 2
Raid.X = 51
Raid.Y = 32
Raid.Map = 14
na = "Faren"

Case 3
Raid.X = 46
Raid.Y = 79
Raid.Map = 193
na = "Mirgan"

Case 4
Raid.X = 76
Raid.Y = 66
Raid.Map = 160
na = "Tirgan"

Case 5
Raid.X = 27
Raid.Y = 37
Raid.Map = 76
na = "Colossus"

Case 6
Raid.X = 47
Raid.Y = 24
Raid.Map = 188
na = "Lostel"



End Select

ini = SpawnNpc(699 + n, Raid, False, True)

If ini <> MAXNPCS Then
Call SendData(ToAll, 0, 0, "||Reaparece el Monster DraG " & na & "´" & FontTypeNames.FONTTYPE_PARTY)
RaidVivos(n).Activo = 1
RaidVivos(n).MiniRaids = 9
Call LogCasino("Reaparece MonsterDraG: " & na)
End If

End Sub
Function NextOpenNPC() As Integer
'Call LogTarea("Sub NextOpenNPC")

On Error GoTo errhandler

Dim loopc As Integer
  
For loopc = 1 To MAXNPCS + 1
    If loopc > MAXNPCS Then Exit For
    If Not Npclist(loopc).flags.NPCActive Then Exit For
Next loopc
  
NextOpenNPC = loopc


Exit Function
errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer, ByVal veneno As Integer)
On Error GoTo fallo
Dim n As Integer
'n = RandomNumber(1, 100)
'If n < 30 Then
'nati: agrego el "Not UserList(UserIndex).flags.Morph = 214" para el berserker no sea envenenado.
  If UCase$(UserList(UserIndex).clase) <> "BARDO" And UserList(UserIndex).flags.Angel = 0 And UserList(UserIndex).flags.Demonio = 0 And Not UserList(UserIndex).flags.Morph = 214 Then
    UserList(UserIndex).flags.Envenenado = veneno
    Call SendData(ToIndex, UserIndex, 0, "||La criatura te ha envenenado" & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 119 & "," & 1)

Else
Call SendData(ToIndex, UserIndex, 0, "||La criatura te ha intentado envenenar, pero eres INMUNE" & "´" & FontTypeNames.FONTTYPE_FIGHT)
End If

'End If
Exit Sub
fallo:
Call LogError("npcenvenenauser " & Err.number & " D: " & Err.Description)

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
'Crea un NPC del tipo Npcindex
'Call LogTarea("Sub SpawnNpc")
On Error GoTo fallo
Dim Newpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean


Dim Map As Integer
Dim X As Integer
Dim Y As Integer
Dim it As Integer
Dim it2 As Byte

nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

it = 0
it2 = 0
'pluto:6.0A
If Npclist(NpcIndex).NPCtype = 60 Then it2 = 100


If nIndex > MAXNPCS Then
    SpawnNpc = nIndex
    Exit Function
End If

'pluto:2.17
'If Npclist(nIndex).NPCtype = 78 Then
'Map = Pos.Map
'X = Pos.X
'Y = Pos.Y
'GoTo yipi
'End If

Do While Not PosicionValida
        Call ClosestLegalPos(Pos, Newpos, Npclist(nIndex).flags.AguaValida) 'Nos devuelve la posicion valida mas cercana
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If LegalPosNPC(Newpos.Map, Newpos.X, Newpos.Y, Npclist(nIndex).flags.AguaValida) Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).Pos.Map = Newpos.Map
            Npclist(nIndex).Pos.X = Newpos.X
            Npclist(nIndex).Pos.Y = Newpos.Y
            PosicionValida = True
        Else
            Newpos.X = 0
            Newpos.Y = 0
        End If
        
        it = it + 1
        
        If it > MAXSPAWNATTEMPS + it2 Then
            Call QuitarNPC(nIndex)
            SpawnNpc = MAXNPCS
            Call LogError("Mas de " & MAXSPAWNATTEMPS + it2 & " iteraciones en SpawnNpc Mapa:" & Pos.Map & " Index:" & Npclist(NpcIndex).Name)
            Exit Function
        End If
Loop

'asignamos las nuevas coordenas
Map = Newpos.Map
X = Npclist(nIndex).Pos.X
Y = Npclist(nIndex).Pos.Y
yipi:
'Crea el NPC
Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)
'pluto:2.10
If Map = 192 And Npclist(nIndex).NPCtype = 21 Then Balon = nIndex

If FX Then
    Call SendData(ToMap, 0, Map, "TW" & SND_WARP)
    Call SendData2(ToMap, 0, Map, 22, Npclist(nIndex).Char.CharIndex & "," & FXWARP & "," & 0)
End If

SpawnNpc = nIndex
Exit Function
fallo:
Call LogError("spawnnpc " & Err.number & " D: " & Err.Description)

End Function

Sub ReSpawnNpc(MinPc As npc)
On Error GoTo fallo
If MinPc.flags.Respawn = 0 Then
Call CrearNPC(MinPc.numero, MinPc.Pos.Map, MinPc.Orig)
End If

Exit Sub
fallo:
Call LogError("respawnnpc Nom:" & MinPc.Name & " Map: " & MinPc.Pos.Map & " D: " & Err.Description)
Exit Sub
End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal Map As Integer) As String
On Error GoTo fallo
Dim NpcIndex As Integer
Dim cont As Integer

'Contador
cont = 0
'NPCHostiles = "NPCS en este Mapa: "
For NpcIndex = 1 To LastNPC

    '¿esta vivo?
    If Npclist(NpcIndex).flags.NPCActive _
       And Npclist(NpcIndex).Pos.Map = Map _
       And Npclist(NpcIndex).Hostile = 1 And _
       Npclist(NpcIndex).Stats.Alineacion = 2 Then
            cont = cont + 1
            NPCHostiles = NPCHostiles & Npclist(NpcIndex).Name & "(" & Npclist(NpcIndex).Pos.X & "-" & Npclist(NpcIndex).Pos.Y & ")" & ", "
           
    End If
    
Next NpcIndex

'NPCHostiles = cont
Exit Function
fallo:
Call LogError("npchostiles " & Err.number & " D: " & Err.Description)

End Function

Sub NPCTirarOro(MinPc As npc)
On Error GoTo fallo
'SI EL NPC TIENE ORO LO TIRAMOS
If MinPc.GiveGLD > 0 Then
    Dim MiObj As obj
    MiObj.Amount = MinPc.GiveGLD
    MiObj.ObjIndex = iORO
 Dim alea As Byte
 Dim alea2 As Integer
alea = RandomNumber(1, 20)
alea2 = CInt(MinPc.GiveGLD + (MinPc.GiveGLD * (alea / 100)))
If alea2 > 10000 Then alea2 = alea2 - 2000
'pluto:2.17
If alea2 > 10000 Then alea2 = 10000
MiObj.Amount = alea2


Call TirarItemAlPiso(MinPc.Pos, MiObj)


End If
Exit Sub
fallo:
Call LogError("npctiraroro " & Err.number & " D: " & Err.Description)

End Sub



Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer
On Error GoTo fallo
Dim NpcIndex As Integer
Dim npcfile As String

If NpcNumber > 499 Then
        npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
        npcfile = DatPath & "NPCs.dat"
End If


NpcIndex = NextOpenNPC

If NpcIndex > MAXNPCS Then 'Limite de npcs
    OpenNPC = NpcIndex
    Exit Function
End If



Npclist(NpcIndex).numero = NpcNumber
'pluto:6.0A
Npclist(NpcIndex).Anima = val(GetVar(npcfile, "NPC" & NpcNumber, "Anima"))
Npclist(NpcIndex).Arquero = val(GetVar(npcfile, "NPC" & NpcNumber, "Arquero"))
Npclist(NpcIndex).Raid = val(GetVar(npcfile, "NPC" & NpcNumber, "Raid"))
'pluto:7.0
Npclist(NpcIndex).LogroTipo = val(GetVar(npcfile, "NPC" & NpcNumber, "LogroTipo"))

Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")

Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

Npclist(NpcIndex).flags.AguaValida = val(GetVar(npcfile, "NPC" & NpcNumber, "AguaValida"))
Npclist(NpcIndex).flags.TierraInvalida = val(GetVar(npcfile, "NPC" & NpcNumber, "TierraInValida"))
Npclist(NpcIndex).flags.Faccion = val(GetVar(npcfile, "NPC" & NpcNumber, "Faccion"))

Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP")) * DifServer

Npclist(NpcIndex).veneno = val(GetVar(npcfile, "NPC" & NpcNumber, "Veneno"))

Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))


Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD")) * DifOro

Npclist(NpcIndex).PoderAtaque = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderAtaque"))
Npclist(NpcIndex).PoderEvasion = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderEvasion"))

Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

'@Nati: NPCS a 1 de vida
Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
'Npclist(NpcIndex).Stats.MaxHP = 1
'Npclist(NpcIndex).Stats.MinHP = 1
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NpcNumber, "ImpactRate"))
'Npclist(NpcIndex).Premio = val(GetVar(npcfile, "NPC" & NpcNumber, "Premio"))       'Delzak sistema premios


Dim loopc As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
For loopc = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & loopc)
    Npclist(NpcIndex).Invent.Object(loopc).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(loopc).Amount = val(ReadField(2, ln, 45))
Next loopc

Npclist(NpcIndex).flags.LanzaSpells = val(GetVar(npcfile, "NPC" & NpcNumber, "LanzaSpells"))
If Npclist(NpcIndex).flags.LanzaSpells > 0 Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
For loopc = 1 To Npclist(NpcIndex).flags.LanzaSpells
    Npclist(NpcIndex).Spells(loopc) = val(GetVar(npcfile, "NPC" & NpcNumber, "Sp" & loopc))
Next loopc


If Npclist(NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
    Npclist(NpcIndex).NroCriaturas = val(GetVar(npcfile, "NPC" & NpcNumber, "NroCriaturas"))
    ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
    For loopc = 1 To Npclist(NpcIndex).NroCriaturas
        Npclist(NpcIndex).Criaturas(loopc).NpcIndex = GetVar(npcfile, "NPC" & NpcNumber, "CI" & loopc)
        Npclist(NpcIndex).Criaturas(loopc).NpcName = GetVar(npcfile, "NPC" & NpcNumber, "CN" & loopc)
    Next loopc
End If


Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Inflacion"))

Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False

If Respawn Then
    Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Else
    Npclist(NpcIndex).flags.Respawn = 1
End If

Npclist(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "PosOrig"))
Npclist(NpcIndex).flags.AfectaParalisis = val(GetVar(npcfile, "NPC" & NpcNumber, "AfectaParalisis"))
'pluto:2.14
Npclist(NpcIndex).flags.PoderEspecial1 = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderEspecial1"))
Npclist(NpcIndex).flags.PoderEspecial2 = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderEspecial2"))
Npclist(NpcIndex).flags.PoderEspecial3 = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderEspecial3"))
Npclist(NpcIndex).flags.PoderEspecial4 = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderEspecial4"))
Npclist(NpcIndex).flags.PoderEspecial5 = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderEspecial5"))
Npclist(NpcIndex).flags.PoderEspecial6 = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderEspecial6"))
Npclist(NpcIndex).flags.Magiainvisible = val(GetVar(npcfile, "NPC" & NpcNumber, "Magiainvisible"))
Npclist(NpcIndex).flags.GolpeExacto = val(GetVar(npcfile, "NPC" & NpcNumber, "GolpeExacto"))


Npclist(NpcIndex).flags.Snd1 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd1"))
Npclist(NpcIndex).flags.Snd2 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd2"))
Npclist(NpcIndex).flags.Snd3 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd3"))
Npclist(NpcIndex).flags.Snd4 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd4"))

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

Dim aux As String
aux = GetVar(npcfile, "NPC" & NpcNumber, "NROEXP")
If aux = "" Then
    Npclist(NpcIndex).NroExpresiones = 0
Else
    Npclist(NpcIndex).NroExpresiones = val(aux)
    ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
    For loopc = 1 To Npclist(NpcIndex).NroExpresiones
        Npclist(NpcIndex).Expresiones(loopc) = GetVar(npcfile, "NPC" & NpcNumber, "Exp" & loopc)
    Next loopc
End If

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

'Update contadores de NPCs
If NpcIndex > LastNPC Then LastNPC = NpcIndex
NumNPCs = NumNPCs + 1

'pluto:2.17
Dim li As Byte
Dim bi As Byte
Dim bi2 As Byte

'pluto:6.8
If Npclist(NpcIndex).numero = BichoDelDia And EventoDia = 1 Then
li = RandomNumber(194, 202)
Else
li = RandomNumber(1, 200)
End If




bi = 1
bi2 = 1
If li > 193 And NpcNumber > 499 And Npclist(NpcIndex).NPCtype <> 60 Then

If li > 193 And li < 197 Then bi = 2: bi2 = 3
If li > 196 And li < 199 Then bi = 3: bi2 = 5
If li > 198 Then bi = 4: bi2 = 7
'pluto:6.8
If li > 200 Then bi = 5: bi2 = 10

'pluto:6.0A evitar bichos fuertes en ciertos npcs
If Npclist(NpcIndex).numero = 699 Or Npclist(NpcIndex).numero = 692 Or Npclist(NpcIndex).numero = 666 Or Npclist(NpcIndex).numero = 667 Or Npclist(NpcIndex).numero = 587 Or Npclist(NpcIndex).numero = 594 Or Npclist(NpcIndex).numero = 633 Or Npclist(NpcIndex).numero = 611 Or Npclist(NpcIndex).numero = 585 Or Npclist(NpcIndex).numero = 621 Or Npclist(NpcIndex).numero = 676 Or Npclist(NpcIndex).numero = 726 Or Npclist(NpcIndex).NPCtype = 6 Or Npclist(NpcIndex).Raid > 0 Then bi = 1: bi2 = 1

If bi = 2 Then Npclist(NpcIndex).Name = Npclist(NpcIndex).Name & " >> Mejorado <<"
If bi = 3 Then Npclist(NpcIndex).Name = Npclist(NpcIndex).Name & " >>  Líder <<"
If bi = 4 Then Npclist(NpcIndex).Name = Npclist(NpcIndex).Name & " >> Especial <<"
If bi = 5 Then Npclist(NpcIndex).Name = Npclist(NpcIndex).Name & " >> Legendario <<"


Npclist(NpcIndex).GiveEXP = Npclist(NpcIndex).GiveEXP * bi2
Npclist(NpcIndex).flags.Domable = 0
Npclist(NpcIndex).GiveGLD = Npclist(NpcIndex).GiveGLD * bi2
If Npclist(NpcIndex).GiveGLD > 10000 Then Npclist(NpcIndex).GiveGLD = 10000
Npclist(NpcIndex).PoderAtaque = Npclist(NpcIndex).PoderAtaque * bi
Npclist(NpcIndex).PoderEvasion = Npclist(NpcIndex).PoderEvasion * bi
Npclist(NpcIndex).Stats.MaxHP = Npclist(NpcIndex).Stats.MaxHP * bi
Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP * bi
Npclist(NpcIndex).Stats.MaxHIT = Npclist(NpcIndex).Stats.MaxHIT * bi
Npclist(NpcIndex).Stats.MinHIT = Npclist(NpcIndex).Stats.MinHIT * bi
Npclist(NpcIndex).Stats.Def = Npclist(NpcIndex).Stats.Def * bi
End If

'pluto:6.8 eventodia 3 y 4
'If Npclist(NpcIndex).Numero = 718 And EventoDia = 3 Then
'Npclist(NpcIndex).Stats.MaxHP = 500
'Npclist(NpcIndex).Stats.MinHP = 500
'Npclist(NpcIndex).Stats.MaxHIT = 1
'Npclist(NpcIndex).Stats.MinHIT = 1
'Npclist(NpcIndex).PoderEvasion = 1
'Npclist(NpcIndex).GiveGLD = 10000
'Npclist(NpcIndex).Name = "Regalo de Dioses"
'End If
'pluto:6.8 arregla que solo sea para el bichodeldia
If Npclist(NpcIndex).numero = BichoDelDia And EventoDia = 4 Then
Npclist(NpcIndex).Stats.MaxHP = Int(Npclist(NpcIndex).Stats.MaxHP / 2)
Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP

If Npclist(NpcIndex).Stats.MaxHP < 1 Then Npclist(NpcIndex).Stats.MaxHP = 1
End If
'-------------------------------------
'Devuelve el nuevo Indice
OpenNPC = NpcIndex



Exit Function
fallo:
Call LogError("opennpc " & Err.number & " D: " & Err.Description)

End Function
Sub EnviarListaCriaturas(ByVal UserIndex As Integer, ByVal NpcIndex)
  On Error GoTo fallo
  Dim SD As String
  Dim k As Integer
  SD = SD & Npclist(NpcIndex).NroCriaturas & ","
  For k = 1 To Npclist(NpcIndex).NroCriaturas
        SD = SD & Npclist(NpcIndex).Criaturas(k).NpcName & ","
  Next k
  Call SendData2(ToIndex, UserIndex, 0, 58, SD)
Exit Sub
fallo:
Call LogError("enviarlistacriaturas " & Err.number & " D: " & Err.Description)

End Sub


Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
On Error GoTo fallo
If Npclist(NpcIndex).flags.Follow Then
  Npclist(NpcIndex).flags.AttackedBy = ""
  Npclist(NpcIndex).flags.Follow = False
  Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
  Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
Else
  Npclist(NpcIndex).flags.AttackedBy = UserName
  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = 4 'follow
  Npclist(NpcIndex).Hostile = 0
End If
Exit Sub
fallo:
Call LogError("dofollow " & Err.number & " D: " & Err.Description)

End Sub

Sub FollowAmo(ByVal NpcIndex As Integer)
On Error GoTo fallo
  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = SIGUE_AMO 'follow
  Npclist(NpcIndex).Hostile = 0
  Npclist(NpcIndex).Target = 0
  Npclist(NpcIndex).TargetNpc = 0
Exit Sub
fallo:
Call LogError("followamo " & Err.number & " D: " & Err.Description)

End Sub

