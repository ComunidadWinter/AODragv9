Attribute VB_Name = "NPCs"
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
    Dim i As Integer
    
    For i = 1 To MAXMASCOTAS
      If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
         UserList(UserIndex).MascotasIndex(i) = 0
         UserList(UserIndex).MascotasType(i) = 0
         
         UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
         Exit For
      End If
    Next i
End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
    NPCList(Maestro).Mascotas = NPCList(Maestro).Mascotas - 1
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
  '********************************************************
  'Author: Unknown
  'Llamado cuando la vida de un NPC llega a cero.
  'Last Modify Date: 24/01/2007
  '22/06/06: (Nacho) Chequeamos si es pretoriano
  '24/01/2007: Pablo (ToxicWaste): Agrego para actualización de tag si cambia de status.
  '********************************************************

  '28/10/2015 Irongete: Si el NPC que muere es un Dummy.. no muere y le restauro la vida al 100%
  If NPCList(NpcIndex).Numero = 624 Then
    NPCList(NpcIndex).Stats.MinHP = NPCList(NpcIndex).Stats.MaxHP
    Exit Sub
  End If

  On Error GoTo errHandler
  Dim MiNPC As npc
  MiNPC = NPCList(NpcIndex)
  Dim EraCriminal As Boolean
  Dim NumeroNPC   As Integer
    
  NumeroNPC = MiNPC.Numero
    
  'SALA DE INVOCACIONES
  If MiNPC.Pos.Map = MapInvocacion Then MapInfo(MapInvocacion).Invocado = 0
  '/SALA DE INVOCACIONES
    
  'Respawn de NPC con retardo
  If NPCList(NpcIndex).flags.Retardo = 1 Then
    RetardoSpawn(MiNPC.Numero).Tiempo = RandomNumber(MiNPC.flags.TiempoRetardoMin, MiNPC.flags.TiempoRetardoMax)
    RetardoSpawn(MiNPC.Numero).Mapa = MiNPC.Orig.Map
    RetardoSpawn(MiNPC.Numero).X = MiNPC.Orig.X
    RetardoSpawn(MiNPC.Numero).Y = MiNPC.Orig.Y
    RetardoSpawn(MiNPC.Numero).NPCNUM = MiNPC.Numero
  End If
  '/Respawn de NPC con retardo
    
  '¿Lo mato un usuario?
  If UserIndex > 0 Then

    With UserList(UserIndex)

      '¿El NPC explota al matarlo?
      If NPCList(NpcIndex).flags.Explota = 1 Then
        Call UserDie(UserIndex)
                     
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(27, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 27, 0))
        Call WriteConsoleMsg(UserIndex, "¡La explosion del Bomber te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
      End If
           
      '¿El NPC corresponde al NPC del castillo?
      If MiNPC.Numero = NPCReyCastle Then
        Call ReyMuere(UserIndex, NpcIndex)
        Exit Sub
      End If
            
      '26/02/2016 Irongete: Comprobar si el NPC es el Defensor de la Fortaleza
      If MiNPC.Numero = NPCDefensorFortaleza Then
        Call DefensorMuere(UserIndex, NpcIndex)
        Exit Sub
      End If
            
      If MiNPC.NPCType = 10 Then
        Call PuertaEsDestruida(UserIndex, NpcIndex)
        Exit Sub
      End If
        
      'Si el NPC tiene sonido de muerte lo reproducimos
      If MiNPC.flags.Snd3 > 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y))
      End If
            
      .flags.TargetNPC = 0
      .flags.TargetNpcTipo = eNPCType.Comun
            
      'El user que lo mato tiene mascotas?
      If .NroMascotas > 0 Then
        Dim T As Integer

        For T = 1 To MAXMASCOTAS

          If .MascotasIndex(T) > 0 Then
            If NPCList(.MascotasIndex(T)).TargetNPC = NpcIndex Then
              Call FollowAmo(.MascotasIndex(T))
            End If
          End If
        Next T

      End If
            
      '¿El NPC soltaba experiencia?
      If MiNPC.flags.ExpCount > 0 Then
        Dim ExpFinal As Long
               
        ExpFinal = MiNPC.flags.ExpCount
               
        '29/10/2015 Irongete: Añadir mas experiencia en funcion de los usuarios que estan online
        ExpFinal = ExpFinal + ((ExpFinal * NumUsers) / 100)
               
        '13/02/2016 Lorwik: Por cada mimbro de la GUILD online, adquieres 1% mas de exp
        'If Not UserList(UserIndex).GuildIndex = 0 Then ExpFinal = ExpFinal + ((ExpFinal * modGuilds.m_ListaDeMiembrosOnline(UserIndex, UserList(UserIndex).GuildIndex) / 100))
               
        '18/11/2015 Irongete: Modificamos la ExpFinal dependiendo si está en party o no
        If .PartyId > 0 Then
          Call Drag_Party.RepartirExpParty(UserIndex, ExpFinal, .PartyId, MiNPC.Pos.X, MiNPC.Pos.Y, MiNPC.Nivel)
                   
        Else 'Si no esta en Party le damos la experiencia directamente.
               
          'Si el usuario esta por encima del nivel del NPC le damos menos exp.
          If UserList(UserIndex).Stats.ELV > MiNPC.Nivel Then
            ExpFinal = MiNPC.flags.ExpCount / 10
          End If
                    
          'Si hay una diferencia de niveles mayor a 10, te da solo un 10% de exp
          'Dim DiferenciadeNivel As Byte
          'DiferenciadeNivel = UserList(UserIndex).Stats.ELV - MiNPC.Nivel
                    
          'If UserList(UserIndex).Stats.ELV < MiNPC.Nivel And DiferenciadeNivel >= 10 Then
          '    ExpFinal = Porcentaje(MiNPC.flags.ExpCount, 10)
          'End If
                    
          .Stats.Exp = .Stats.Exp + ExpFinal

          If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
          Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpFinal & " puntos de experiencia.", FontTypeNames.FONTTYPE_exp)
                    
          '<Edurne>
          'Si tiene montura le damos experiencia.
          If .flags.QueMontura Then
            .flags.Montura(.flags.QueMontura).Exp = .flags.Montura(.flags.QueMontura).Exp + (ExpFinal / 2)
            Call WriteConsoleMsg(UserIndex, "Tu montura gano " & (ExpFinal / 2) & " puntos de experiencia.", FontTypeNames.FONTTYPE_exp)
            Call CheckMonturaLevel(UserIndex)
          End If
        End If
        MiNPC.flags.ExpCount = 0
        ExpFinal = 0
      End If
            
      'Notificamos la muerte de la criatura
      Call WriteConsoleMsg(UserIndex, "¡Has matado a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
      'Le damos el oro que tiraba la criatura
      Call NPCTirarOro(MiNPC, UserIndex)
            
      'Sumamos una muerte al contador de criaturas matadas
      If .Stats.NPCsMuertos < 32000 Then .Stats.NPCsMuertos = .Stats.NPCsMuertos + 1
            
      'Dependiendo de que criatura fuera, le quitamos criminabilidad o se la sumamos
      EraCriminal = criminal(UserIndex)
            
      If MiNPC.Stats.Alineacion = 0 Then
        If MiNPC.Numero = Guardias Then
          .Reputacion.NobleRep = 0
          .Reputacion.PlebeRep = 0
          .Reputacion.AsesinoRep = .Reputacion.AsesinoRep + 500

          If .Reputacion.AsesinoRep > MAXREP Then .Reputacion.AsesinoRep = MAXREP
        End If

        If MiNPC.MaestroUser = 0 Then
          .Reputacion.AsesinoRep = .Reputacion.AsesinoRep + vlASESINO

          If .Reputacion.AsesinoRep > MAXREP Then .Reputacion.AsesinoRep = MAXREP
        End If
      ElseIf MiNPC.Stats.Alineacion = 1 Then

        If esCaos(UserIndex) = False Then
          .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlCAZADOR

          If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        End If
      ElseIf MiNPC.Stats.Alineacion = 2 Then

        If esCaos(UserIndex) = False Then
          .Reputacion.NobleRep = .Reputacion.NobleRep + vlASESINO / 2

          If .Reputacion.NobleRep > MAXREP Then .Reputacion.NobleRep = MAXREP
        End If
      ElseIf MiNPC.Stats.Alineacion = 4 Then

        If esCaos(UserIndex) = False Then
          .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlCAZADOR

          If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        End If
      End If

      If criminal(UserIndex) And esArmada(UserIndex) Then Call ExpulsarFaccionReal(UserIndex)
      If Not criminal(UserIndex) And esCaos(UserIndex) Then Call ExpulsarFaccionCaos(UserIndex)
            
      If EraCriminal And Not criminal(UserIndex) Then
        Call RefreshCharStatus(UserIndex)
      ElseIf Not EraCriminal And criminal(UserIndex) Then
        Call RefreshCharStatus(UserIndex)
      End If
            
      '¿Subiria de nivel?
      Call CheckUserLevel(UserIndex)
            
    End With
  End If ' Userindex > 0
     
  'Quitamos el NPC
  Call QuitarNPC(NpcIndex)
    
  If MiNPC.MaestroUser = 0 Then
    'Tiramos el inventario
    Call NPC_TIRAR_ITEMS(MiNPC, UserIndex)
    
    '09/12/2018 Irongete: Si pertenece a una zona el respawn es diferente
    If MiNPC.flags.zona > 0 Then
      Call RespawnNpcZona(MiNPC)
      Exit Sub
    Else
      If MiNPC.flags.Retardo = 0 Then
        'ReSpawn o no
        Call ReSpawnNpc(MiNPC)
      End If
    End If
  End If
   
  'Si esta en arenas y las arenas esta en curso, comprobamos cuandos bichos quedan.
  If UserList(UserIndex).flags.ArenaRinkel = True Then Call BichosVivos(False)
    
  Exit Sub

errHandler:
  Call LogError("Error en MuereNpc - Error: " & Err.Number & " - Desc: " & Err.Description)
End Sub

Private Sub ResetNpcFlags(ByVal NpcIndex As Integer)
    'Clear the npc's flags
    
    With NPCList(NpcIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedFirstBy = vbNullString
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .AtacaDoble = 0
        .LanzaSpells = 0
        .invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
    End With
End Sub

Private Sub ResetNpcCounters(ByVal NpcIndex As Integer)
    With NPCList(NpcIndex).Contadores
        .Paralisis = 0
        .TiempoExistencia = 0
    End With
End Sub

Private Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
    With NPCList(NpcIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Private Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
    Dim j As Long
    
    With NPCList(NpcIndex)
        For j = 1 To .NroCriaturas
            .Criaturas(j).NpcIndex = 0
            .Criaturas(j).NpcName = vbNullString
        Next j
        
        .NroCriaturas = 0
    End With
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
    Dim j As Long
    
    With NPCList(NpcIndex)
        For j = 1 To .NroExpresiones
            .Expresiones(j) = vbNullString
        Next j
        
        .NroExpresiones = 0
    End With
End Sub

Private Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
    With NPCList(NpcIndex)
        .Attackable = 0
        .CanAttack = 0
        .Comercia = 0
        .GiveEXP = 0
        .GiveGLD = 0
        .Hostile = 0
        .InvReSpawn = 0
        
        If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, NpcIndex)
        If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)
        
        .MaestroUser = 0
        .MaestroNpc = 0
        
        .Mascotas = 0
        .Movement = 0
        .Name = vbNullString
        .NPCType = 0
        .Numero = 0
        .Orig.Map = 0
        .Orig.X = 0
        .Orig.Y = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .SkillDomar = 0
        .Target = 0
        .TargetNPC = 0
        .TipoItems = 0
        .Veneno = 0
        .desc = vbNullString
        
        
        Dim j As Long
        For j = 1 To .NroSpells
            .Spells(j) = 0
        Next j
    End With
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)
End Sub

Public Sub QuitarNPC(ByVal NpcIndex As Integer)

On Error GoTo errHandler

    With NPCList(NpcIndex)
        .flags.Active = False
        
        If InMapBounds(.Pos.Map, .Pos.X, .Pos.Y) Then
            Call EraseNPCChar(NpcIndex)
        End If
    End With
    
    'Si el que esta siendo borrado es el Saqueador, reseteamos el index
    If NPCList(NpcIndex).Numero = NumSaqueador Then _
        SaqueadorIndex = 0
    
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then
        Do Until NPCList(LastNPC).flags.Active
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
        
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If
    
Exit Sub

errHandler:
    Call LogError("Error en QuitarNPC")
End Sub

Public Sub QuitarPet(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 18/11/2009
'Kills a pet
'***************************************************
On Error GoTo errHandler

    Dim i As Integer
    Dim PetIndex As Integer

    With UserList(UserIndex)
        
        ' Busco el indice de la mascota
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) = NpcIndex Then PetIndex = i
        Next i
        
        ' Poco probable que pase, pero por las dudas..
        If PetIndex = 0 Then Exit Sub
        
        ' Limpio el slot de la mascota
        .NroMascotas = .NroMascotas - 1
        .MascotasIndex(PetIndex) = 0
        .MascotasType(PetIndex) = 0
        
        ' Elimino la mascota
        Call QuitarNPC(NpcIndex)
    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en QuitarPet. Error: " & Err.Number & " Desc: " & Err.Description & " NpcIndex: " & NpcIndex & " UserIndex: " & UserIndex & " PetIndex: " & PetIndex)
End Sub

Private Function TestSpawnTrigger(Pos As worldPos, Optional PuedeAgua As Boolean = False) As Boolean
    
    If LegalPos(Pos.Map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = _
        MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 3 And _
        MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 2 And _
        MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 1
    End If

End Function

Sub CrearNPC(NroNPC As Integer, Mapa As Integer, OrigPos As worldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC

Dim Pos As worldPos
Dim newpos As worldPos
Dim altpos As worldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim Map As Integer
Dim X As Integer
Dim Y As Integer

    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
    If nIndex > MAXNPCS Then Exit Sub
    PuedeAgua = NPCList(nIndex).flags.AguaValida
    PuedeTierra = IIf(NPCList(nIndex).flags.TierraInvalida = 1, False, True)
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        
        Map = OrigPos.Map
        X = OrigPos.X
        Y = OrigPos.Y
        NPCList(nIndex).Orig = OrigPos
        NPCList(nIndex).Pos = OrigPos
       
    Else
        
        Pos.Map = Mapa 'mapa
        altpos.Map = Mapa
        
        Do While Not PosicionValida
            Pos.X = RandomNumber(MinXBorder, MaxXBorder)    'Obtenemos posicion al azar en x
            Pos.Y = RandomNumber(MinYBorder, MaxYBorder)    'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
            If newpos.X <> 0 And newpos.Y <> 0 Then
                altpos.X = newpos.X
                altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn, pero intentando qeu si tenía que ser en el agua, sea en el agua.)
            Else
                Call ClosestLegalPos(Pos, newpos, PuedeAgua)
                If newpos.X <> 0 And newpos.Y <> 0 Then
                    altpos.X = newpos.X
                    altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)
                End If
            End If
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, PuedeAgua) And _
               Not HayPCarea(newpos) And TestSpawnTrigger(newpos, PuedeAgua) Then
                'Asignamos las nuevas coordenas solo si son validas
                NPCList(nIndex).Pos.Map = newpos.Map
                NPCList(nIndex).Pos.X = newpos.X
                NPCList(nIndex).Pos.Y = newpos.Y
                PosicionValida = True
            Else
                newpos.X = 0
                newpos.Y = 0
            
            End If
                
            'for debug
            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.X <> 0 And altpos.Y <> 0 Then
                    Map = altpos.Map
                    X = altpos.X
                    Y = altpos.Y
                    NPCList(nIndex).Pos.Map = Map
                    NPCList(nIndex).Pos.X = X
                    NPCList(nIndex).Pos.Y = Y
                    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
                    Exit Sub
                Else
                    altpos.X = 50
                    altpos.Y = 50
                    Call ClosestLegalPos(altpos, newpos)
                    If newpos.X <> 0 And newpos.Y <> 0 Then
                        NPCList(nIndex).Pos.Map = newpos.Map
                        NPCList(nIndex).Pos.X = newpos.X
                        NPCList(nIndex).Pos.Y = newpos.Y
                        Call MakeNPCChar(True, newpos.Map, nIndex, newpos.Map, newpos.X, newpos.Y)
                        Exit Sub
                    Else
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & Mapa & " NroNpc:" & NroNPC)
                        Exit Sub
                    End If
                End If
            End If
        Loop
        
        'asignamos las nuevas coordenas
        Map = newpos.Map
        X = NPCList(nIndex).Pos.X
        Y = NPCList(nIndex).Pos.Y
    End If
    
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)

End Sub

Public Sub MakeNPCChar(ByVal toMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
Dim CharIndex As Integer
Dim nombre As String
Dim bType As Byte
Dim Quest As eEstadoQuest

    With NPCList(NpcIndex)
        If .Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            .Char.CharIndex = CharIndex
            CharList(CharIndex) = NpcIndex
        End If
        
        MapData(Map, X, Y).NpcIndex = NpcIndex
        bType = NPCList(NpcIndex).Hostile
        
        If .Hostile = 0 Or .NPCType = MonsterDrag Then nombre = .Name
        
        
        '14/11/2018 calcular el estado de quest para el usuario
        If NPCList(NpcIndex).Quest = 1 Then
           Quest = Drag_Quest.EstadoQuest(NPCList(NpcIndex).Numero, UserList(sndIndex).id)
        End If
        
        
        If Not toMap Then
            Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.heading, .Char.CharIndex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, nombre, vbNullString, .Stats.TipoNick, 0, bType, NPCList(NpcIndex).Char.AnimAtaque, .flags.Speed, 0, NPCList(NpcIndex).NPCType, NPCList(NpcIndex).Numero, Quest)
            Call FlushBuffer(sndIndex)
        Else
            Call AgregarNpc(NpcIndex)
        End If
    End With
End Sub

Public Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading)
    If NpcIndex > 0 Then
        With NPCList(NpcIndex).Char
            .body = body
            .Head = Head
            .heading = heading
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, 0, 0, 0, 0, 0))
        End With
    End If
End Sub

Private Sub EraseNPCChar(ByVal NpcIndex As Integer)

If NPCList(NpcIndex).Char.CharIndex <> 0 Then CharList(NPCList(NpcIndex).Char.CharIndex) = 0

If NPCList(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar <= 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(NPCList(NpcIndex).Pos.Map, NPCList(NpcIndex).Pos.X, NPCList(NpcIndex).Pos.Y).NpcIndex = 0

'Actualizamos los clientes
Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(NPCList(NpcIndex).Char.CharIndex))

'Update la lista npc
NPCList(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Public Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 06/04/2009
'06/04/2009: ZaMa - Now npcs can force to change position with dead character
'01/08/2009: ZaMa - Now npcs can't force to chance position with a dead character if that means to change the terrain the character is in
'***************************************************

On Error GoTo errh
    Dim nPos As worldPos
    Dim UserIndex As Integer
    
    With NPCList(NpcIndex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)
        
        ' es una posicion legal
        If LegalPosNPC(.Pos.Map, nPos.X, nPos.Y, .flags.AguaValida = 1, .MaestroUser <> 0) Then
            
            If .flags.AguaValida = 0 And HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Sub
            If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Sub
            
            UserIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex
            ' Si hay un usuario a donde se mueve el npc, entonces esta muerto
            If UserIndex > 0 Then
                
                ' No se traslada caspers de agua a tierra
                If HayAgua(.Pos.Map, nPos.X, nPos.Y) And Not HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Sub
                ' No se traslada caspers de tierra a agua
                If Not HayAgua(.Pos.Map, nPos.X, nPos.Y) And HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Sub
                
                With UserList(UserIndex)
                    ' Actualizamos posicion y mapa
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
                    .Pos.X = NPCList(NpcIndex).Pos.X
                    .Pos.Y = NPCList(NpcIndex).Pos.Y
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                    ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, .Pos.X, .Pos.Y))
                    Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))
                End With
            End If
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))

            'Update map and user pos
            MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex = 0
            .Pos = nPos
            .Char.heading = nHeading
            MapData(.Pos.Map, nPos.X, nPos.Y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
        ElseIf .MaestroUser = 0 Then
            If .Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                .PFINFO.PathLenght = 0
            End If
        End If
    End With
Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)
End Sub

Function NextOpenNPC() As Integer
'Call LogTarea("Sub NextOpenNPC")

On Error GoTo errHandler
    Dim LoopC As Long
      
    For LoopC = 1 To MAXNPCS + 1
        If LoopC > MAXNPCS Then Exit For
        If Not NPCList(LoopC).flags.Active Then Exit For
    Next LoopC
      
    NextOpenNPC = LoopC
Exit Function

errHandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)

Dim n As Integer
n = RandomNumber(1, 100)
If n < 30 Then
    UserList(UserIndex).flags.Envenenado = 1
    Call WriteConsoleMsg(UserIndex, "¡¡La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
End If

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As worldPos, ByVal FX As Boolean, ByVal Respawn As Boolean, Optional ByVal OrigPos As Boolean = False, Optional ByVal IncrementoVida As Integer = 0) As Integer
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 06/15/2008
'23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
'06/15/2008 -> Optimizé el codigo. (NicoNZ)
'***************************************************
Dim newpos As worldPos
Dim altpos As worldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer
    
    nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice
    
    If nIndex > MAXNPCS Then
        SpawnNpc = 0
        Exit Function
    End If
    
    '¿Hay usuario en esa pos?
    If MapData(Pos.Map, Pos.X, Pos.Y).UserIndex > 0 Then _
        Call WarpUserChar(MapData(Pos.Map, Pos.X, Pos.Y).UserIndex, Pos.Map, Pos.X + 1, Pos.Y + 1, False)
        
    PuedeAgua = NPCList(nIndex).flags.AguaValida
    PuedeTierra = Not NPCList(nIndex).flags.TierraInvalida = 1
                
    Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
    Call ClosestLegalPos(Pos, altpos, PuedeAgua)
    'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        
    If newpos.X <> 0 And newpos.Y <> 0 Then
        'Asignamos las nuevas coordenas solo si son validas
        NPCList(nIndex).Pos.Map = newpos.Map
        NPCList(nIndex).Pos.X = newpos.X
        NPCList(nIndex).Pos.Y = newpos.Y
        PosicionValida = True
    Else
        If altpos.X <> 0 And altpos.Y <> 0 Then
            NPCList(nIndex).Pos.Map = altpos.Map
            NPCList(nIndex).Pos.X = altpos.X
            NPCList(nIndex).Pos.Y = altpos.Y
            PosicionValida = True
        Else
            PosicionValida = False
        End If
    End If
    
    If Not PosicionValida Then
        Call QuitarNPC(nIndex)
        SpawnNpc = 0
        Exit Function
    End If
    
    'asignamos las nuevas coordenas
    Map = newpos.Map
    X = NPCList(nIndex).Pos.X
    Y = NPCList(nIndex).Pos.Y
    
    '17/02/2016 - Lorwik: Se utiliza principalmente para los NPC con retardo de Spawn
    If OrigPos Then
        NPCList(nIndex).Orig.Map = Map
        NPCList(nIndex).Orig.X = X
        NPCList(nIndex).Orig.Y = Y
    End If
    
    If IncrementoVida > 0 Then
        NPCList(nIndex).Stats.MaxHP = NPCList(nIndex).Stats.MaxHP * IncrementoVida
        NPCList(nIndex).Stats.MinHP = NPCList(nIndex).Stats.MinHP * IncrementoVida
    End If
    
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
    
    If FX Then
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(NPCList(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))
    End If
    
    SpawnNpc = nIndex
    
    'Si el que esta haciendo spawn es el saqueador, guardamos su Index
    If NpcIndex = NumSaqueador Then _
        SaqueadorIndex = nIndex

End Function

Sub ReSpawnNpc(MiNPC As npc)

If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

End Sub

Private Sub NPCTirarOro(ByRef MiNPC As npc, ByVal UserIndex As Integer)
Dim GLDFinal As Long
'SI EL NPC TIENE ORO LO TIRAMOS
    If MiNPC.GiveGLD > 0 Then
    
        GLDFinal = MiNPC.GiveGLD
        
        '13/02/2016 Lorwik: Por cada mimbro de la GUILD online, adquieres 1% mas de oro
        'If Not UserList(UserIndex).GuildIndex = 0 Then GLDFinal = GLDFinal + ((GLDFinal * modGuilds.m_ListaDeMiembrosOnline(UserIndex, UserList(UserIndex).GuildIndex) / 100))
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + GLDFinal
        Call WriteUpdateUserStats(UserIndex)
        Call WriteConsoleMsg(UserIndex, "¡Obtienes " & GLDFinal & " monedas de oro!", FontTypeNames.FONTTYPE_oro)
    End If
End Sub

Public Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'conmigo. Para leer los NPCS se deberá usar la
'nueva clase clsIniReader.
'
'Alejo
'
'###################################################
    Dim NpcIndex As Integer
    Dim Leer As clsIniReader
    Dim LoopC As Long
    Dim ln As String
    Dim aux As String
    
    Set Leer = LeerNPCs
    
    'If requested index is invalid, abort
    If Not Leer.KeyExists("NPC" & NpcNumber) Then
        OpenNPC = MAXNPCS + 1
        Exit Function
    End If
    
    NpcIndex = NextOpenNPC
    
    If NpcIndex > MAXNPCS Then 'Limite de npcs
        OpenNPC = NpcIndex
        Exit Function
    End If
    
    With NPCList(NpcIndex)
        .Numero = NpcNumber
        .Name = Leer.GetValue("NPC" & NpcNumber, "Name")
        .desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
        .Nivel = Leer.GetValue("NPC" & NpcNumber, "Nivel")
        .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
        .flags.OldMovement = .Movement
        
        .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
        .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
        .flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
        .flags.AtacaDoble = val(Leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))
        
        .NPCType = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
        
        .Char.body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
        .Char.AnimAtaque = val(Leer.GetValue("NPC" & NpcNumber, "AnimAtaque"))

        .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
        .Char.heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
        .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "ShieldAnim"))
        .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "WeaponAnim "))
        .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "CascoAnim"))
        
        .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
        .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
        .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
        .flags.OldHostil = .Hostile
        
        .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP"))
        
        .flags.ExpCount = .GiveEXP
        
        .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
        
        .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
        
        .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))
        
        .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
        .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
        
        .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
        
        With .Stats
            .MaxHP = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
            .MinHP = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
            .MaxHIT = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
            .MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
            .def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
            .defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
            .Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
            .TipoNick = val(Leer.GetValue("NPC" & NpcNumber, "TipoNick"))
        End With
        
        .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
        For LoopC = 1 To .Invent.NroItems
            ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
            .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
            .Invent.Object(LoopC).ProbTirar = val(ReadField(3, ln, 45))
        Next LoopC
        
        .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
        If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)
        For LoopC = 1 To .flags.LanzaSpells
            .Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
        Next LoopC
        
        If .NPCType = eNPCType.Entrenador Then
            .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
            For LoopC = 1 To .NroCriaturas
                .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
                .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
            Next LoopC
        End If
        
        With .flags
            .Active = True
            
            If Respawn Then
                .Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
            Else
                .Respawn = 1
            End If
            
            .BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
            .RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
            .AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
            
            .Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
            .Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
            .Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))
            
            .SoloParty = val(Leer.GetValue("NPC" & NpcNumber, "SoloParty"))
            .Mensaje = Leer.GetValue("NPC" & NpcNumber, "Mensaje")
            .LanzaMensaje = val(Leer.GetValue("NPC" & NpcNumber, "LanzaMensaje"))
            .AumentaPotencia = val(Leer.GetValue("NPC" & NpcNumber, "AumentaPotencia"))
            .TiempoRetardoMax = val(Leer.GetValue("NPC" & NpcNumber, "TiempoRetardoMax"))
            .TiempoRetardoMin = val(Leer.GetValue("NPC" & NpcNumber, "TiempoRetardoMin"))
            .Retardo = val(Leer.GetValue("NPC" & NpcNumber, "Retardo"))
            .Explota = val(Leer.GetValue("NPC" & NpcNumber, "Explota"))
            .VerInvi = val(Leer.GetValue("NPC" & NpcNumber, "VerInvi"))
            .ArenasRinkel = val(Leer.GetValue("NPC" & NpcNumber, "ArenasRinkel"))
            
            .ActivoPotencia = False
            .DijoMensaje = False
        End With
        
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        .NroExpresiones = val(Leer.GetValue("NPC" & NpcNumber, "NROEXP"))
        If .NroExpresiones > 0 Then ReDim .Expresiones(1 To .NroExpresiones) As String
        For LoopC = 1 To .NroExpresiones
            .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
        Next LoopC
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        
        'Tipo de items con los que comercia
        .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
        
        
        .Quest = val(Leer.GetValue("NPC" & NpcNumber, "Quest"))
    End With
    
    'Update contadores de NPCs
    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNPCs = NumNPCs + 1
    
    'Devuelve el nuevo Indice
    OpenNPC = NpcIndex
End Function

Public Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
    With NPCList(NpcIndex)
        If .flags.Follow Then
            .flags.AttackedBy = vbNullString
            .flags.Follow = False
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
        Else
            .flags.AttackedBy = UserName
            .flags.Follow = True
            .Movement = TipoAI.NPCDEFENSA
            .Hostile = 0
        End If
    End With
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
    With NPCList(NpcIndex)
        .flags.Follow = True
        .Movement = TipoAI.SigueAmo
        .Hostile = 0
        .Target = 0
        .TargetNPC = 0
    End With
End Sub
