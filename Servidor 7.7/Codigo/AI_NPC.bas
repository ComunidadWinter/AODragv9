Attribute VB_Name = "AI"
Option Explicit

Public Const ESTATICO = 1
Public Const MUEVE_AL_AZAR = 2
Public Const NPC_MALO_ATACA_USUARIOS_BUENOS = 3
Public Const NPCDEFENSA = 4
Public Const GUARDIAS_ATACAN_CRIMINALES = 5
Public Const SIGUE_AMO = 8
Public Const NPC_ATACA_NPC = 9
Public Const NPC_PATHFINDING = 10
Public Const GUARDIAS_ATACAN_CIUDADANOS = 11

Private Sub GuardiasAI(ByVal NpcIndex As Integer)
 On Error GoTo fallo
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer

  'pluto:2.15
   'If MapInfo(Npclist(NpcIndex).Pos.Map).Dueño = 2 Then
   
  ' End If


For HeadingLoop = NORTH To WEST
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(HeadingLoop, nPos)
    If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
        UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
        If UI > 0 Then
              If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Privilegios = 0 Then
                     If Criminal(UI) Then
                            Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop, 0)
                            Call NpcAtacaUser(NpcIndex, UI)
                            Exit Sub
                     ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).Name _
                               And Not Npclist(NpcIndex).flags.Follow Then
                           Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop, 0)
                           Call NpcAtacaUser(NpcIndex, UI)
                           Exit Sub
                     End If
              End If
        End If
    End If
Next HeadingLoop

Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("GUARDIASAI " & Err.number & " D: " & Err.Description)

End Sub
Private Sub GuardiasAIcaos(ByVal NpcIndex As Integer)
 On Error GoTo fallo
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer

For HeadingLoop = NORTH To WEST
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(HeadingLoop, nPos)
    If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
        UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
        If UI > 0 Then
              If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Privilegios = 0 Then
                     If Not Criminal(UI) Then
                            Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop, 0)
                            Call NpcAtacaUser(NpcIndex, UI)
                            Exit Sub
                     ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).Name _
                               And Not Npclist(NpcIndex).flags.Follow Then
                           Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop, 0)
                           Call NpcAtacaUser(NpcIndex, UI)
                           Exit Sub
                     End If
              End If
        End If
    End If
Next HeadingLoop

Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("GUARDIASAICAOS " & Err.number & " D: " & Err.Description)

End Sub
Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
 On Error GoTo fallo
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
For HeadingLoop = NORTH To WEST
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(HeadingLoop, nPos)
    If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
        UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
        If UI > 0 Then
            If UserList(UI).flags.Muerto = 0 Then
'pluto:2.4
 If UserList(UI).flags.Privilegios > 0 Or UserList(UI).flags.Incor = True Then GoTo aqui2
                
                '¿ES del clan del castillo1?
                'pluto:2.11
                Set UserList(UI).GuildRef = FetchGuild(UserList(UI).GuildInfo.GuildName)
            If Not UserList(UI).GuildRef Is Nothing Then
                If Npclist(NpcIndex).Pos.Map = mapa_castillo1 And UserList(UI).GuildRef.IsAllie(castillo1) Then GoTo aqui2
                If Npclist(NpcIndex).Pos.Map = mapa_castillo2 And UserList(UI).GuildRef.IsAllie(castillo2) Then GoTo aqui2
                If Npclist(NpcIndex).Pos.Map = mapa_castillo3 And UserList(UI).GuildRef.IsAllie(castillo3) Then GoTo aqui2
                If Npclist(NpcIndex).Pos.Map = mapa_castillo4 And UserList(UI).GuildRef.IsAllie(castillo4) Then GoTo aqui2
                If Npclist(NpcIndex).Pos.Map = 185 And UserList(UI).GuildRef.IsAllie(fortaleza) Then GoTo aqui2
            End If
             '-----------------------
                
                If Npclist(NpcIndex).Pos.Map = mapa_castillo1 And UserList(UI).GuildInfo.GuildName = castillo1 Then GoTo aqui2
                If Npclist(NpcIndex).Pos.Map = mapa_castillo2 And UserList(UI).GuildInfo.GuildName = castillo2 Then GoTo aqui2
                If Npclist(NpcIndex).Pos.Map = mapa_castillo3 And UserList(UI).GuildInfo.GuildName = castillo3 Then GoTo aqui2
                If Npclist(NpcIndex).Pos.Map = mapa_castillo4 And UserList(UI).GuildInfo.GuildName = castillo4 Then GoTo aqui2
If Npclist(NpcIndex).Pos.Map = 185 And UserList(UI).GuildInfo.GuildName = fortaleza Then GoTo aqui2

                    If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then
                    Dim k As Integer
                    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                    Call NpcLanzaUnSpell(NpcIndex, UI)
                    End If
                    'pluto:6.0A
                    If Npclist(NpcIndex).Arquero > 0 Then
                    If Porcentaje(100, Npclist(NpcIndex).Arquero * 10) Then
                    Call NpcAtacaUser(NpcIndex, UI)
                    Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop, 0)
                    Exit Sub
                    End If
                    End If
                    '-------------
                Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop, 0)
             
                Call NpcAtacaUser(NpcIndex, MapData(nPos.Map, nPos.X, nPos.Y).UserIndex)
                Exit Sub
            End If
        
        End If
        End If
aqui2:
Next HeadingLoop

Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("HOSTILMALVADOAI " & Err.number & " D: " & Err.Description)

End Sub


Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
 On Error GoTo fallo
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
For HeadingLoop = NORTH To WEST
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(HeadingLoop, nPos)
    If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
        UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
        If UI > 0 Then
            If UserList(UI).Name = Npclist(NpcIndex).flags.AttackedBy Then
                If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Privilegios = 0 Then
                        If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                          Dim k As Integer
                          k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                          Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                        Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop, 0)
                        Call NpcAtacaUser(NpcIndex, UI)
                        Exit Sub
                End If
            End If
        End If
    End If
Next HeadingLoop

Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("HOSTILBUENOAI " & Err.number & " D: " & Err.Description)

End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
 On Error GoTo fallo
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
    For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
               UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
               If UI > 0 Then

                  
                  If UserList(UI).flags.Muerto = 1 Or UserList(UI).flags.Incor = True Then GoTo aqui
'pluto:2.4
 If UserList(UI).flags.Privilegios > 0 Then GoTo aqui
   'pluto:6.0A
   If UserList(UI).raza = "Vampiro" And (UserList(UI).Char.Body = 9 Or UserList(UI).Char.Body = 260) Then GoTo aqui
  If UserList(UI).flags.Minotauro > 0 And UserList(UI).Char.Body = 380 Then GoTo aqui
            
                  'castillo clan
                
                'pluto:2.11 aliados no ataca rey
            Set UserList(UI).GuildRef = FetchGuild(UserList(UI).GuildInfo.GuildName)
            If Not UserList(UI).GuildRef Is Nothing Then
                If Npclist(NpcIndex).Pos.Map = mapa_castillo1 And UserList(UI).GuildRef.IsAllie(castillo1) Then GoTo aqui
                If Npclist(NpcIndex).Pos.Map = mapa_castillo2 And UserList(UI).GuildRef.IsAllie(castillo2) Then GoTo aqui
                If Npclist(NpcIndex).Pos.Map = mapa_castillo3 And UserList(UI).GuildRef.IsAllie(castillo3) Then GoTo aqui
                If Npclist(NpcIndex).Pos.Map = mapa_castillo4 And UserList(UI).GuildRef.IsAllie(castillo4) Then GoTo aqui
                If Npclist(NpcIndex).Pos.Map = 185 And UserList(UI).GuildRef.IsAllie(fortaleza) Then GoTo aqui
            End If
            'End If
   
   
                If Npclist(NpcIndex).Pos.Map = mapa_castillo1 And (UserList(UI).GuildInfo.GuildName = castillo1) Then GoTo aqui
                If Npclist(NpcIndex).Pos.Map = mapa_castillo2 And (UserList(UI).GuildInfo.GuildName = castillo2) Then GoTo aqui
                If Npclist(NpcIndex).Pos.Map = mapa_castillo3 And (UserList(UI).GuildInfo.GuildName = castillo3) Then GoTo aqui
                If Npclist(NpcIndex).Pos.Map = mapa_castillo4 And (UserList(UI).GuildInfo.GuildName = castillo4) Then GoTo aqui
                If Npclist(NpcIndex).Pos.Map = 185 And UserList(UI).GuildInfo.GuildName = fortaleza Then GoTo aqui

                If UserList(UI).flags.Invisible = 1 And Npclist(NpcIndex).flags.Magiainvisible = 0 Then GoTo aqui
                If UserList(UI).flags.AdminInvisible = 1 Then GoTo aqui
                   If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                   'pluto:6.0A
                   If Npclist(NpcIndex).Arquero > 0 Then Call NpcAtacaUser(NpcIndex, UI)
                   '---------
                   'pluto:2.4.1
                   If UserList(UI).flags.Muerto = 0 And Npclist(NpcIndex).flags.Paralizado = 0 Then
                   
                           tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                       
                       'pluto:2.14
If Npclist(NpcIndex).flags.PoderEspecial3 > 0 Then
If Npclist(NpcIndex).Char.Body <> UserList(UI).Char.Body Or Npclist(NpcIndex).Char.Head <> UserList(UI).Char.Head Then
Npclist(NpcIndex).Char.Body = UserList(UI).Char.Body
Npclist(NpcIndex).Char.Head = UserList(UI).Char.Head
Call ChangeNPCChar(ToMap, 0, Npclist(NpcIndex).Pos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, tHeading, 0)
End If
End If
'----------------------------------------------
                       Call MoveNPCChar(NpcIndex, tHeading)
                       Exit Sub
                  End If
               End If
        End If
aqui:
    Next X
Next Y

Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("IRUSUARIOCERCANO " & Err.number & " D: " & Err.Description)

End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
On Error GoTo fallo
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer

'pluto:2.22-----------------------------------------------------
'If Npclist(NpcIndex).Name = "NPC SIN INICIAR" Then
'Npclist(NpcIndex).flags.NPCActive = False
'Exit Sub
'End If
'-------------------------------------------------------------
 
'pluto:6.0A MIRAR ESTO
If Npclist(NpcIndex).Pos.Map = 0 Then Exit Sub

'nati:agrego NPCType = 33
If Npclist(NpcIndex).NPCtype = 78 Or Npclist(NpcIndex).NPCtype = 33 Or Npclist(NpcIndex).NPCtype = 61 Then Exit Sub
For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
    For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
            UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
            If UI > 0 Then
            'pluto:6.5--------------
            'If UserList(UI).flags.Privilegios > 0 Then Exit Sub
            '-----------------------------
                If UserList(UI).Name = Npclist(NpcIndex).flags.AttackedBy Then
                    
                    If Npclist(NpcIndex).flags.LanzaSpells > 0 Or Npclist(NpcIndex).Raid > 0 Then
                              Dim k As Integer
                              k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                         End If
                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.AdminInvisible = 0 Then
                         
                         tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                         Call MoveNPCChar(NpcIndex, tHeading)
                         'pluto:2.4.5
                         If Distancia(Npclist(NpcIndex).Pos, UserList(UI).Pos) < 2 Then
                         Call NpcAtacaUser(NpcIndex, UI)
                         End If
                         
                         Exit Sub
                    End If
                End If
            End If
        End If
    Next X
Next Y

Call RestoreOldMovement(NpcIndex)


Exit Sub
fallo:
Call LogError("SEGUIRAGRESOR " & UserList(UI).Name & " -> nº: " & NpcIndex & " nom: " & Npclist(NpcIndex).Name & " D: " & Err.Description)

End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
On Error GoTo fallo
'pluto:6.5 añado raids
If Npclist(NpcIndex).MaestroUser = 0 And Npclist(NpcIndex).Raid = 0 Then
    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    Npclist(NpcIndex).flags.AttackedBy = ""
End If
Exit Sub
fallo:
Call LogError("RESTOREOLDMOVEMENT " & Err.number & " D: " & Err.Description)


End Sub


Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
 On Error GoTo fallo
Dim UI As Integer
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
    For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
           UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
           If UI > 0 Then
                If Criminal(UI) And UserList(UI).flags.Privilegios = 0 Then
                   If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Dim k As Integer
                              k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                   If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.AdminInvisible = 0 And UserList(UI).flags.Privilegios = 0 Then
                        
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
           End If
        End If
    Next X
Next Y

Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("PERSIGUECRIMINAL " & Err.number & " D: " & Err.Description)

End Sub
Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
 On Error GoTo fallo
Dim UI As Integer
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
    For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
           UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
           If UI > 0 Then
                If Not Criminal(UI) And UserList(UI).flags.Privilegios = 0 Then
                   If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Dim k As Integer
                              k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                   If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.AdminInvisible = 0 And UserList(UI).flags.Privilegios = 0 Then
                        
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
           End If
        End If
    Next X
Next Y

Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("PERSIGUECIUDADANO " & Err.number & " D: " & Err.Description)

End Sub
Private Sub SeguirAmo(ByVal NpcIndex As Integer)
On Error GoTo fallo
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
    For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
            If Npclist(NpcIndex).Target = 0 And Npclist(NpcIndex).TargetNpc = 0 Then
                UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
                If UI > 0 Then
                   If UserList(UI).flags.Muerto = 0 _
                   And UserList(UI).flags.Invisible = 0 _
                   And UserList(UI).flags.AdminInvisible = 0 _
                   And UI = Npclist(NpcIndex).MaestroUser _
                   And Distancia(Npclist(NpcIndex).Pos, UserList(UI).Pos) > 3 Then
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
            End If
        End If
    Next X
Next Y

Call RestoreOldMovement(NpcIndex)
Exit Sub
fallo:
Call LogError("SEGUIRAMO " & Err.number & " D: " & Err.Description)

End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
 On Error GoTo fallo
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim NI As Integer
Dim bNoEsta As Boolean
For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
    For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
           NI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).NpcIndex
           If NI > 0 Then
                If Npclist(NpcIndex).TargetNpc = NI Then
                     bNoEsta = True
                     tHeading = FindDirection(Npclist(NpcIndex).Pos, Npclist(MapData(Npclist(NpcIndex).Pos.Map, X, Y).NpcIndex).Pos)
                     Call MoveNPCChar(NpcIndex, tHeading)
                     'pluto:2.4.5
                     If Distancia(Npclist(NpcIndex).Pos, Npclist(NI).Pos) < 2 Then
                               Call NpcAtacaNpc(NpcIndex, NI)
                                 'pluto:6.0A-----
                     'Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, tHeading, 1)
                    '--------------
        
                     End If
                     
                     Exit Sub
                End If
           End If
           
        End If
    Next X
Next Y

If Not bNoEsta Then
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call FollowAmo(NpcIndex)
    Else
        Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
        Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    End If
End If
   Exit Sub
fallo:
Call LogError("AINPCATACANPC " & Err.number & " D: " & Err.Description)
 
End Sub
Sub HablaPirata(NpcIndex)
Dim n As Integer
n = RandomNumber(1, 300)
If n > 8 Then Exit Sub
Call SendData(ToNPCArea, val(NpcIndex), Npclist(NpcIndex).Pos.Map, "!;8°" & n & "°" & Npclist(NpcIndex).Char.CharIndex)
End Sub
Function NPCAI(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
 'pluto:2.22-----------------------
'If Npclist(NpcIndex).flags.NPCActive = False Then
'Dim MiNPC As npc
'MiNPC = Npclist(NpcIndex)
'Call QuitarNPC(NpcIndex)
'Call ReSpawnNpc(MiNPC)
'Exit Function
'End If
'----------------------------------
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If Npclist(NpcIndex).MaestroUser = 0 Then
            'Busca a alguien para atacar
            '¿Es un guardia?
            If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                    Call GuardiasAI(NpcIndex)
            ElseIf Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS2 Then
                    Call GuardiasAIcaos(NpcIndex)
            ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion <> 0 Then
                    Call HostilMalvadoAI(NpcIndex)
            ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Call HostilBuenoAI(NpcIndex)
           
            End If
        Else
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
   
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
           'quitar esto
       ' If Npclist(NpcIndex).TargetNpc > 0 Then Npclist(NpcIndex).Movement = 9
     
        Select Case Npclist(NpcIndex).Movement
        
            Case MUEVE_AL_AZAR
                If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                    If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                    End If
                    Call PersigueCriminal(NpcIndex)
                ElseIf Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS2 Then
                    If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                    End If
                    Call PersigueCiudadano(NpcIndex)
                Else
                    If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                    End If
                End If
            'Va hacia el usuario cercano
            Case NPC_MALO_ATACA_USUARIOS_BUENOS
                Call IrUsuarioCercano(NpcIndex)
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case NPCDEFENSA
                Call SeguirAgresor(NpcIndex)
            'Persigue criminales
            Case GUARDIAS_ATACAN_CRIMINALES
                Call PersigueCriminal(NpcIndex)
            'Persigue CIUDAS
            Case GUARDIAS_ATACAN_CIUDADANOS
                Call PersigueCiudadano(NpcIndex)
            Case SIGUE_AMO
                Call SeguirAmo(NpcIndex)
                If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                End If
            Case NPC_ATACA_NPC
                Call AiNpcAtacaNpc(NpcIndex)
            Case NPC_PATHFINDING
                
                If ReCalculatePath(NpcIndex) Then
                    Call PathFindingAI(NpcIndex)
                    'Existe el camino?
                    If Npclist(NpcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        Call MoveNPCChar(NpcIndex, Int(RandomNumber(1, 4)))
                    End If
                Else
                    If Not PathEnd(NpcIndex) Then
                        Call FollowPath(NpcIndex)
                    Else
                        Npclist(NpcIndex).PFINFO.PathLenght = 0
                    End If
                End If

        End Select


Exit Function


ErrorHandler:
    'pluto:2.4.5
    Call LogError("NPCAI " & Npclist(NpcIndex).Name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).Pos.Map & " x:" & Npclist(NpcIndex).Pos.X & " y:" & Npclist(NpcIndex).Pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN: " & Npclist(NpcIndex).TargetNpc & " N: " & Err.number & " D: " & Err.Description)
    Dim MinPc As npc
    MinPc = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MinPc)
    
End Function


Function UserNear(ByVal NpcIndex As Integer) As Boolean
On Error GoTo fallo
'#################################################################
'Returns True if there is an user adjacent to the npc position.
'#################################################################
UserNear = Not Int(Distance(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.X, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.Y)) > 1
Exit Function
fallo:
Call LogError("USERNEAR " & Err.number & " D: " & Err.Description)

End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
 On Error GoTo fallo
'#################################################################
'Returns true if we have to seek a new path
'#################################################################
If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
    ReCalculatePath = True
ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
    ReCalculatePath = True
End If
Exit Function
fallo:
Call LogError("RECALCULATEPATH " & Err.number & " D: " & Err.Description)

End Function

Function SimpleAI(ByVal NpcIndex As Integer) As Boolean
On Error GoTo fallo
'#################################################################
'Old Ore4 AI function
'#################################################################
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer

For Y = Npclist(NpcIndex).Pos.Y - 5 To Npclist(NpcIndex).Pos.Y + 5    'Makes a loop that looks at
    For X = Npclist(NpcIndex).Pos.X - 5 To Npclist(NpcIndex).Pos.X + 5   '5 tiles in every direction
           'Make sure tile is legal
            If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                'look for a user
                If MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex > 0 Then
                    'Move towards user
                    tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).Pos)
                    MoveNPCChar NpcIndex, tHeading
                    'Leave
                    Exit Function
                End If
            End If
    Next X
Next Y
Exit Function
fallo:
Call LogError("SIMPLEAI " & Err.number & " D: " & Err.Description)

End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
 On Error GoTo fallo
'#################################################################
'Coded By Gulfas Morgolock
'Returns if the npc has arrived to the end of its path
'#################################################################
PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght

Exit Function
fallo:
Call LogError("PATHEND " & Err.number & " D: " & Err.Description)


End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
 On Error GoTo fallo
'#################################################################
'Coded By Gulfas Morgolock
'Moves the npc.
'#################################################################

Dim tmpPos As WorldPos
Dim tHeading As Byte

tmpPos.Map = Npclist(NpcIndex).Pos.Map
tmpPos.X = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).Y ' invertí las coordenadas
tmpPos.Y = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).X

'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"

tHeading = FindDirection(Npclist(NpcIndex).Pos, tmpPos)

MoveNPCChar NpcIndex, tHeading

Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.CurPos + 1
Exit Function
fallo:
Call LogError("FOLLOWPATH " & Err.number & " D: " & Err.Description)

End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
 On Error GoTo fallo
'#################################################################
'Coded By Gulfas Morgolock / 11-07-02
'www.geocities.com/gmorgolock
'morgolock@speedy.com.ar
'This function seeks the shortest path from the Npc
'to the user's location.
'#################################################################
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer

For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10    'Makes a loop that looks at
     For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10   '5 tiles in every direction

         'Make sure tile is legal
         If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
         
             'look for a user
             If MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex > 0 Then
                 'pluto:2.11
                If UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).flags.Privilegios > 0 Or UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex).flags.Muerto > 0 Then GoTo yop
                 
                 'Move towards user
                  Dim tmpUserIndex As Integer
                  tmpUserIndex = MapData(Npclist(NpcIndex).Pos.Map, X, Y).UserIndex
                  
                  'We have to invert the coordinates, this is because
                  'ORE refers to maps in converse way of my pathfinding
                  'routines.
                  Npclist(NpcIndex).PFINFO.Target.X = UserList(tmpUserIndex).Pos.Y
                  Npclist(NpcIndex).PFINFO.Target.Y = UserList(tmpUserIndex).Pos.X 'ops!
                  Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                                      'pluto:2.10
                   

                   If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, tmpUserIndex)
                   

                  Call SeekPath(NpcIndex)
                  Exit Function
             End If
             
         End If
yop:
     Next X
 Next Y

Exit Function
fallo:
Call LogError("PATHFINDINGAI " & Err.number & " D: " & Err.Description)


End Function


Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
 On Error GoTo fallo
'pluto:2.8.0
Dim pro As Byte

'PLUTO:6.0a
'If Npclist(NpcIndex).Raid > 0 Then
    'If Npclist(NpcIndex).Stats.MinHP < 7000 Then
'If RandomNumber(1, 1000) < Npclist(NpcIndex).Raid * 10 Then
'Call SendData2(ToMap, 0, Npclist(NpcIndex).Pos.Map, 22, Npclist(NpcIndex).Char.CharIndex & "," & 31 & "," & 1)
'Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP + Npclist(NpcIndex).Raid * 50
'Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & 18)
'If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
'Call SendData(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "H4" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Stats.MinHP)
'Exit Sub ' si se cura no lanza otra cosa
'End If
   ' End If
'End If 'RAID

'pluto:6.0A


If UserList(UserIndex).raza = "Vampiro" And (UserList(UserIndex).Char.Body = 9 Or UserList(UserIndex).Char.Body = 260) Then Exit Sub
If UserList(UserIndex).flags.Minotauro > 0 And UserList(UserIndex).Char.Body = 380 Then Exit Sub

'pluto:6.5 añado privilegio>0
If UserList(UserIndex).flags.AdminInvisible = 1 Or UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub

If UserList(UserIndex).flags.Invisible = 1 And Npclist(NpcIndex).flags.Magiainvisible = 0 Then Exit Sub
Dim k As Integer
k = RandomNumber(1, 10)
'pluto:2.10
If Npclist(NpcIndex).Movement = 10 Then pro = 8 Else pro = 4

If k > pro Then Exit Sub
'pluto:6.0A
If Npclist(NpcIndex).flags.LanzaSpells = 0 Then Exit Sub

k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))
Exit Sub
fallo:
Call LogError("NPCLANZAUNSPELL " & Npclist(NpcIndex).Name & "-->" & UserList(UserIndex).Name & " " & Err.number & " D: " & Err.Description)

End Sub

