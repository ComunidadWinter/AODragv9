Attribute VB_Name = "Drag_AccionesRaton"
Option Explicit


Sub ClickIzquierdo(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
  '***************************************************
  'Autor: Unknown (orginal version)
  'Last Modification: 26/03/2009
  '13/02/2009: ZaMa - EL nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
  '***************************************************
  
  On Error GoTo Errhandler
  
  'Responde al click del usuario sobre el mapa
  Dim FoundChar As Byte
  Dim FoundSomething As Byte
  Dim TempCharIndex As Integer
  Dim Stat As String
  Dim ft As FontTypeNames
  
  '¿Rango Visión? (ToxicWaste)
  If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
      Exit Sub
  End If
  
  '¿Posicion valida?
  If InMapBounds(Map, X, Y) Then
      UserList(UserIndex).flags.TargetMap = Map
      UserList(UserIndex).flags.TargetX = X
      UserList(UserIndex).flags.TargetY = Y
      '¿Es un obj?
      If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
          'Informa el nombre
          UserList(UserIndex).flags.TargetObjMap = Map
          UserList(UserIndex).flags.TargetObjX = X
          UserList(UserIndex).flags.TargetObjY = Y
          FoundSomething = 1
      ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
          'Informa el nombre
          If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
              UserList(UserIndex).flags.TargetObjMap = Map
              UserList(UserIndex).flags.TargetObjX = X + 1
              UserList(UserIndex).flags.TargetObjY = Y
              FoundSomething = 1
          End If
      ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
          If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
              'Informa el nombre
              UserList(UserIndex).flags.TargetObjMap = Map
              UserList(UserIndex).flags.TargetObjX = X + 1
              UserList(UserIndex).flags.TargetObjY = Y + 1
              FoundSomething = 1
          End If
      ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
          If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
              'Informa el nombre
              UserList(UserIndex).flags.TargetObjMap = Map
              UserList(UserIndex).flags.TargetObjX = X
              UserList(UserIndex).flags.TargetObjY = Y + 1
              FoundSomething = 1
          End If
      End If
      
      If FoundSomething = 1 Then
          UserList(UserIndex).flags.targetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
          If ObjData(UserList(UserIndex).flags.targetObj).DefensaMagicaMax > 0 Or ObjData(UserList(UserIndex).flags.targetObj).DañoMagico > 0 Then
              Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.targetObj).Name & " - " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.Amount & " || DañoMagico: " & ObjData(UserList(UserIndex).flags.targetObj).DañoMagico & " - DefensaMagica: Max " & ObjData(UserList(UserIndex).flags.targetObj).DefensaMagicaMax & "\Min " & ObjData(UserList(UserIndex).flags.targetObj).DefensaMagicaMin, FontTypeNames.FONTTYPE_INFO)
          Else
              Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.targetObj).Name & " - " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.Amount & "", FontTypeNames.FONTTYPE_INFO)
          End If
      End If
      '¿Es un personaje?
      If Y + 1 <= YMaxMapSize Then
          If MapData(Map, X, Y + 1).UserIndex > 0 Then
              TempCharIndex = MapData(Map, X, Y + 1).UserIndex
              FoundChar = 1
          End If
          If MapData(Map, X, Y + 1).NPCIndex > 0 Then
              TempCharIndex = MapData(Map, X, Y + 1).NPCIndex
              FoundChar = 2
          End If
      End If
      '¿Es un personaje?
      If FoundChar = 0 Then
          If MapData(Map, X, Y).UserIndex > 0 Then
              TempCharIndex = MapData(Map, X, Y).UserIndex
              FoundChar = 1
          End If
          If MapData(Map, X, Y).NPCIndex > 0 Then
              TempCharIndex = MapData(Map, X, Y).NPCIndex
              FoundChar = 2
          End If
      End If
      
      
      'Reaccion al personaje
      If FoundChar = 1 Then '  ¿Encontro un Usuario?
              
         If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios And PlayerType.Dios Then
              
              If LenB(UserList(TempCharIndex).DescRM) = 0 And UserList(TempCharIndex).showName Then 'No tiene descRM y quiere que se vea su nombre.
                  
                  If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                      Stat = Stat & " <Ejército Real> " & "<" & TituloReal(TempCharIndex) & ">"
                  ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                      Stat = Stat & " <Legión Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"
                  End If
                  
                  If UserList(TempCharIndex).GuildIndex > 0 Then
                      Stat = Stat & " <" & modGuilds.GuildName(UserList(TempCharIndex).GuildIndex) & ">"
                  End If
                  
                  If Len(UserList(TempCharIndex).desc) > 0 Then
                      Stat = "Ves a " & UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).desc
                  Else
                      Stat = "Ves a " & UserList(TempCharIndex).Name & Stat
                  End If
                  
                                  
                  If UserList(TempCharIndex).flags.Privilegios And PlayerType.RoyalCouncil Then
                      Stat = Stat & " [CONSEJO DE BANDERBILL]"
                      ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                  ElseIf UserList(TempCharIndex).flags.Privilegios And PlayerType.ChaosCouncil Then
                      Stat = Stat & " [CONCILIO DE LAS SOMBRAS]"
                      ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                  Else
                      If Not UserList(TempCharIndex).flags.Privilegios And PlayerType.User Then
                          Stat = Stat & " <GAME MASTER>"
                          
                          ' Elijo el color segun el rango del GM:
                          ' Dios
                          If UserList(TempCharIndex).flags.Privilegios = PlayerType.Dios Then
                              ft = FontTypeNames.FONTTYPE_DIOS
                          ' Gm
                          ElseIf UserList(TempCharIndex).flags.Privilegios = PlayerType.SemiDios Then
                              ft = FontTypeNames.FONTTYPE_GM
                          ' Conse
                          ElseIf UserList(TempCharIndex).flags.Privilegios = PlayerType.Consejero Then
                              ft = FontTypeNames.FONTTYPE_CONSE
                          ' Rm o Dsrm
                          ElseIf UserList(TempCharIndex).flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Consejero) Or UserList(TempCharIndex).flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Dios) Then
                              ft = FontTypeNames.FONTTYPE_EJECUCION
                          End If
                      ElseIf EsNewbie(TempCharIndex) = True Then
                          Stat = Stat & " <NEWBIE>"
                          ft = FontTypeNames.FONTTYPE_DIOS
                      ElseIf criminal(TempCharIndex) Then
                          Stat = Stat & " <CRIMINAL>"
                          ft = FontTypeNames.FONTTYPE_FIGHT
                      Else
                          Stat = Stat & " <CIUDADANO>"
                          ft = FontTypeNames.FONTTYPE_CITIZEN
                      End If
                  End If
              Else  'Si tiene descRM la muestro siempre.
                  Stat = UserList(TempCharIndex).DescRM
                  ft = FontTypeNames.FONTTYPE_INFOBOLD
              End If
              
              If UserList(TempCharIndex).flags.Muerto = 1 Then Stat = Stat & " <MUERTO>"
              
              If GranPoder = TempCharIndex Then Stat = Stat & " [Bendecido con el Gran Poder]"
              
              'If UserList(UserIndex).flags.Privilegios Then Stat = Stat & "[PartyId: " & UserList(TempCharIndex).PartyId & "]"
              
              
              If LenB(Stat) > 0 Then
                  Call WriteConsoleMsg(UserIndex, Stat, ft)
              End If
              
              FoundSomething = 1
              UserList(UserIndex).flags.targetUser = TempCharIndex
              UserList(UserIndex).flags.TargetNPC = 0
              UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
         End If
  
      End If
      If FoundChar = 2 Then '¿Encontro un NPC?
              Dim estatus As String
              
              If NPCList(TempCharIndex).Nivel = 0 Then
                  estatus = " Nivel: ?? -"
              Else
                  estatus = " Nivel: " & NPCList(TempCharIndex).Nivel & " -"
              End If
              
              estatus = estatus + " [" & NPCList(TempCharIndex).Stats.MinHP & "/" & NPCList(TempCharIndex).Stats.MaxHP & " -"
              
              If NPCList(TempCharIndex).Stats.MinHP < (NPCList(TempCharIndex).Stats.MaxHP * 0.05) Then
                  estatus = estatus + " Agonizando] "
              ElseIf NPCList(TempCharIndex).Stats.MinHP < (NPCList(TempCharIndex).Stats.MaxHP * 0.1) Then
                  estatus = estatus + " Casi muerto] "
              ElseIf NPCList(TempCharIndex).Stats.MinHP < (NPCList(TempCharIndex).Stats.MaxHP * 0.25) Then
                  estatus = estatus + " Muy Malherido] "
              ElseIf NPCList(TempCharIndex).Stats.MinHP < (NPCList(TempCharIndex).Stats.MaxHP * 0.5) Then
                  estatus = estatus + " Herido] "
              ElseIf NPCList(TempCharIndex).Stats.MinHP < (NPCList(TempCharIndex).Stats.MaxHP * 0.75) Then
                  estatus = estatus + " Levemente herido] "
              ElseIf NPCList(TempCharIndex).Stats.MinHP < (NPCList(TempCharIndex).Stats.MaxHP) Then
                  estatus = estatus + " Sano] "
              Else
                  estatus = estatus + " Intacto] "
              End If
              
              If NPCList(TempCharIndex).flags.Paralizado = 1 Then
                  estatus = estatus + "[Paralizado]"
              ElseIf NPCList(TempCharIndex).flags.Inmovilizado = 1 Then
                  estatus = estatus + "[Inmovilizado]"
              End If
              
              Dim CentinelaIndex As Integer
              CentinelaIndex = EsCentinela(TempCharIndex)
              
              If Len(NPCList(TempCharIndex).desc) > 1 Then
                  
                  '26/02/2016 Irongete: Al hacer click al Protector de la Mina de la Fortaleza si eres el dueño de la misma eres transportado dentro.
                  If NPCList(TempCharIndex).Numero = 604 Then _
                      If UserList(UserIndex).GuildIndex = Castillos(5).Dueño Then Call WarpUserChar(UserIndex, 43, 80, 71, True)
                  
                  Call WriteChatOverHead(UserIndex, NPCList(TempCharIndex).desc, NPCList(TempCharIndex).Char.CharIndex, vbWhite)
              ElseIf CentinelaIndex <> 0 Then
                  'Enviamos nuevamente el texto del centinela según quien pregunta
                  Call modCentinela.CentinelaSendClave(UserIndex, CentinelaIndex)
              Else
                  If NPCList(TempCharIndex).MaestroUser > 0 Then
                      Call WriteConsoleMsg(UserIndex, NPCList(TempCharIndex).Name & estatus & " es mascota de " & UserList(NPCList(TempCharIndex).MaestroUser).Name, FontTypeNames.FONTTYPE_INFO)
                  Else
                  
                  '18/11/2018 Lorwik: Si es un Arbol o un Yacimiento solo veremos su nombre
                If NPCList(TempCharIndex).NPCType = Arbol Or NPCList(TempCharIndex).NPCType = Yacimiento Then
                      Call WriteConsoleMsg(UserIndex, NPCList(TempCharIndex).Name, FontTypeNames.FONTTYPE_INFO)
                Else
                      '27/10/2015 Irongete: Oculto el mensaje con el nombre, la vida, el nivel y el estado de salud del NPC al hacer click encima de el
                      ' SOLAMENTE cuando se esté casteando una mágia sobre el
                      If UserList(UserIndex).flags.Hechizo = 0 Then
                          Call WriteConsoleMsg(UserIndex, NPCList(TempCharIndex).Name & " -" & estatus, FontTypeNames.FONTTYPE_INFO)
                      End If
                      
                      If UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                          Call WriteConsoleMsg(UserIndex, "Le pegó primero: " & NPCList(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                      End If
                  End If
                End If
                  
              End If
              FoundSomething = 1
              UserList(UserIndex).flags.TargetNpcTipo = NPCList(TempCharIndex).NPCType
              UserList(UserIndex).flags.TargetNPC = TempCharIndex
              UserList(UserIndex).flags.targetUser = 0
              UserList(UserIndex).flags.targetObj = 0
          
      End If
      
      If FoundChar = 0 Then
          UserList(UserIndex).flags.TargetNPC = 0
          UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
          UserList(UserIndex).flags.targetUser = 0
      End If
      
      '*** NO ENCOTRO NADA ***
      If FoundSomething = 0 Then
          UserList(UserIndex).flags.TargetNPC = 0
          UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
          UserList(UserIndex).flags.targetUser = 0
          UserList(UserIndex).flags.targetObj = 0
          UserList(UserIndex).flags.TargetObjMap = 0
          UserList(UserIndex).flags.TargetObjX = 0
          UserList(UserIndex).flags.TargetObjY = 0
      End If
  
  Else
      If FoundSomething = 0 Then
          UserList(UserIndex).flags.TargetNPC = 0
          UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
          UserList(UserIndex).flags.targetUser = 0
          UserList(UserIndex).flags.targetObj = 0
          UserList(UserIndex).flags.TargetObjMap = 0
          UserList(UserIndex).flags.TargetObjX = 0
          UserList(UserIndex).flags.TargetObjY = 0
      End If
  End If
  
  Exit Sub
  
Errhandler:
      Call LogError("Error en LookAtTile. Error " & Err.Number & " : " & Err.Description)

End Sub

Sub ClickDerecho(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
  On Error Resume Next
  Dim tempIndex As Integer
  
  '¿Rango Visión? (ToxicWaste)
  If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
    Exit Sub
  End If
  
  '¿Está trabajando?
  If UserList(UserIndex).flags.Makro <> 0 Then
    Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
  End If
  
  '¿Posicion valida?
  If InMapBounds(Map, X, Y) Then
    With UserList(UserIndex)
      If MapData(Map, X, Y).NPCIndex > 0 Then     'Acciones NPCs
        tempIndex = MapData(Map, X, Y).NPCIndex
        
        'Set the target NPC
        .flags.TargetNPC = tempIndex
        
        If NPCList(tempIndex).NPCType = eNPCType.Quest Then
        
          Debug.Print "CLICK EN QUEST"; NPCList(tempIndex).Numero
        
        
        ElseIf NPCList(tempIndex).Comercia = 1 Then
          '¿Esta el user muerto? Si es asi no puede comerciar
          If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
          End If
          
          'Is it already in commerce mode??
          If .flags.Comerciando Then
            Exit Sub
          End If
          
          If Distancia(NPCList(tempIndex).Pos, .Pos) > 3 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
          End If
          
          'Iniciamos la rutina pa' comerciar.
          Call IniciarComercioNPC(UserIndex)
          
        ElseIf NPCList(tempIndex).NPCType = eNPCType.Banquero Then
          '¿Esta el user muerto? Si es asi no puede comerciar
          If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
          End If
          
          'Is it already in commerce mode??
          If .flags.Comerciando Then
            Exit Sub
          End If
          
          If Distancia(NPCList(tempIndex).Pos, .Pos) > 3 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
          End If
          
          'A depositar de una
          Call IniciarDeposito(UserIndex)
          
        ElseIf NPCList(tempIndex).NPCType = eNPCType.Subastador Then
          '¿Esta el user muerto? Si es asi no puede comerciar
          If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
          End If
          
          'Is it already in commerce mode??
          If .flags.Comerciando Then
            Exit Sub
          End If
          
          If Distancia(NPCList(tempIndex).Pos, .Pos) > 3 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
          End If
          
          'A depositar de una
          Call IniciarSubasta(UserIndex)
          
        ElseIf NPCList(tempIndex).NPCType = eNPCType.Cirujano Then
          '¿Esta el user muerto? Si es asi no puede comerciar
          If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
          End If
          
          If .Stats.DragCredits < 1 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("¡No Creditos necesarios!. Mis honorarios es de 1 DragCredito.", NPCList(tempIndex).Char.CharIndex, vbWhite))
            Exit Sub
          End If
          
          Call CambiarCabeza(UserIndex)
          .Stats.DragCredits = .Stats.DragCredits - 1
          Call WriteUpdateUserStats(UserIndex)
          Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("¡Creo que es uno de mis mejores trabajos, espero que te guste!", NPCList(tempIndex).Char.CharIndex, vbWhite))
          Exit Sub
          
          
        ElseIf NPCList(tempIndex).NPCType = eNPCType.Guardiafalso Then
          Call WriteEnviaManual(UserIndex)
          
        ElseIf NPCList(tempIndex).NPCType = eNPCType.Revividor Or NPCList(tempIndex).NPCType = eNPCType.ResucitadorNewbie Then
          If Distancia(.Pos, NPCList(tempIndex).Pos) > 10 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
          End If
          
          'Revivimos si es necesario
          If .flags.Muerto = 1 And (NPCList(tempIndex).NPCType = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
            Call RevivirUsuario(UserIndex)
          End If
          
          If NPCList(tempIndex).NPCType = eNPCType.Revividor Or EsNewbie(UserIndex) Then
            'curamos veneno
            If iniSacerdoteCuraVeneno Then ' GSZAO
              If .flags.Envenenado <> 0 Then
                .flags.Envenenado = 0
              End If
            End If
            'curamos totalmente
            .Stats.MinHP = VidaMaxima(UserIndex)
            Call WriteUpdateUserStats(UserIndex)
          End If
        End If
        
      '¿Es un obj?
      ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
        tempIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
        
        .flags.targetObj = tempIndex
        
        Select Case ObjData(tempIndex).OBJType
        Case eOBJType.otPuertas 'Es una puerta
          Call AccionParaPuerta(Map, X, Y, UserIndex)
        Case eOBJType.otCarteles 'Es un cartel
          Call AccionParaCartel(Map, X, Y, UserIndex)
        Case eOBJType.otForos 'Foro
          Call AccionParaForo(Map, X, Y, UserIndex)
        Case eOBJType.otLeña    'Leña
          If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
            Call AccionParaRamita(Map, X, Y, UserIndex)
          End If
        End Select
        
        
        
      '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
      ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
        tempIndex = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
        .flags.targetObj = tempIndex
        
        Select Case ObjData(tempIndex).OBJType
          
        Case eOBJType.otPuertas 'Es una puerta
          Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
          
        End Select
        
      ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
        tempIndex = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex
        .flags.targetObj = tempIndex
        
        Select Case ObjData(tempIndex).OBJType
        Case eOBJType.otPuertas 'Es una puerta
          Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
        End Select
        
      ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
        tempIndex = MapData(Map, X, Y + 1).ObjInfo.ObjIndex
        .flags.targetObj = tempIndex
        
        Select Case ObjData(tempIndex).OBJType
        Case eOBJType.otPuertas 'Es una puerta
          Call AccionParaPuerta(Map, X, Y + 1, UserIndex)
        End Select
      End If
      
      
      
    End With
  End If
End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
  On Error Resume Next
  
  Dim Pos As WorldPos
  Pos.Map = Map
  Pos.X = X
  Pos.Y = Y
  
  If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call WriteMultiMessage(UserIndex, eMessages.Lejos)
    Exit Sub
  End If
  
  Call WriteShowForumForm(UserIndex)
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
  On Error Resume Next
  
  If MapData(Map, X, Y).ObjInfo.ObjIndex = 0 Then Exit Sub
  
  If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
      If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then
        'Abre la puerta
        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
          
          MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexAbierta
          
          Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
          
          'Desbloquea
          MapData(Map, X, Y).Blocked = 0
          MapData(Map, X - 1, Y).Blocked = 0
          
          'Bloquea todos los mapas
          Call Bloquear(True, Map, X, Y, 0)
          Call Bloquear(True, Map, X - 1, Y, 0)
          
          
          'Sonido
          Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
          
        Else
          Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
        End If
      Else
        'Cierra puerta
        MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexCerrada
        
        Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
        
        MapData(Map, X, Y).Blocked = 1
        MapData(Map, X - 1, Y).Blocked = 1
        
        
        Call Bloquear(True, Map, X - 1, Y, 1)
        Call Bloquear(True, Map, X, Y, 1)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
      End If
      
      UserList(UserIndex).flags.targetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
    Else
      Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
    End If
  Else
    Call WriteMultiMessage(UserIndex, eMessages.Lejos)
  End If
  
End Sub


Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
  On Error Resume Next
  
  If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
    
    If Len(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
      Call WriteShowSignal(UserIndex, MapData(Map, X, Y).ObjInfo.ObjIndex)
    End If
    
  End If
  
End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
  On Error Resume Next
  
  Dim Obj As Obj
  
  Dim Pos As WorldPos
  Pos.Map = Map
  Pos.X = X
  Pos.Y = Y
  
  If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call WriteMultiMessage(UserIndex, eMessages.Lejos)
    Exit Sub
  End If
  
  If MapData(Map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
    Call WriteConsoleMsg(UserIndex, "En zona segura no puedes hacer fogatas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
  End If
  
  If MapInfo(UserList(UserIndex).Pos.Map).Zona <> Ciudad Then
    Obj.ObjIndex = FOGATA
    Obj.Amount = 1
    
    Call WriteConsoleMsg(UserIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
    
    Call MakeObj(Obj, Map, X, Y)
    
    'Las fogatas prendidas se deben eliminar
    Call aLimpiarMundo.AddItem(Map, X, Y) ' GSZAO
  Else
    Call WriteConsoleMsg(UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
  End If
End Sub




