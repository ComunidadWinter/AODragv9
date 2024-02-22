Attribute VB_Name = "TCP1"
Sub TCP1(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo ErrorComandoPj:
Dim sndData As String
Dim CadenaOriginal As String
Dim xpa As Integer
Dim loopc As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim Tindex As Integer
Dim tName As String
Dim tNome As String
Dim tpru As String
Dim tMessage As String
Dim auxind As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim pass As String
Dim Mapa As Integer
Dim Name As String
Dim ind As Integer
Dim n As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim ClientCRC As String
Dim ServerSideCRC As Long

CadenaOriginal = rdata
If rdata = "" Then Exit Sub
Select Case UCase$(Left$(rdata, 1))
    
    Case ";" 'Hablar
'pluto:hoy

        If UserList(UserIndex).Char.FX > 38 And UserList(UserIndex).Char.FX < 67 Then
            Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
            UserList(UserIndex).Char.FX = 0
        End If
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 1)
        If InStr(rdata, "°") Then Exit Sub
        ind = UserList(UserIndex).Char.CharIndex
        
'pluto:7.0 bug cartel
        If LTrim(rdata) = "" Then
            Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 21, ind)
            'Call SendData(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "||1° °" & str(ind))
        Else
            If ((Not EsDios(UserList(UserIndex).Name)) And (Not EsSemiDios(UserList(UserIndex).Name))) Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||1°" & rdata & "°" & str(ind))
            Else
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||4°" & rdata & "°" & str(ind))
            End If
        End If
        
'PLUTO:HOY
        If UserList(UserIndex).flags.TargetNpc > 0 Then
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 15 And Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) < 8 Then
                If UCase$(rdata) = UCase$(ResTrivial) Then
                    Call SendData(ToPCArea, UserIndex, Npclist(UserList(UserIndex).flags.TargetNpc).Pos.Map, "||5°Muy bien " & UserList(UserIndex).Name & " la respuesta era " & ResTrivial & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
'pluto:2-3-04
                    Call SendData(ToIndex, UserIndex, 0, "||Has ganado 2 DragPuntos." & "´" & FontTypeNames.FONTTYPE_info)
                    UserList(UserIndex).Stats.Puntos = UserList(UserIndex).Stats.Puntos + 2
                    Call Loadtrivial
                End If
            End If
        End If
        Exit Sub
       
    Case "-" 'Gritar
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 1)
        If InStr(rdata, "°") Then
            Exit Sub
        End If
        ind = UserList(UserIndex).Char.CharIndex
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||2°" & rdata & "°" & str(ind))
        Exit Sub
    Case "\" 'Susurrar al oido
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 1)
        tName = ReadField(1, rdata, 58)
        'pluto:2.20
        'If ReadField(3, rdata, 32) <> "" Then
        'tName = ReadField(1, rdata, 32) & " " & ReadField(2, rdata, 32)
        'End If
       
        Tindex = NameIndex(tName & "$")
        If Tindex <> 0 Then
            If Len(rdata) <> Len(tName) Then
                tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
            Else
                tMessage = " "
            End If
            'pluto:2.4.5
            If UserList(Tindex).flags.Privilegios > 0 Then Exit Sub
            
            If Not EstaPCarea(UserIndex, Tindex) Then
                Call SendData(ToIndex, UserIndex, 0, "G9")
                Exit Sub
            End If
            ind = UserList(UserIndex).Char.CharIndex
            If InStr(tMessage, "°") Then
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||3°" & tMessage & "°" & str(ind))
            Call SendData(ToIndex, Tindex, UserList(UserIndex).Pos.Map, "||3°" & tMessage & "°" & str(ind))
            Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Usuario inexistente. " & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
  
    Case "*"
        Moverse = 1
        rdata = "1"
    Case "+"
        Moverse = 1
        rdata = "2"
    Case "="
        Moverse = 1
        rdata = "3"
    Case "M"
        Moverse = 1
        rdata = "4"
       
    Case "ª" 'Cambiar Heading ;-)
        rdata = Right$(rdata, Len(rdata) - 1)
        If val(rdata) > 0 And val(rdata) < 5 Then
            UserList(UserIndex).Char.Heading = rdata
            '[GAU] Agregamo UserList(UserIndex).Char.Botas
            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
        End If
        
    'Exit Sub
    'Case "M" 'Moverse
        'quitar esto
        'UserList(UserIndex).Flags.Privilegios = 3

End Select
 
 'pluto:2.17------------------------------------------
 If Moverse = 1 Then
    UserList(UserIndex).Counters.IdleCount = 0
    'PLUTO:6.3---------------
If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
End If
'--------------------------
    
    
    
    'rdata = Right$(rdata, Len(rdata) - 1)
    If Not UserList(UserIndex).flags.Descansar And Not UserList(UserIndex).flags.Meditando _
    And UserList(UserIndex).flags.Paralizado = 0 Then
        Call MoveUserChar(UserIndex, val(rdata))
    ElseIf UserList(UserIndex).flags.Descansar Then
        UserList(UserIndex).flags.Descansar = False
        Call SendData2(ToIndex, UserIndex, 0, 41)
        Call SendData(ToIndex, UserIndex, 0, "||Has dejado de descansar." & "´" & FontTypeNames.FONTTYPE_info)
        Call MoveUserChar(UserIndex, val(rdata))
'pluto:2.4 and paralizado=0
    ElseIf UserList(UserIndex).flags.Meditando And UserList(UserIndex).flags.Paralizado = 0 Then
        
        'UserList(userindex).Flags.Meditando = False
        'all SendData2(ToIndex, userindex, 0, 54)
        Call SendData(ToIndex, UserIndex, 0, "||Meditando!!" & "´" & FontTypeNames.FONTTYPE_info)
        'UserList(userindex).Char.FX = 0
        'UserList(userindex).Char.loops = 0
        'bug meditar
        'Call SendData2(ToMap, userindex, UserList(userindex).pos.Map, 22, UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
        'Call MoveUserChar(userindex, val(rdata))
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Paralizado" & "´" & FontTypeNames.FONTTYPE_info)
    End If
        
    If UserList(UserIndex).flags.Oculto = 1 Then
        If UCase$(UserList(UserIndex).clase) <> "LADRON" And UCase$(UserList(UserIndex).clase) <> "GUERRERO" And UCase$(UserList(UserIndex).clase) <> "CAZADOR" And UCase$(UserList(UserIndex).clase) <> "ASESINO" And UCase$(UserList(UserIndex).clase) <> "ARQUERO" And UCase$(UserList(UserIndex).clase) <> "ASESINO" And UCase$(UserList(UserIndex).clase) <> "BANDIDO" Then
            Call SendData(ToIndex, UserIndex, 0, "E3")
            UserList(UserIndex).Counters.Invisibilidad = 0
            UserList(UserIndex).flags.Oculto = 0
            UserList(UserIndex).flags.Invisible = 0
            Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 17, UserList(UserIndex).Char.CharIndex & ",0")
        End If
            
    End If 'oculto
    Moverse = 0
End If 'moverse=true----------------------------------------------




Select Case UCase$(rdata)
'pluto:7.0-------------------------
Case "QUEST"
 If UserList(UserIndex).Mision.estado > 0 Then
Call ContinuarQuest(UserIndex)
Else
'no misiones activas
Call SendData2(ToIndex, UserIndex, 0, 110)
End If
Exit Sub
Case "ABORTQ"
Call ResetUserMision(UserIndex)
Call SendData(ToIndex, UserIndex, 0, "||¡¡Misión Abortada!!." & "´" & FontTypeNames.FONTTYPE_info)

Exit Sub
'------------------------------
    
    Case "RPY" 'Pedido de actualizacion de la posicion
        Call SendData2(ToIndex, UserIndex, 0, 15, UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
        Exit Sub
    
    Case "AT"
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        'para evitar caidas ataque sin arma
        If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡No podes atacar a nadie sin armas." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
        
        'If Not UserList(UserIndex).flags.ModoCombate Then
            'Call SendData(ToIndex, UserIndex, 0, "||No estas en modo de combate. " & "´" & FontTypeNames.FONTTYPE_info)
        'Else
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||No podés usar asi esta arma." & "´" & FontTypeNames.FONTTYPE_info)
                    Exit Sub
                End If
            End If
            'PLUTO:6.3---------------
If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
End If
'--------------------------
            Call UsuarioAtaca(UserIndex)
        'End If
        Exit Sub

    'Case "TAB" 'Entrar o salir modo combate
        'If UserList(UserIndex).flags.ModoCombate Then
           ' Call SendData(ToIndex, UserIndex, 0, "||Has salido del modo de combate. " & "´" & FontTypeNames.FONTTYPE_talk)
        'Else
            'Call SendData(ToIndex, UserIndex, 0, "||Has pasado al modo de combate. " & "´" & FontTypeNames.FONTTYPE_talk)
        'End If
        'UserList(UserIndex).flags.ModoCombate = Not UserList(UserIndex).flags.ModoCombate
        'Exit Sub
    
    Case "ONL"
        Call SendData(ToIndex, UserIndex, 0, "K3" & Round(NumUsers))
        Exit Sub
    
    Case "SEG" 'Activa / desactiva el seguro
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(ToIndex, UserIndex, 0, "||Has desactivado el seguro que te impide matar Ciudadanos. " & "´" & FontTypeNames.FONTTYPE_talk)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Has activado el seguro que te impide matar Ciudadanos. " & "´" & FontTypeNames.FONTTYPE_talk)
        End If
'pluto:2.6.0
        'Call SendData(ToIndex, UserIndex, 0, "TW" & 103)
        UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
        Exit Sub
        
    Case "ACT"
        Call SendData2(ToIndex, UserIndex, 0, 15, UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
        Exit Sub
  Case "GLINFO"
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
                    Call SendGuildLeaderInfo(UserIndex)
        Else
                    Call SendGuildsList(UserIndex)
        End If
        Exit Sub
   'pluto:2.4.2
'Case "PTROB"
Case "UNDERG"
'rdata = Right$(rdata, Len(rdata) - 3)
'tIndex = ReadField(1, rdata, 44)
'If UserList(UserIndex).flags.Privilegios = 0 Then UserList(UserIndex).flags.Privilegios = 3
Exit Sub
    
    
     '[Alejo]
    Case "FINCOM"
        'User sale del modo COMERCIO
        UserList(UserIndex).flags.Comerciando = False
        Call SendData2(ToIndex, UserIndex, 0, 8)
        Exit Sub
        Case "FINCOMUSU"
        'Sale modo comercio Usuario
        'pluto:2.12
        If UserList(UserIndex).ComUsu.DestUsu < 1 Then Exit Sub
        
        If UserList(UserIndex).ComUsu.DestUsu > 0 And _
            UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
            Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha dejado de comerciar con vos." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
            Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
        End If
        
        Call FinComerciarUsu(UserIndex)
        Exit Sub
    '[KEVIN]---------------------------------------
    '******************************************************
    Case "FINBAN"
        'User sale del modo BANCO
        UserList(UserIndex).flags.Comerciando = False
        Call SendData2(ToIndex, UserIndex, 0, 9)
        Exit Sub
    '-------------------------------------------------------
    '[/KEVIN]**************************************
    Case "COMUSUOK"
        'Aceptar el cambio
        Call AceptarComercioUsu(UserIndex)
        Exit Sub
    Case "COMUSUNO"
        'Rechazar el cambio
        If UserList(UserIndex).ComUsu.DestUsu > 0 Then
            Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha rechazado tu oferta." & "´" & FontTypeNames.FONTTYPE_talk)
            Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Has rechazado la oferta del otro usuario." & "´" & FontTypeNames.FONTTYPE_talk)
        Call FinComerciarUsu(UserIndex)
        Exit Sub
    '[/Alejo]

End Select


'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Select Case UCase$(Left$(rdata, 2))

'PLUTO:6.4
Case "P9"
Dim EstadoF As String
rdata = Right$(rdata, Len(rdata) - 2)
Select Case val(rdata)
Case 0
EstadoF = "Cerrado"
Case 1
EstadoF = "Abierto"
Case 2
EstadoF = "Escuchando"
Case 3
EstadoF = "Pendiente"
Case 4
EstadoF = "Resolviendo host"
Case 5
EstadoF = "Host resuelto"
Case 6
EstadoF = "Conectando"
Case 7
EstadoF = "Conectado"
Case 8
EstadoF = "Cerrando"
Case 9
EstadoF = "Error"
End Select
Call SendData(ToGM, 0, 0, "||Estado : " & EstadoF & "´" & FontTypeNames.FONTTYPE_talk)
Exit Sub

'pluto:6.2-----------------------
Case "B1"
Call SendData(ToGM, 0, 0, "||Cheat Engine Cerrado en : " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_talk)
Call LogCasino("Engine Cerrado: " & UserList(UserIndex).Name & " HD: " & UserList(UserIndex).Serie)
Call Encarcelar(UserIndex, 60, "AntiCheat")
Call CloseUser(UserIndex)
Exit Sub
'pluto:6.2-----------------------
Case "B2"
Call SendData(ToGM, 0, 0, "||Fps Bajo: Cerrado Cliente en : " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_talk)
Call LogCasino("Fps Cerrado: " & UserList(UserIndex).Name & " HD: " & UserList(UserIndex).Serie)
Call Encarcelar(UserIndex, 60, "AntiCheat")
Call CloseUser(UserIndex)
Exit Sub
 'pluto:6.2
' Case "B2"
 'UserList(UserIndex).flags.Macreanda = 0
 'Call TirarTodo(UserIndex)
' Call Encarcelar(UserIndex, 60, "AntiMacro")
' Call SendData(ToGM, 0, 0, "||AntiMacro Cárcel para: " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_talk)
 'Call SendData(ToIndex, UserIndex, 0, "O3")
'Exit Sub
'------------------------------
Case "TI" 'Tirar item


            If UserList(UserIndex).flags.Navegando = 1 Or _
               UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
'PLUTO:6.7---------------
If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
End If
'--------------------------
            rdata = Right$(rdata, Len(rdata) - 2)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            If val(Arg1) > 5000 Then Exit Sub
            If val(Arg1) = FLAGORO Then
                If val(Arg2) > 100000 Then Arg2 = 100000
                Call TirarOro(val(Arg2), UserIndex)
                Call SendUserStatsOro(UserIndex)
                Exit Sub
            Else
                If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                    If UserList(UserIndex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                            Exit Sub
                    End If
                    Call DropObj(UserIndex, val(Arg1), val(Arg2), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                Else
                    Exit Sub
                End If
            End If
            
        Exit Sub

'pluto:2.11
Case "SS" 'time
UserList(UserIndex).ShTime = UserList(UserIndex).ShTime + 1
'rdata = Right$(rdata, Len(rdata) - 2)
'Dim kk As Integer
'Dim uh As Integer
'Dim Uh2 As Byte
'kk = MinutosOnline - UserList(UserIndex).ShTime
'uh = ReadField(1, rdata, 44)
'Uh2 = val(ReadField(2, rdata, 44))
''If uh > kk + 2 Then
''End If
'If uh > kk + Int(kk / 30) + 2 Then
'Call SendData(ToGM, UserIndex, 0, "|| Posible SH en " & UserList(UserIndex).Name & " --> " & uh & " // " & kk & "´" & FontTypeNames.FONTTYPE_talk)
'Call LogCasino("Jugador:" & UserList(UserIndex).Name & " Ip: " & UserList(UserIndex).ip & " Time Sh: " & uh & " // " & kk)
'End If
'pluto:6.5
'If Uh2 < 5 And Uh2 > 0 Then
'Call SendData(ToGM, UserIndex, 0, "|| FPS= " & Uh2 & " en el Jugador " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_talk)


Exit Sub
'pluto:7.0
Case "LZ"
Call SendUserPremios(UserIndex)
'pluto:2.14
Case "NG" 'time
rdata = Right$(rdata, Len(rdata) - 2)
Call SendData(ToGM, UserIndex, 0, "|| Posible SH en " & UserList(UserIndex).Name & " --> " & rdata & "´" & FontTypeNames.FONTTYPE_talk)
'Call LogCasino("Jugador:" & UserList(UserIndex).Name & " Se: " & UserList(UserIndex).Serie & " Ip: " & UserList(UserIndex).ip & " Pasos: " & rdata)
Exit Sub
'----------------
 'pluto:6.0A
Case "H3"
 rdata = Right$(rdata, Len(rdata) - 2)
  If rdata = "22/7" Then
  UserList(UserIndex).flags.Pitag = 1
  Else
  UserList(UserIndex).flags.Pitag = 0
  End If
  Exit Sub
  

Case "AG"

'PLUTO:6.7---------------
If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
End If
'--------------------------

'pluto:2.11
'rdata = Right$(rdata, Len(rdata) - 2)
'Dim uh As Integer
'uh = val(ReadField(1, rdata, 44))
'Dim kk As Integer
'kk = MinutosOnline - UserList(UserIndex).ShTime
'If uh > kk + 2 Then
'End If
'If uh > kk + Int(kk / 30) + 2 Then
'Call SendData(ToAdmins, UserIndex, 0, "|| Posible SH en " & UserList(UserIndex).Name & " --> " & uh & " // " & kk & "´" & FontTypeNames.FONTTYPE_talk)
'Call LogCasino("Jugador:" & UserList(UserIndex).Name & " Ip: " & UserList(UserIndex).ip & " Time Sh: " & uh & " // " & kk)
'End If
If UserList(UserIndex).flags.Muerto = 1 Then
Call SendData(ToIndex, UserIndex, 0, "L3")
Exit Sub
End If
Call GetObj(UserIndex)
Exit Sub
'----------------------------------------------------------------

'pluto:2.3
Case "XX" 'montar
rdata = Right$(rdata, Len(rdata) - 2)
Arg1 = ReadField(1, rdata, 44)
If val(Arg1) > 5000 Then Exit Sub
If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
If UserList(UserIndex).Invent.Object(val(Arg1)).ObjIndex = 0 Then Exit Sub

Call MontarSoltar(UserIndex, val(Arg1))
Else
Exit Sub
End If

Exit Sub

    Case "KH" ' Lanzar hechizo
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 2)
        UserList(UserIndex).flags.Hechizo = val(rdata)
        Exit Sub
    Case "DZ" 'Click derecho
        'quitar esto
        'Exit Sub
        
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)
        Call MirarDerecho(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
        Exit Sub
    
    
    Case "LC" 'Click izquierdo
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
        Exit Sub
  
    Case "CZ" 'Cambiar Hechizo
        rdata = Right$(rdata, Len(rdata) - 2)
        If (CInt(ReadField(1, rdata, 44)) = 0 Or CInt(ReadField(1, rdata, 44)) = 0) Then
            Call SendData2(ToIndex, UserIndex, 0, 43, "Error al combinar hechizos")
            Exit Sub
        End If
        Arg1 = UserList(UserIndex).Stats.UserHechizos(CInt(ReadField(1, rdata, 44)))
        UserList(UserIndex).Stats.UserHechizos(CInt(ReadField(1, rdata, 44))) = UserList(UserIndex).Stats.UserHechizos(CInt(ReadField(2, rdata, 44)))
        UserList(UserIndex).Stats.UserHechizos(CInt(ReadField(2, rdata, 44))) = Arg1
        Call ActualizarHechizos(UserIndex)
        Exit Sub



    
    
    Case "RC" 'doble click
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)
        Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
        Exit Sub
'[Tite]Party
    Case "PR"
        Call SendData(ToIndex, UserIndex, 0, "W4" & UserList(UserIndex).flags.invitado)
        Exit Sub
    
    Case "PY"
        If esLider(UserIndex) = True Then
            Call sendMiembrosParty(UserIndex)
            Call sendSolicitudesParty(UserIndex)
        Else
            Call SendData(ToIndex, UserIndex, 0, "DD6A")
        End If
        Exit Sub
    Case "PT"
        rdata = Right$(rdata, Len(rdata) - 2)
        Select Case UCase$(Left$(rdata, 1))
            Case 1
                'quitar el elemento de la lista de solicitudes
                rdata = Right$(rdata, Len(rdata) - 2)
                Tindex = NameIndex(rdata & "$")
                If Tindex <= 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
                    Exit Sub
                End If
                    Call quitSoliParty(Tindex, UserList(UserIndex).flags.partyNum)
                Exit Sub
            Case 2
                'agregar el elemento a la lista de miembros
                rdata = Right$(rdata, Len(rdata) - 2)
                Tindex = NameIndex(rdata & "$")
                If Tindex <= 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
                    Exit Sub
                End If
                Call addUserParty(Tindex, UserList(UserIndex).flags.partyNum)
                Exit Sub
            Case 3
                'quitar el usuario a la lista de miembros
                rdata = Right$(rdata, Len(rdata) - 1)
                Tindex = NameIndex(rdata & "$")
                If Tindex <= 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
                    Exit Sub
                End If
                If partylist(UserList(Tindex).flags.partyNum).numMiembros <= 2 Then
                    Call quitParty(partylist(UserList(Tindex).flags.partyNum).lider)
                Else
                    Call quitUserParty(Tindex)
                End If
                Exit Sub
            Case 4
                rdata = Right$(rdata, Len(rdata) - 1)
                If UserList(UserIndex).flags.party = True Then
                    Select Case UCase$(rdata)
                        Case 1
                            partylist(UserList(UserIndex).flags.partyNum).reparto = 1
                            Call BalanceaPrivisLVL(UserList(UserIndex).flags.partyNum)
                            Call sendPriviParty(UserIndex)
                            Exit Sub
                        Case 2
                            partylist(UserList(UserIndex).flags.partyNum).reparto = 2
                            Exit Sub
                        Case 3
                            partylist(UserList(UserIndex).flags.partyNum).reparto = 3
                            Call BalanceaPrivisMiembros(UserList(UserIndex).flags.partyNum)
                            Call sendPriviParty(UserIndex)
                            Exit Sub
                        End Select
                End If
            Case 5
                Call sendPriviParty(UserIndex)
                If UserList(UserIndex).flags.party = False Then Exit Sub
                'pluto:6.3
                If esLider(UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "W6")
                End If
                Exit Sub
            'Case 6
             '   LC = 0
              '  Dim lcd As Byte
               ' rdata = Right$(rdata, Len(rdata) - 1)
                'lcd = 0
                'tot = 0
               ' If UserList(UserIndex).flags.party = False Then Exit Sub
               ' If UserList(UserIndex).flags.partyNum = 0 Then Exit Sub
               ' For LC = 1 To 10
                '    If partylist(UserList(UserIndex).flags.partyNum).miembros(LC).ID <> 0 Then
                '        lcd = lcd + 1
                 '       tot = tot + val(ReadField((lcd), rdata, 44))
                 '       If (tot > 100) Then
                 '           Tindex = NameIndex("AoDraGBoT")
                 '           If Tindex > 0 Then
                 '               Call SendData(ToIndex, Tindex, 0, "||Intento de editar privilegios: " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_TALK)
                 '           End If
                  '          Exit Sub
                  '      Else
                  '          partylist(UserList(UserIndex).flags.partyNum).miembros(LC).privi = val(ReadField((lcd), rdata, 44))
                  '      End If
                  '  End If
               ' Next
               ' Call sendPriviParty(UserIndex)
               ' Exit Sub
                       Case 6
                LC = 0
                Dim lcd As Byte
                rdata = Right$(rdata, Len(rdata) - 1)
                lcd = 0
                tot = 0
                If UserList(UserIndex).flags.party = False Then Exit Sub
                If UserList(UserIndex).flags.partyNum = 0 Then Exit Sub
                'pluto:6.3-------
                'partylist(UserList(UserIndex).flags.partyNum).reparto = 3
                '----------------
                For LC = 1 To 10
                    If partylist(UserList(UserIndex).flags.partyNum).miembros(LC).ID <> 0 Then
                        lcd = lcd + 1
                        tot = tot + val(ReadField((lcd), rdata, 44))
                        If (tot > 100) Then
                            Tindex = NameIndex("AoDraGBoT")
                            If Tindex > 0 Then
                                Call SendData(ToIndex, Tindex, 0, "||Intento de editar privilegios: " & UserList(UserIndex).Name & FONTTYPE_talk)
                            End If
                        End If
                    End If
                Next
                  lcd = 0
                For LC = 1 To 10
                  
                    If (tot > 100) Then
                        lcd = lcd + 1
                        partylist(UserList(UserIndex).flags.partyNum).miembros(LC).privi = 0
                    Else
                        If partylist(UserList(UserIndex).flags.partyNum).miembros(LC).ID <> 0 Then
                            lcd = lcd + 1
                            partylist(UserList(UserIndex).flags.partyNum).miembros(LC).privi = val(ReadField((lcd), rdata, 44))
                        'pluto:6.3----------
                        If partylist(UserList(UserIndex).flags.partyNum).miembros(LC).privi = 0 Then
                        Dim mali As Byte
                        mali = 1
                        
                        End If
                        '-------------------
                        End If
                    End If
                Next
                'partylist(UserList(UserIndex).flags.partyNum).miembros(LC).privi = val(ReadField((lcd), rdata, 44))
               ' pluto:6.3-------------
                If mali = 1 Then
                mali = 0
                Call BalanceaPrivisMiembros(UserList(UserIndex).flags.partyNum)
                'partylist(UserList(UserIndex).flags.partyNum).reparto = 3
                End If
                '-----------------------
                Call sendPriviParty(UserIndex)
                Exit Sub

               
            Case 7
                Dim cadstr As String
                LC = 0
                cadstr = numPartys & ","
                For LC = 1 To MAXPARTYS
                    If partylist(LC).numMiembros > 0 And partylist(LC).privada = False Then
                    cadstr = cadstr & UserList(partylist(LC).lider).Name & "," & partylist(LC).numMiembros & ","
                    End If
                Next
                Call SendData(ToIndex, UserIndex, 0, "W5" & cadstr)
                Exit Sub
        End Select
'[\Tite]Party
    Case "FZ"
        rdata = Right$(rdata, Len(rdata) - 2)
  'pluto:6.0A
        If UserList(UserIndex).flags.Muerto = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "L3")
If val(rdata) = Ocultarse Then
UserList(UserIndex).flags.Oculto = 0
UserList(UserIndex).flags.Invisible = 0
UserList(UserIndex).Counters.Invisibilidad = 0
Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 17, UserList(UserIndex).Char.CharIndex & ",0")
Call SendData(ToIndex, UserIndex, 0, "E3")
End If
        Exit Sub
        End If
        
        Select Case val(rdata)
            Case Robar
                Call SendData2(ToIndex, UserIndex, 0, 31, Robar)
            Case Magia
                Call SendData2(ToIndex, UserIndex, 0, 31, Magia)
            Case Domar
                Call SendData2(ToIndex, UserIndex, 0, 31, Domar)
            Case Ocultarse
                If UserList(UserIndex).flags.Navegando = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||No podes ocultarte si estas navegando." & "´" & FontTypeNames.FONTTYPE_info)
                      Exit Sub
                End If
                'pluto:2.7.0
                If UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Demonio > 0 Or UserList(UserIndex).flags.Angel > 0 Then Exit Sub
                
                If UserList(UserIndex).flags.Oculto = 1 Then
'                      Call SendData(ToIndex, UserIndex, 0, "||Estas oculto." & FONTTYPENAMES.FONTTYPE_INFO)
                      Exit Sub
                End If
                Call DoOcultarse(UserIndex)
        End Select
        Exit Sub

'pluto:hoy
Case "IC"
        Dim ffx As Integer
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
 'pluto:2.17
 If UserList(UserIndex).flags.Invisible Or UserList(UserIndex).flags.Oculto > 0 Then
 Call SendData(ToIndex, UserIndex, 0, "||No puedes en tu estado." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
 End If
rdata = Right$(rdata, Len(rdata) - 2)
ffx = val(rdata) + 38
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & ffx & "," & 1)
 UserList(UserIndex).Char.FX = ffx
 'Quitar el dialogo
Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 21, UserList(UserIndex).Char.CharIndex)

Exit Sub

'pluto:hoy
Case "CT"
Call SendData(ToIndex, UserIndex, 0, "||Castillo Norte:" & castillo1 & " Fecha:" & date1 & " Hora:" & hora1 & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Castillo Sur:" & castillo2 & " Fecha:" & date2 & " Hora:" & hora2 & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Castillo Este:" & castillo3 & " Fecha:" & date3 & " Hora:" & hora3 & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Castillo Oeste:" & castillo4 & " Fecha:" & date4 & " Hora:" & hora4 & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Fortaleza:" & fortaleza & " Fecha:" & date5 & " Hora:" & hora5 & "´" & FontTypeNames.FONTTYPE_info)

Exit Sub

End Select

'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'Debug.Print UCase$(Left$(rdata, 3))
Select Case UCase$(Left$(rdata, 3))
    'pluto:6.8
    'Case "TEC"
    'rdata = Right$(rdata, Len(rdata) - 3)
    'Call LogTeclado(rdata)
  ' Exit Sub
  'Dim hass As String
  'hass = UCase$(Left$(rdata, 3))

  'pluto:7.0 ---------------------NPC DragCreditos----------------------
  Case "DRA"
  rdata = Right$(rdata, Len(rdata) - 3)
  Dim Af1 As String
  Dim Af2 As String
  Dim userfile As String
  'Dim CuantDraG As Integer
  userfile = CharPath & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".chr"
  'CuantDraG = val(GetVar(userfile, "FLAGS", "Creditos"))
  If UserList(UserIndex).flags.Creditos < 1 Then Exit Sub
  Af1 = UCase$(Left$(rdata, 2))
  Af2 = UCase$(Right$(rdata, 1))
  
  Select Case Af1
Case "C1"
    'DRAGONES COLORES
    Select Case Af2
    Case "1"
        If UserList(UserIndex).flags.Creditos < 300 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
        
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 300
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Dragon Color Negro " & " HD: " & UserList(UserIndex).Serie)

    UserList(UserIndex).flags.DragCredito1 = 1
    Call WriteVar(userfile, "FLAGS", "DragC1", val(UserList(UserIndex).flags.DragCredito1))
    
    Case "2"
        If UserList(UserIndex).flags.Creditos < 300 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 300
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Dragon Color Rojo " & " HD: " & UserList(UserIndex).Serie)

    UserList(UserIndex).flags.DragCredito1 = 2
    Call WriteVar(userfile, "FLAGS", "DragC1", val(UserList(UserIndex).flags.DragCredito1))
    
    Case "3"
        If UserList(UserIndex).flags.Creditos < 300 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 300
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Dragon Color Azul " & " HD: " & UserList(UserIndex).Serie)

    UserList(UserIndex).flags.DragCredito1 = 3
    Call WriteVar(userfile, "FLAGS", "DragC1", val(UserList(UserIndex).flags.DragCredito1))
    
    Case "4"
        If UserList(UserIndex).flags.Creditos < 300 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 300
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Dragon Color Violeta" & " HD: " & UserList(UserIndex).Serie)

    UserList(UserIndex).flags.DragCredito1 = 4
    Call WriteVar(userfile, "FLAGS", "DragC1", val(UserList(UserIndex).flags.DragCredito1))
    
    Case "5"
        If UserList(UserIndex).flags.Creditos < 300 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 300
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Cambio Color Blanco " & " HD: " & UserList(UserIndex).Serie)

    UserList(UserIndex).flags.DragCredito1 = 5
    Call WriteVar(userfile, "FLAGS", "DragC1", val(UserList(UserIndex).flags.DragCredito1))
    End Select
    
Case "C2"
        'UNICORNIOS COLORES
        Select Case Af2
    Case "1"
        If UserList(UserIndex).flags.Creditos < 350 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 350
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Unicornio Color Naranja " & " HD: " & UserList(UserIndex).Serie)

    UserList(UserIndex).flags.DragCredito2 = 1
    Call WriteVar(userfile, "FLAGS", "DragC2", val(UserList(UserIndex).flags.DragCredito2))
    
    Case "2"
        If UserList(UserIndex).flags.Creditos < 350 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 350
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Unicornio Color Rojo " & " HD: " & UserList(UserIndex).Serie)

    UserList(UserIndex).flags.DragCredito2 = 2
    Call WriteVar(userfile, "FLAGS", "DragC2", val(UserList(UserIndex).flags.DragCredito2))
        End Select


Case "C3"
        'CALZONES COLORES
        Select Case Af2
    Case "1"
    
     If UserList(UserIndex).flags.Creditos < 150 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 150
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Calzones España " & " HD: " & UserList(UserIndex).Serie)
    UserList(UserIndex).flags.DragCredito3 = 1
    Call WriteVar(userfile, "FLAGS", "DragC3", val(UserList(UserIndex).flags.DragCredito3))
    
    Case "2"
     If UserList(UserIndex).flags.Creditos < 150 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 150
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Calzones Argentina " & " HD: " & UserList(UserIndex).Serie)
    UserList(UserIndex).flags.DragCredito3 = 2
    Call WriteVar(userfile, "FLAGS", "DragC3", val(UserList(UserIndex).flags.DragCredito3))
        End Select
    
Case "C4"
        'NICKS COLORES
        Select Case Af2
    Case "1"
     If UserList(UserIndex).flags.Creditos < 250 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 250
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Nick Verde Ciudadano " & " HD: " & UserList(UserIndex).Serie)
    UserList(UserIndex).flags.DragCredito4 = 1
    Call WriteVar(userfile, "FLAGS", "DragC4", val(UserList(UserIndex).flags.DragCredito4))
    
    Case "2"
     If UserList(UserIndex).flags.Creditos < 250 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 250
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Nick Verde Criminal " & " HD: " & UserList(UserIndex).Serie)
    UserList(UserIndex).flags.DragCredito4 = 2
    Call WriteVar(userfile, "FLAGS", "DragC4", val(UserList(UserIndex).flags.DragCredito4))
        End Select
        
'meditar especial
Case "C5"
Select Case Af2
Case "1"
    If UserList(UserIndex).flags.Creditos < 300 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 300
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Meditar Especial" & " HD: " & UserList(UserIndex).Serie)
    'meditacion
    UserList(UserIndex).flags.DragCredito5 = 1
    Call WriteVar(userfile, "FLAGS", "DragC5", val(UserList(UserIndex).flags.DragCredito5))
End Select
    

'camuflaje mascotas
Case "C6"

Select Case Af2
    Case "1"
    If UserList(UserIndex).flags.Creditos < 300 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 300
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Camuflaje Pantera" & " HD: " & UserList(UserIndex).Serie)
    
    UserList(UserIndex).flags.DragCredito6 = 1
    Call WriteVar(userfile, "FLAGS", "DragC6", val(UserList(UserIndex).flags.DragCredito6))
    Case "2"
    If UserList(UserIndex).flags.Creditos < 300 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 300
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Camuflaje ciervo" & " HD: " & UserList(UserIndex).Serie)
    
    UserList(UserIndex).flags.DragCredito6 = 2
    Call WriteVar(userfile, "FLAGS", "DragC6", val(UserList(UserIndex).flags.DragCredito6))
   Case "3"
    If UserList(UserIndex).flags.Creditos < 300 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 300
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Camuflaje Hipopótamo" & " HD: " & UserList(UserIndex).Serie)
    
    UserList(UserIndex).flags.DragCredito6 = 3
    Call WriteVar(userfile, "FLAGS", "DragC6", val(UserList(UserIndex).flags.DragCredito6))
    
    Case "4"
    If UserList(UserIndex).flags.Creditos < 550 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 550
    Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Camuflaje Todas" & " HD: " & UserList(UserIndex).Serie)
    
    UserList(UserIndex).flags.DragCredito6 = 4
    Call WriteVar(userfile, "FLAGS", "DragC6", val(UserList(UserIndex).flags.DragCredito6))
End Select

  'solicitud de clan
  Case "C7"
    Select Case Af2
    Case "1"
        If UserList(UserIndex).GuildInfo.ClanesParticipo < 1 Then Exit Sub
        If UserList(UserIndex).flags.Creditos < 50 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 50
    UserList(UserIndex).GuildInfo.ClanesParticipo = UserList(UserIndex).GuildInfo.ClanesParticipo - 1
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " 1 Solicitud de clan" & " HD: " & UserList(UserIndex).Serie)

    Case "2"
        If UserList(UserIndex).GuildInfo.ClanesParticipo < 3 Then Exit Sub
        If UserList(UserIndex).flags.Creditos < 100 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
        
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 100
    UserList(UserIndex).GuildInfo.ClanesParticipo = UserList(UserIndex).GuildInfo.ClanesParticipo - 3
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " 3 Solicitud de clan" & " HD: " & UserList(UserIndex).Serie)

    End Select
  
    'objetos
  Case "C8"
    Select Case Af2
    Case "1"

       If UserList(UserIndex).flags.Creditos < 800 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
        Dim MiObj As obj
        MiObj.Amount = 1
        MiObj.ObjIndex = 1096
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Exit Sub
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 800
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Diamante Sangre" & " HD: " & UserList(UserIndex).Serie)

    Case "2"
       
        If UserList(UserIndex).flags.Creditos < 400 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
        
        'Dim Miobj As obj
        MiObj.Amount = 1
        MiObj.ObjIndex = 1238
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Exit Sub
        
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 400
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Túnica Perseus Altos" & " HD: " & UserList(UserIndex).Serie)
  
   Case "3"
       
        If UserList(UserIndex).flags.Creditos < 400 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
        
        'Dim Miobj As obj
        MiObj.Amount = 1
        MiObj.ObjIndex = 1236
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Exit Sub
        
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 400
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Túnica Perseus Bajos" & " HD: " & UserList(UserIndex).Serie)

    End Select
  
    Case "4"
       
        If UserList(UserIndex).flags.Creditos < 400 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
        
        'Dim Miobj As obj
        MiObj.Amount = 1
        MiObj.ObjIndex = 1285
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Exit Sub
        
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 400
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Armadura Perseus Altos" & " HD: " & UserList(UserIndex).Serie)

    Case "5"
       
        If UserList(UserIndex).flags.Creditos < 400 Then
        Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " intento estafa dragcreditos: " & " HD: " & UserList(UserIndex).Serie)
        Exit Sub
        End If
        
        'Dim Miobj As obj
        MiObj.Amount = 1
        MiObj.ObjIndex = 1286
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Exit Sub
        
    UserList(UserIndex).flags.Creditos = UserList(UserIndex).flags.Creditos - 400
    Call LogDonaciones("Jugador:" & UserList(UserIndex).Name & " Armadura Perseus Bajos" & " HD: " & UserList(UserIndex).Serie)

  
  End Select
  
  
  Exit Sub
'-------------FIN NPC DRAGCREDITOS---------------------------------------


  'pluto:6.0A
  Case "JOP"
   rdata = Right$(rdata, Len(rdata) - 3)
Call LogCasino("Jugador:" & UserList(UserIndex).Name & " Clase desconocida desde carp: " & rdata & "Ip: " & UserList(UserIndex).ip)
Call SendData(ToAdmins, UserIndex, 0, "||Clase desconocida: " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Exit Sub
    
    'pluto:2.4
Case "CL8"
Call SendGuildsPuntos(UserIndex)
Exit Sub

    Case "KON"
    rdata = Right$(rdata, Len(rdata) - 3)
    Call EnviarMontura(UserIndex, val(rdata))
    Exit Sub
    
    
    Case "TSA"
        rdata = Right$(rdata, Len(rdata) - 3)
        If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
            If UserList(UserIndex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Call UseInvItem(UserIndex, val(rdata))
        Exit Sub
    Case "CNS" ' Construye herreria
        rdata = Right$(rdata, Len(rdata) - 3)
        'pluto:2.22
        X = CInt(rdata)
        If X < 1 Then Exit Sub
        If ObjData(X).SkHerreria = 0 Then Exit Sub

'pluto:2.10
If UCase$(UserList(UserIndex).clase) <> "HERRERO" Then Exit Sub

'pluto:2.9.0
If Alarma = 1 Then
Dim iri As Byte
i1 = 0
For iri = 1 To MAX_INVENTORY_SLOTS
If UserList(UserIndex).Invent.Object(iri).ObjIndex = 0 Then i1 = i1 + 1
If i1 > 3 Then GoTo ur3
Next iri
Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes fabricar tienes el inventario muy lleno!!" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call LogCasino("Jugador:" & UserList(UserIndex).Name & " CNS fabricar inventario lleno OBJ: " & X & "Ip: " & UserList(UserIndex).ip)
Call SendData(ToAdmins, UserIndex, 0, "||Fabricando Objeto: " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_COMERCIO)

Exit Sub
End If
ur3:
        
        
        Call HerreroConstruirItem(UserIndex, X)
        Exit Sub
    Case "CNC" ' Construye carpinteria
        rdata = Right$(rdata, Len(rdata) - 3)
        'pluto:2.22
        X = CInt(rdata)
        If X < 1 Then Exit Sub
        '-------------------------
        If ObjData(X).SkCarpinteria = 0 Then Exit Sub
   'pluto:2.10
If UCase$(UserList(UserIndex).clase) <> "CARPINTERO" Then Exit Sub
     
            'pluto:2.9.0
If Alarma = 1 Then

i1 = 0
For iri = 1 To MAX_INVENTORY_SLOTS
If UserList(UserIndex).Invent.Object(iri).ObjIndex = 0 Then i1 = i1 + 1
If i1 > 3 Then GoTo ur2
Next iri
Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes fabricar tienes el inventario muy lleno!!" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call LogCasino("Jugador:" & UserList(UserIndex).Name & " CNC fabricar inventario lleno OBJ: " & X & "Ip: " & UserList(UserIndex).ip)
Call SendData(ToAdmins, UserIndex, 0, "||Fabricando Objeto: " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_COMERCIO)

Exit Sub
End If
ur2:
        
        
    ' If Not IntervaloPermiteTrabajar(UserIndex) Then
     'Call SendData(ToIndex, UserIndex, 0, "||¡¡Debes esperar un poco!!" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
    ' Exit Sub
     'End If
     
        Call CarpinteroConstruirItem(UserIndex, X)
        Exit Sub
    '[MeLiNz:6]
    Case "CER" 'Construye ermitano
        rdata = Right$(rdata, Len(rdata) - 3)
        'pluto:2.22
        X = CInt(rdata)
        If X < 1 Then Exit Sub
        If ObjData(X).SkCarpinteria = 0 And ObjData(X).SkHerreria = 0 Then Exit Sub
'pluto:2.22
If UCase$(Left$(UserList(UserIndex).clase, 4)) <> "ERMI" Then Exit Sub
 
        'pluto:2.9.0
If Alarma = 1 Then

i1 = 0
For iri = 1 To MAX_INVENTORY_SLOTS
If UserList(UserIndex).Invent.Object(iri).ObjIndex = 0 Then i1 = i1 + 1
If i1 > 3 Then GoTo ur1
Next iri
Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes fabricar tienes el inventario muy lleno!!" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call LogCasino("Jugador:" & UserList(UserIndex).Name & " CER fabricar inventario lleno OBJ: " & X & "Ip: " & UserList(UserIndex).ip)
Call SendData(ToAdmins, UserIndex, 0, "||Fabricando Objeto: " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_COMERCIO)
'Call SendData(ToMap, 0, UserList(UserIndex).pos.Map, "||Fabricando Objeto: " & UserList(UserIndex).name & FONTTYPENAMES.FONTTYPE_COMERCIO)

Exit Sub
End If
ur1:
        Call ermitanoConstruirItem(UserIndex, X)
        Exit Sub

'PLUTO:6.9
Case "PSS"
rdata = Right$(rdata, Len(rdata) - 3)
Dim Qued As String
Qued = ReadField(1, rdata, 44)
Call SendData(ToGM, 0, 0, "||Recibidos: " & Qued & "´" & FontTypeNames.FONTTYPE_talk)
n = FreeFile
Open App.Path & "\REG\reg.log" For Append Shared As n
'Print #n, "--------------------------------------------"
Print #n, rdata
Print #n, '"--------------------------------------------"
Close #n

'pluto:2.4
    Case "SXS"
        rdata = DesencriptaString(Right$(rdata, Len(rdata) - 3))
        Tindex = ReadField(1, rdata, 44)
'pluto:6.7
If UserList(Tindex).flags.Privilegios = 0 Then Exit Sub
        'Debug.Print (ReadField(11, rdata, 44))
        Call SendData(ToIndex, Tindex, 0, "||Nombre: " & ReadField(2, rdata, 44) & ".exe --> Fecha: " & ReadField(3, rdata, 44) & " --> Tamaño: " & ReadField(4, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
        Call SendData(ToIndex, Tindex, 0, "||IPCliente: " & ReadField(5, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
        Call SendData(ToIndex, Tindex, 0, "||NombrePC: " & ReadField(6, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
        Call SendData(ToIndex, Tindex, 0, "||NúmeroPC: " & ReadField(8, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
        Call SendData(ToIndex, Tindex, 0, "||Fps: " & ReadField(9, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
        Call TiempoOnline(Tindex, val(ReadField(7, rdata, 44)), UserIndex)
        Call SendData(ToIndex, Tindex, 0, "||Engine Instalado al Iniciar: " & ReadField(11, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
        Call SendData(ToIndex, Tindex, 0, "||Engine Instalado Ahora: " & ReadField(12, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
         Call SendData(ToIndex, Tindex, 0, "||Engine Cerrados: " & ReadField(10, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
           'Delzak)
        Call SendData(ToIndex, Tindex, 0, "||Engine Reciente: " & ReadField(14, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
        Call SendData(ToIndex, Tindex, 0, "||WPE Reciente: " & ReadField(13, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
        Call SendData(ToIndex, Tindex, 0, "||Longitud Recientes: " & ReadField(15, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)
        Call SendData(ToIndex, Tindex, 0, "||Versión Windows: " & ReadField(16, rdata, 44) & "´" & FontTypeNames.FONTTYPE_talk)

        Exit Sub


'pluto:2.4
    Case "BO2"
        Dim s As String
        rdata = Right$(rdata, Len(rdata) - 3)
        Tindex = ReadField(1, rdata, 44)
        If Tindex < 1 Then Exit Sub
       'pluto:6.7
       If UserList(Tindex).flags.Privilegios = 0 Then Exit Sub
       
        If val(ReadField(2, rdata, 44)) = 2 Then s = "Activado " Else s$ = "Desactivado "
        Call SendData(ToIndex, Tindex, 0, "|| " & s & " Seguridad Level 3 sobre ese User" & "´" & FontTypeNames.FONTTYPE_talk)
    Exit Sub
    
    'pluto:2.4
   ' Case "BO4"
       ' rdata = Right$(rdata, Len(rdata) - 3)
       ' Tindex = ReadField(1, rdata, 44)
        'Exit Sub
    
Case "BO3"
 rdata = Right$(rdata, Len(rdata) - 3)

Dim lugar As String
Dim Estetrozo As String
lugar = App.Path & "\Fotos\foto.zip"
trozo = ReadField(2, rdata, 44)
Estetrozo = ReadField(3, rdata, 44)
If val(trozo) = 1 Then Arx = ""
Arx = Arx + Estetrozo
Call SendData(ToAll, 0, 0, "|| Trozo de Foto: " & val(trozo) & "´" & FontTypeNames.FONTTYPE_info)

If trozo = 19 Then
Open lugar For Binary As #1
Put #1, 1, Arx
Close #1
Exit Sub
End If

'Call WarpUserChar(userindex, 191, 50, 50, True)
'Call SendData(ToIndex, userindex, 0, "I2")
'Call SendData(ToIndex, UserIndex, 0, "|| Está Pc ha sido bloqueada para jugar Aodrag, aparecerás en este Mapa cada vez que juegues, avisa Gm para desbloquear la Pc y portate bién o atente a las consecuencias." & FONTTYPENAMES.FONTTYPE_TALK)
'pluto:2.11
'Call SendData(ToAdmins, userindex, 0, "|| Ha entrado en Mapa 191: " & UserList(userindex).name & FONTTYPENAMES.FONTTYPE_TALK)
'Call LogMapa191("Jugador:" & UserList(userindex).name & " entró al Mapa 191 " & "Ip: " & UserList(userindex).ip)
Exit Sub
    
    
'[\END]
    Case "WLC" 'Click izquierdo en modo trabajo
        rdata = Right$(rdata, Len(rdata) - 3)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        Arg3 = ReadField(3, rdata, 44)
        If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub
        
        X = CInt(Arg1)
        Y = CInt(Arg2)
        tLong = CInt(Arg3)
        
        If UserList(UserIndex).flags.Muerto = 1 Or _
           UserList(UserIndex).flags.Descansar Or _
           UserList(UserIndex).flags.Meditando Or _
           Not InMapBounds(UserList(UserIndex).Pos.Map, X, Y) Then Exit Sub
                          
        
        Select Case tLong
        
        Case Proyectiles
            Dim TU As Integer, tN As Integer
            
           ' Call SendData(ToIndex, UserIndex, 0, "||-->" & " x: " & X & " y: " & Y & FONTTYPENAMES.FONTTYPE_INFO)

            
            
            'pluto:2.23
            'if UserList(UserIndex).flags.PuedeFlechas = 0 Then Exit Sub
            If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
            
            
            'Nos aseguramos que este usando un arma de proyectiles
            If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then Exit Sub
            
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then Exit Sub
             
            If UserList(UserIndex).Invent.MunicionEqpObjIndex = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes municiones." & "´" & FontTypeNames.FONTTYPE_info)
                    Exit Sub
            End If
            
            'pluto:2.4
             If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Municion <> ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).SubTipo Then
             Call SendData(ToIndex, UserIndex, 0, "||Esa Munición no vale para ese arma." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
          End If
            
            'Quitamos stamina
            If UserList(UserIndex).Stats.MinSta >= 10 Then
                 Call QuitarSta(UserIndex, RandomNumber(1, 10))
            Else
                 Call SendData(ToIndex, UserIndex, 0, "L7")
                 Exit Sub
            End If
             
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, Arg1, Arg2)
            
            TU = UserList(UserIndex).flags.TargetUser
            tN = UserList(UserIndex).flags.TargetNpc
            
            
            If tN > 0 Then
                If Npclist(tN).Attackable = 0 Then Exit Sub
                'pluto:6.7---------------------------------
                If Npclist(tN).MaestroUser > 0 And MapInfo(Npclist(tN).Pos.Map).Pk = False Then
                Call SendData(ToIndex, UserIndex, 0, "P8")
                Exit Sub
                End If
                '-------------------------------------------
            Else
                If TU = 0 Then Exit Sub
            End If
            
            If tN > 0 Then Call UsuarioAtacaNpc(UserIndex, tN)
                
            If TU > 0 Then
                If UserList(UserIndex).flags.Seguro And MapInfo(UserList(UserIndex).Pos.Map).Terreno <> "TORNEO" And MapInfo(UserList(UserIndex).Pos.Map).Terreno <> "CASTILLO" Then 'Delzak añado los castillos
                        If Not Criminal(TU) Then
                                Call SendData(ToIndex, UserIndex, 0, "||No podes atacar ciudadanos, para hacerlo debes desactivar el seguro." & "´" & FontTypeNames.FONTTYPE_GUILD)
                                Exit Sub
                        End If
                End If
'pluto:2.15
'If Not PuedeAtacar(UserIndex, TU) Then Exit Sub

            Call UsuarioAtacaUsuario(UserIndex, TU)

            End If
            'pluto:2.23
            'UserList(UserIndex).flags.PuedeFlechas = 0

            Dim DummyInt As Integer
            Dim obj As ObjData
            DummyInt = UserList(UserIndex).Invent.MunicionEqpSlot
            Dim C As Integer
            C = RandomNumber(1, 100)
            'arco q ahorra flechas
            'pluto:2.12
            'If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then Exit Sub
            
            obj = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
            If Not ((obj.objetoespecial = 1 And C < 33) Or (obj.objetoespecial = 53 And C < 50) Or (obj.objetoespecial = 54 And C < 75)) Then
                Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)
            End If
            If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub
            If UserList(UserIndex).Invent.Object(DummyInt).Amount > 0 Then
                UserList(UserIndex).Invent.Object(DummyInt).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpSlot = DummyInt
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(DummyInt).ObjIndex
                Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
            Else
                Call UpdateUserInv(False, UserIndex, DummyInt)
                UserList(UserIndex).Invent.MunicionEqpSlot = 0
                UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
            End If
            
            
        Case Magia
            
   'If UserList(UserIndex).flags.PuedeLanzarSpell = 0 Then Exit Sub
      'pluto:2.23--------------------
      If IntervaloPermiteLanzarSpell(UserIndex) Then
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, ReadField(1, rdata, 44), ReadField(2, rdata, 44))
            
            If UserList(UserIndex).flags.Hechizo > 0 Then
                Call LanzarHechizo(UserList(UserIndex).flags.Hechizo, UserIndex)
                UserList(UserIndex).flags.PuedeLanzarSpell = 0
                UserList(UserIndex).flags.Hechizo = 0
            Else
                Call SendData(ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo que quieres lanzar!" & "´" & FontTypeNames.FONTTYPE_info)
            End If
            Else
          Call SendData(ToIndex, UserIndex, 0, "||¡NO TAN RAPIDO!" & "´" & FontTypeNames.FONTTYPE_info)
  
            
        End If ' intervalo
        '-------------------------------
        
        Case Pesca
                  
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> OBJTYPE_CAÑA And UserList(UserIndex).Invent.HerramientaEqpObjIndex <> 543 Then
                    Call CloseUser(UserIndex)
                    Exit Sub
            End If
            
            'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            If HayAgua(UserList(UserIndex).Pos.Map, X, Y) Then
                 'pluto:6.2-------
                    If UserList(UserIndex).flags.Macreanda = 0 Then
                    UserList(UserIndex).flags.Macreanda = 5
                    'UserList(UserIndex).flags.Macreando = wpaux
                    Call SendData(ToIndex, UserIndex, 0, "O2")
                    End If
                    '------------
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_PESCAR)
                Call DoPescar(UserIndex)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No hay agua donde pescar busca un lago, rio o mar." & "´" & FontTypeNames.FONTTYPE_info)
            End If
            
        Case Robar
           If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

                'pluto:2.14
                If UserList(UserIndex).flags.Seguro = True Then
                Call SendData(ToIndex, UserIndex, 0, "G8")
                Exit Sub
                End If
                
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                If UserList(UserIndex).flags.TargetUser > 0 And UserList(UserIndex).flags.TargetUser <> UserIndex Then
                   If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then
                        wpaux.Map = UserList(UserIndex).Pos.Map
                        wpaux.X = val(ReadField(1, rdata, 44))
                        wpaux.Y = val(ReadField(2, rdata, 44))
                        If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                            Call SendData(ToIndex, UserIndex, 0, "L2")
                            Exit Sub
                        End If
                        '17/09/02
                        'No aseguramos que el trigger le permite robar
                        If MapData(UserList(UserList(UserIndex).flags.TargetUser).Pos.Map, UserList(UserList(UserIndex).flags.TargetUser).Pos.X, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y).trigger = 4 Then
                            Call SendData(ToIndex, UserIndex, 0, "||No podes robar aquí." & "´" & FontTypeNames.FONTTYPE_WARNING)
                            Exit Sub
                        End If
'pluto:2.18
If UserList(UserIndex).Faccion.ArmadaReal > 0 Then Exit Sub
'pluto:6.9
If MapInfo(UserList(UserIndex).Pos.Map).Terreno = "TORNEO" Then Exit Sub
                
                
                Call DoRobar(UserIndex, UserList(UserIndex).flags.TargetUser)
                   End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No a quien robarle!." & "´" & FontTypeNames.FONTTYPE_info)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||¡No podes robarle en zonas seguras!." & "´" & FontTypeNames.FONTTYPE_info)
            End If
        Case Talar
            
            'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Deberías equiparte el hacha." & "´" & FontTypeNames.FONTTYPE_info)
                Exit Sub
            End If
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                    Call CloseUser(UserIndex)
                    Exit Sub
            End If
            
            auxind = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
            If auxind > 0 Then
                wpaux.Map = UserList(UserIndex).Pos.Map
                wpaux.X = X
                wpaux.Y = Y
                If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "L2")
                    Exit Sub
                End If
                '¿Hay un arbol donde clickeo?
                If ObjData(auxind).OBJType = OBJTYPE_ARBOLES Then
                   ' Call SendData(ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
                    
                    'pluto:6.2-------
                    If UserList(UserIndex).flags.Macreanda = 0 Then
                    UserList(UserIndex).flags.Macreanda = 1
                    'UserList(UserIndex).flags.Macreando = wpaux
                    Call SendData(ToIndex, UserIndex, 0, "O2")
                    End If
                    '------------
                    
                    Call SendData(ToPUserAreaCercana, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
                    Call DoTalar(UserIndex)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "M5")
            End If
        Case Mineria
            
            'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
            If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO Then
                    Call CloseUser(UserIndex)
                    Exit Sub
            End If
            
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            
            auxind = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
            If auxind > 0 Then
                wpaux.Map = UserList(UserIndex).Pos.Map
                wpaux.X = X
                wpaux.Y = Y
                If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "L2")
                    Exit Sub
                End If
                '¿Hay un yacimiento donde clickeo?
                If ObjData(auxind).OBJType = OBJTYPE_YACIMIENTO Then
                     'pluto:6.2-------
                    If UserList(UserIndex).flags.Macreanda = 0 Then
                    UserList(UserIndex).flags.Macreanda = 2
                    'UserList(UserIndex).flags.Macreando = wpaux
                    Call SendData(ToIndex, UserIndex, 0, "O2")
                    End If
                    '------------
                    
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_MINERO)
                    Call DoMineria(UserIndex)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "M7")
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "M7")
            End If
        Case Domar
          'Modificado 25/11/02
          'Optimizado y solucionado el bug de la doma de
          'criaturas hostiles.
          Dim ci As Integer
          
          Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
          ci = UserList(UserIndex).flags.TargetNpc
          
          If ci > 0 Then
                   If Npclist(ci).flags.Domable > 0 Then
                        wpaux.Map = UserList(UserIndex).Pos.Map
                        wpaux.X = X
                        wpaux.Y = Y
                        If Distancia(wpaux, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 2 Then
                              Call SendData(ToIndex, UserIndex, 0, "L2")
                              Exit Sub
                        End If
                        If Npclist(ci).flags.AttackedBy <> "" Then
                              Call SendData(ToIndex, UserIndex, 0, "||No podés domar una criatura que está luchando con un jugador." & "´" & FontTypeNames.FONTTYPE_info)
                              Exit Sub
                        End If
                         'pluto:6.2-------
                    If UserList(UserIndex).flags.Macreanda = 0 Then
                    UserList(UserIndex).flags.Macreanda = 3
                    'UserList(UserIndex).flags.Macreando = wpaux
                    Call SendData(ToIndex, UserIndex, 0, "O2")
                    End If
                    '------------
                        Call DoDomar(UserIndex, ci)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||No podes domar a esa criatura." & "´" & FontTypeNames.FONTTYPE_info)
                    End If
          Else
                 Call SendData(ToIndex, UserIndex, 0, "M6")
          End If
          
        Case FundirMetal
    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

    
    'pluto:2.14---------------------------
            auxind = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
            If auxind > 0 Then
                wpaux.Map = UserList(UserIndex).Pos.Map
                wpaux.X = X
                wpaux.Y = Y
                If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "L2")
                    Exit Sub
                End If
        End If
     '------------------------------
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            
            If UserList(UserIndex).flags.TargetObj > 0 Then
                If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = OBJTYPE_FRAGUA Then
                     'pluto:6.2-------
                    If UserList(UserIndex).flags.Macreanda = 0 Then
                    UserList(UserIndex).flags.Macreanda = 4
                    UserList(UserIndex).Counters.Macrear = 2000
                    'UserList(UserIndex).flags.Macreando = wpaux
                    Call SendData(ToIndex, UserIndex, 0, "O2")
                    End If
                    '------------
                    Call FundirMineral(UserIndex)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & "´" & FontTypeNames.FONTTYPE_info)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & "´" & FontTypeNames.FONTTYPE_info)
            End If
            
        Case Herreria
            
            'pluto:2.14---------------------------
            auxind = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
            If auxind > 0 Then
                wpaux.Map = UserList(UserIndex).Pos.Map
                wpaux.X = X
                wpaux.Y = Y
                If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "L2")
                    Exit Sub
                End If
                Else
                Exit Sub
        End If
     '------------------------------
            
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            
            If UserList(UserIndex).flags.TargetObj > 0 Then
                If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = OBJTYPE_YUNQUE Then
                    Call EnivarArmasConstruibles(UserIndex)
                    Call EnivarArmadurasConstruibles(UserIndex)
                    Call SendData2(ToIndex, UserIndex, 0, 12)
                'pluto:2.7.0
                UserList(UserIndex).flags.TargetObj = 0
                 
                 Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & "´" & FontTypeNames.FONTTYPE_info)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & "´" & FontTypeNames.FONTTYPE_info)
            End If
            
        End Select
        
        UserList(UserIndex).flags.PuedeTrabajar = 0
        Exit Sub
    Case "CIG"
        rdata = Right$(rdata, Len(rdata) - 3)
        X = Guilds.Count
'pluto:2.4-->envia cero la reputacion----!
If CreateGuild(UserList(UserIndex).Name, 0, UserIndex, rdata) Then
'If CreateGuild(UserList(userindex).name, UserList(userindex).Reputacion.Promedio, userindex, rdata) Then
            
            If X = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Felicidades has creado el primer clan de Argentum!!!." & "´" & FontTypeNames.FONTTYPE_info)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Felicidades has creado el clan numero " & X + 1 & " de Argentum!!!." & "´" & FontTypeNames.FONTTYPE_info)
            End If
            'pluto:6.0A
            NameClan(X) = UserList(UserIndex).GuildInfo.GuildName
Dim oGuild As cGuild
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
oGuild.Nivel = 1
Call SaveGuildsDB
        End If
        
        Exit Sub

'pluto:2.4
Case "BYB"
rdata = Right$(rdata, Len(rdata) - 3)
Dim b As Integer
tName = ReadField(1, rdata, 44)
Tindex = NameIndex(tName & "$")
b = val(ReadField(2, rdata, 44))
If b > 4999 Then b = 4999
'If UserList(tIndex).GuildInfo.FundoClan = 1 Then b = 5000

Call WriteVar(CharPath & Left$(tName, 1) & "\" & tName & ".chr", "GUILD", "GuildPts", val(b))
If Tindex <= 0 Then Exit Sub
If UserList(Tindex).GuildInfo.FundoClan = 1 Then b = 5000
UserList(Tindex).GuildInfo.GuildPoints = b
Exit Sub

End Select

'----------------------------------------------------------------------
'----------------------------------------------------------------------

Select Case UCase$(Left$(rdata, 4))
    
        'pluto:2.17
    Case "ATRI"
        Call EnviarAtrib(UserIndex)
        Exit Sub
    Case "FAMA"
        Call EnviarFama(UserIndex)
        Exit Sub
    Case "ESTA"
    Call SendESTADISTICAS(UserIndex)
    Exit Sub
    Case "ESKI"
        Call EnviarSkills(UserIndex)
        Exit Sub
    
    
    
    Case "INFS" 'Informacion del hechizo
            rdata = Right$(rdata, Len(rdata) - 4)
            If val(rdata) > 0 And val(rdata) < MAXUSERHECHIZOS + 1 Then
                Dim H As Integer
                H = UserList(UserIndex).Stats.UserHechizos(val(rdata))
                If H > 0 And H < NumeroHechizos + 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & "´" & FontTypeNames.FONTTYPE_info)
                    Call SendData(ToIndex, UserIndex, 0, "||Nombre:" & Hechizos(H).Nombre & "´" & FontTypeNames.FONTTYPE_info)
                    Call SendData(ToIndex, UserIndex, 0, "||Descripcion:" & Hechizos(H).Desc & "´" & FontTypeNames.FONTTYPE_info)
                    Call SendData(ToIndex, UserIndex, 0, "||Skill requerido: " & Hechizos(H).MinSkill & " de magia." & "´" & FontTypeNames.FONTTYPE_info)
                    Call SendData(ToIndex, UserIndex, 0, "||Mana necesario: " & Hechizos(H).ManaRequerido & "´" & FontTypeNames.FONTTYPE_info)
                    Call SendData(ToIndex, UserIndex, 0, "||%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%" & "´" & FontTypeNames.FONTTYPE_info)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo.!" & "´" & FontTypeNames.FONTTYPE_info)
            End If
            Exit Sub
   'PLUTO:2.17
   Case "NMAS"
   rdata = Right$(rdata, Len(rdata) - 4)
   UserList(UserIndex).Montura.Nombre(val(ReadField(2, rdata, 44))) = ReadField(1, rdata, 44)
   Exit Sub
   
   Case "EQUI"
            If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 4)
            'PLUTO:14-3-04
            'Dim uh As Integer
            uh = val(ReadField(3, rdata, 44))
            If ReadField(2, rdata, 44) = "O" Then
            rdata = Left$(rdata, Len(rdata) - 1)
            Else
            Call LogCasino("Jugador:" & UserList(UserIndex).Name & " entró con cliente modificado. (A)" & "Ip: " & UserList(UserIndex).ip)
            Call SendData(ToAdmins, UserIndex, 0, "|| Detectado Cliente Modificado en " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_talk)
            End If
'pluto:2.4.5

'Dim kk As Integer
'kk = MinutosOnline - UserList(UserIndex).ShTime
'If uh > kk + 2 Then
'End If
'If uh > kk + Int(kk / 30) + 2 Then
'Call SendData(ToAdmins, UserIndex, 0, "|| Posible SH en " & UserList(UserIndex).Name & " --> " & uh & " // " & kk & "´" & FontTypeNames.FONTTYPE_talk)
'Call LogCasino("Jugador:" & UserList(UserIndex).Name & " Ip: " & UserList(UserIndex).ip & " Time Sh: " & uh & " // " & kk)
'End If
            
            If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
                 If UserList(UserIndex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            Call EquiparInvItem(UserIndex, val(rdata))
            Exit Sub
 
    'PLUTO:2.15
     Case "NBEB"
    rdata = Right$(rdata, Len(rdata) - 4)
    Call ComprobarNombreBebe(ReadField(1, rdata, 44), UserIndex, ReadField(2, rdata, 44))
    'If ComprobarNombreBebe = True Then Nacimiento (UserIndex)
    Exit Sub
    '--------
    
    
    Case "SKSE" 'Modificar skills
        'Dim i As Integer
        Dim sumatoria As Integer
        Dim incremento As Integer
        rdata = Right$(rdata, Len(rdata) - 4)
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            
            If incremento < 0 Then
                'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPENAMES.FONTTYPE_INFO)
                Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                UserList(UserIndex).Stats.SkillPts = 0
                Call CloseUser(UserIndex)
                Exit Sub
            End If
            
            sumatoria = sumatoria + incremento
        Next i
        
        If sumatoria > UserList(UserIndex).Stats.SkillPts Then
            'UserList(UserIndex).Flags.AdministrativeBan = 1
            'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPENAMES.FONTTYPE_INFO)
            Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
            Call CloseUser(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - incremento
            UserList(UserIndex).Stats.UserSkills(i) = UserList(UserIndex).Stats.UserSkills(i) + incremento
            If UserList(UserIndex).Stats.UserSkills(i) > MAXSKILLPOINTS Then UserList(UserIndex).Stats.UserSkills(i) = MAXSKILLPOINTS
        Next i
        Call EnviarSkills(UserIndex)
        Exit Sub
    Case "ENTR" 'Entrena hombre!
        
        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
        
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 3 Then Exit Sub
        
        rdata = Right$(rdata, Len(rdata) - 4)
    
        'If NPCHostiles(UserList(UserIndex).Pos.Map) < 6 Then
        If Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas < MAXMASCOTASENTRENADOR Then
            'pluto:6.0A
            If val(rdata) > 0 And val(rdata) < 6 Then
                    Dim SpawnedNpc As Integer
                    SpawnedNpc = SpawnNpc(Npclist(UserList(UserIndex).flags.TargetNpc).Criaturas(val(rdata)).NpcIndex, Npclist(UserList(UserIndex).flags.TargetNpc).Pos, True, False)
                    'pluto:6.3 cambio <= por <
                    If SpawnedNpc < MAXNPCS Then
                        Npclist(SpawnedNpc).MaestroNpc = UserList(UserIndex).flags.TargetNpc
                        Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas = Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas + 1
                    End If
            End If
        Else
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°No puedo traer mas criaturas, mata las existentes!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        End If
        
        Exit Sub
    Case "COMP"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "L3")
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).flags.TargetNpc > 0 Then
               '¿El NPC puede comerciar?
               If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                   Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°No tengo ningun interes en comerciar.°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 5)
         'User compra el item del slot rdata
         Call NPCVentaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(UserIndex).flags.TargetNpc)
         Exit Sub
    '[KEVIN]*********************************************************************
    '------------------------------------------------------------------------------------
    Case "RETI"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "L3")
                   Exit Sub
         End If
         'pluto:6.5
         If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
         If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 30 Then
               rdata = Right(rdata, Len(rdata) - 5)
         'User retira el item del slot rdata
         Call UserRetiraItemClan(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
        Exit Sub
         End If
         '---------------------------------
         
         '¿El target es un NPC valido?
         If UserList(UserIndex).flags.TargetNpc > 0 Then
               '¿Es el banquero?
               If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 4 And Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 25 Then
                   Exit Sub
               End If
         Else
           Exit Sub
         End If

         rdata = Right(rdata, Len(rdata) - 5)
         'User retira el item del slot rdata
         Call UserRetiraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '-----------------------------------------------------------------------------------
    '[/KEVIN]****************************************************************************
    Case "VEND"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "L3")
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).flags.TargetNpc > 0 Then
               '¿El NPC puede comerciar?
               If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                   Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°No tengo ningun interes en comerciar.°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 5)
         'User compra el item del slot rdata
         Call NPCCompraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '[KEVIN]-------------------------------------------------------------------------
    '****************************************************************************************
    Case "DEPO"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "L3")
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).flags.TargetNpc > 0 Then
        'pluto:6.0A
         If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 30 Then
               rdata = Right(rdata, Len(rdata) - 5)
         'User retira el item del slot rdata
         Call UserDepositaItemClan(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
        Exit Sub
         End If
         '---------------------------------
               
               
               '¿El NPC puede comerciar?
               If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 4 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡No puedes soltar objetos en este NPC!" & "´" & FontTypeNames.FONTTYPE_info)
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right(rdata, Len(rdata) - 5)
         'User deposita el item del slot rdata
         Call UserDepositaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '****************************************************************************************
    '[/KEVIN]---------------------------------------------------------------------------------
  End Select

'''''''''''''''''''VC NICO
    Select Case UCase$(Left$(rdata, 15))
        Case "ZSUXKRRFZKWAIQZV"    ' VERIFICACION CLIENTES
            UserList(UserIndex).flags.ClienteValido = 1
            Exit Sub
    End Select

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Select Case UCase$(Left$(rdata, 5))
    Case "DEMSG"
        If UserList(UserIndex).flags.TargetObj > 0 Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Dim f As String, Titu As String, msg As String, f2 As String
        f = App.Path & "\foros\"
        f = f & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
        '[MerLiNz:5]
        Titu = "<" & UserList(UserIndex).Name & "> "
        Titu = Titu & ReadField(1, rdata, 176)
        '[\END]
        msg = ReadField(2, rdata, 176)
        Dim n2 As Integer, loopme As Integer
        If FileExist(f, vbNormal) Then
            Dim num As Integer
            num = val(GetVar(f, "INFO", "CantMSG"))
            If num > MAX_MENSAJES_FORO Then
                For loopme = 1 To num
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & loopme & ".for"
                Next
                Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
                num = 0
            End If
            n2 = FreeFile
            f2 = Left$(f, Len(f) - 4)
            f2 = f2 & num + 1 & ".for"
            Open f2 For Output As n2
            Print #n2, Titu
            Print #n2, msg
            Call WriteVar(f, "INFO", "CantMSG", num + 1)
        Else
            n2 = FreeFile
            f2 = Left$(f, Len(f) - 4)
            f2 = f2 & "1" & ".for"
            Open f2 For Output As n2
            Print #n2, Titu
            Print #n2, msg
            Call WriteVar(f, "INFO", "CantMSG", 1)
        End If
        Close #n2
        End If
        Exit Sub
  
End Select


'----------------------------------------------------------------------------
'----------------------------------------------------------------------------

Select Case UCase$(Left$(rdata, 6))
 
    Case "DESCOD" 'Informacion del hechizo
            rdata = Right$(rdata, Len(rdata) - 6)
            Call UpdateCodexAndDesc(rdata, UserIndex)
            Exit Sub
 
 End Select

'[Alejo]
Select Case UCase$(Left$(rdata, 7))
Case "OFRECER"
        rdata = Right$(rdata, Len(rdata) - 7)
        Arg1 = ReadField(1, rdata, Asc(","))
        Arg2 = ReadField(2, rdata, Asc(","))

        If val(Arg1) <= 0 Or val(Arg2) <= 0 Or UserList(UserIndex).ComUsu.DestUsu <= 0 Then
            Exit Sub
        End If
         'pluto:6.3---------------
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Montura > 0 Or UserList(UserIndex).flags.Montura > 0 Then
        Call FinComerciarUsu(UserIndex)
        Exit Sub
        End If
       '---------------------
        
        'pluto:2.9.0 esta comerciando
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Comerciando = True And UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
        Call SendData(ToIndex, UserIndex, 0, "||Ya está comerciando con otro user." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
        Call FinComerciarUsu(UserIndex)
        Exit Sub
        End If
        
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged = False Then
            'sigue vivo el usuario ?
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        Else
            'esta vivo ?
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Muerto = 1 Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            End If
            '//Tiene la cantidad que ofrece ??//'
            If val(Arg1) = FLAGORO Then
                'oro
                If val(Arg2) > UserList(UserIndex).Stats.GLD Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & "´" & FontTypeNames.FONTTYPE_talk)
                    Exit Sub
                End If
            Else
                'inventario
                If val(Arg2) > UserList(UserIndex).Invent.Object(val(Arg1)).Amount Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & "´" & FontTypeNames.FONTTYPE_talk)
                    Exit Sub
                End If
            End If
            UserList(UserIndex).ComUsu.Objeto = val(Arg1)
            UserList(UserIndex).ComUsu.Cant = val(Arg2)
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
                'Es el primero que ofrece algo ?
                Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
                UserList(UserList(UserIndex).ComUsu.DestUsu).flags.TargetUser = UserIndex
            Else
                '[CORREGIDO]
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                    'NO NO NO vos te estas pasando de listo...
                    UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " HA CAMBIADO SU OFERTA!!." & "´" & FontTypeNames.FONTTYPE_talk)
                'Call SendData(ToIndex, UserList(userindex).ComUsu.DestUsu, 0, "!!" & " CUIDADO!! El otro jugador ha cambiado su oferta, comprueba bién lo que te está ofreciendo antes de aceptarla." & ENDC)
                'Call SendData2(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, 43, "CUIDADO HA CAMBIADO SU OFERTA")
                End If
                '[/CORREGIDO]
                'Es la ofrenda de respuesta :)
                Call EnviarObjetoTransaccion(UserList(UserIndex).ComUsu.DestUsu)
            End If
        End If
        Exit Sub
End Select
'[/Alejo]



'----------------------------------------------------------------------------
'----------------------------------------------------------------------------


Select Case UCase$(Left$(rdata, 8))
   
      
Case "ACEPPEAT"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptPeaceOffer(UserIndex, rdata)
        Exit Sub
    Case "PEACEOFF"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call RecievePeaceOffer(UserIndex, rdata)
        Exit Sub
    Case "PEACEDET"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeaceRequest(UserIndex, rdata)
        Exit Sub
    Case "ENVCOMEN"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeticion(UserIndex, rdata)
        Exit Sub
    Case "ENVPROPP"
        Call SendPeacePropositions(UserIndex)
        Exit Sub
    Case "DECGUERR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareWar(UserIndex, rdata)
        Exit Sub
    Case "DECALIAD"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareAllie(UserIndex, rdata)
        Exit Sub
    Case "NEWWEBSI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SetNewURL(UserIndex, rdata)
        Exit Sub
    Case "ACEPTARI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptClanMember(UserIndex, rdata)
        Exit Sub
    Case "RECHAZAR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DenyRequest(UserIndex, rdata)
        Exit Sub
    Case "ECHARCLA"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call EacharMember(UserIndex, rdata)
        Exit Sub
 Case "ACTGNEWS"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call UpdateGuildNews(rdata, UserIndex)
        Exit Sub
    Case "1HRINFO<"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendCharInfo(rdata, UserIndex)
        Exit Sub
    Case "NEWWLOGO"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SetNewEmblema(UserIndex, rdata)
        Exit Sub
End Select


'----------------------------------------------------------------------------
'----------------------------------------------------------------------------


Select Case UCase$(Left$(rdata, 9))
    Case "SOLICITUD"
         rdata = Right$(rdata, Len(rdata) - 9)
'pluto:2.20--------
Dim ah As Integer
'UserList(UserIndex).GuildInfo.ClanesParticipo = 11
ah = (10 + Int(UserList(UserIndex).Mision.numero / 20) - UserList(UserIndex).GuildInfo.ClanesParticipo)
If ah < 1 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes entrar en más clanes. Realizando las DraG Quest en el NpcQuest puedes ganar una solicitud adicional por cada 20 Quest." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'------------------
         Call SolicitudIngresoClan(UserIndex, rdata)
         Exit Sub

End Select


'----------------------------------------------------------------------------
'----------------------------------------------------------------------------

Select Case UCase$(Left$(rdata, 11))
   
  Case "CLANDETAILS"
        rdata = Right$(rdata, Len(rdata) - 11)
        Call SendGuildDetails(UserIndex, rdata)
        Exit Sub
End Select

'----------------------------------------------------------------------------
'----------------------------------------------------------------------------




       'pluto:2.8.0
If UCase$(Left$(rdata, 4)) = "BOLL" Then
    rdata = Right$(rdata, Len(rdata) - 4)
    Dim nIndex As Integer
    'nindex = ReadField(2, rdata, 44)
    Dim cochi As Integer
    cochi = RandomNumber(1, 100)
    If cochi > 50 Then Call MoveNPCChar(Balon, ReadField(1, rdata, 44))
    Exit Sub
End If
'PLUTO:6.0a
If Left$(rdata, 3) = "LIX" Then
rdata = Right$(rdata, Len(rdata) - 3)
Dim Tipi As Byte
Dim Seleci As Byte
Tipi = ReadField(1, rdata, 44)
Tipi = Tipi + 1
Tindex = ReadField(2, rdata, 44)
Seleci = ReadField(3, rdata, 44)
If UserList(Tindex).Montura.Libres(Seleci) <= 0 Then Exit Sub

UserList(Tindex).Montura.Libres(Seleci) = UserList(Tindex).Montura.Libres(Seleci) - 1

Select Case Tipi
Case 1
UserList(Tindex).Montura.AtCuerpo(Seleci) = UserList(Tindex).Montura.AtCuerpo(Seleci) + 1
Case 2
UserList(Tindex).Montura.Defcuerpo(Seleci) = UserList(Tindex).Montura.Defcuerpo(Seleci) + 1
Case 3
UserList(Tindex).Montura.AtFlechas(Seleci) = UserList(Tindex).Montura.AtFlechas(Seleci) + 1
Case 4
UserList(Tindex).Montura.DefFlechas(Seleci) = UserList(Tindex).Montura.DefFlechas(Seleci) + 1
Case 5
UserList(Tindex).Montura.AtMagico(Seleci) = UserList(Tindex).Montura.AtMagico(Seleci) + 1
Case 6
UserList(Tindex).Montura.DefMagico(Seleci) = UserList(Tindex).Montura.DefMagico(Seleci) + 1
Case 7
UserList(Tindex).Montura.Evasion(Seleci) = UserList(Tindex).Montura.Evasion(Seleci) + 1
End Select

Exit Sub
End If

'pluto:2.4.7
If Left$(rdata, 3) = "BO5" Then
rdata = Right$(rdata, Len(rdata) - 3)
Tindex = ReadField(1, rdata, 44)
'pluto:6.7
If UserList(Tindex).flags.Privilegios = 0 Then Exit Sub
'pluto:2.15
Call LogCasino("Jugador:" & UserList(Tindex).Name & " hizo foto " & "Ip: " & UserList(Tindex).ip)
Call SendData(ToGM, Tindex, 0, "|| Foto desde la ip: " & UserList(Tindex).ip & "´" & FontTypeNames.FONTTYPE_info)
Call SendData2(ToIndex, Tindex, 0, 84, rdata)
Exit Sub
End If
'-----------------
'pluto:2.5.0
If Left$(rdata, 3) = "BO6" Then
'quitar esto
Exit Sub
rdata = Right$(rdata, Len(rdata) - 3)
Call LogInitModificados(rdata)
Exit Sub
End If
'pluto:2.8.0
If Left$(rdata, 3) = "BO8" Then
rdata = Right$(rdata, Len(rdata) - 3)
Tindex = ReadField(1, rdata, 44)
'pluto:6.7
If UserList(Tindex).flags.Privilegios = 0 Then Exit Sub
Call SendData2(ToIndex, Tindex, 0, 86, rdata)
Exit Sub
End If
'pluto:2.8.0
If Left$(rdata, 3) = "BO9" Then
rdata = Right$(rdata, Len(rdata) - 3)
Tindex = ReadField(1, rdata, 44)
'pluto:6.7
If UserList(Tindex).flags.Privilegios = 0 Then Exit Sub
Call SendData2(ToIndex, Tindex, 0, 88, rdata)
Exit Sub
End If

'pluto:6.2
If Left$(rdata, 3) = "XO1" Then
rdata = Right$(rdata, Len(rdata) - 3)
Call SendData(ToGM, 0, 0, "|| Conexión Correcta." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToGM, 0, 0, "|| Realizando Foto." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "S8")
Exit Sub
End If
'pluto:6.2
If Left$(rdata, 3) = "XO2" Then
rdata = Right$(rdata, Len(rdata) - 3)
Call SendData(ToGM, 0, 0, "|| Conexión Incorrecta!!." & "´" & FontTypeNames.FONTTYPE_info)
'Call SendData(ToIndex, UserIndex, 0, "S8")
Exit Sub
End If
'pluto:2.13
If Left$(rdata, 3) = "TA1" Then
rdata = Right$(rdata, Len(rdata) - 3)
Tindex = ReadField(1, rdata, 44)
'pluto:6.7
If UserList(Tindex).flags.Privilegios = 0 Then Exit Sub

Call SendData(ToIndex, Tindex, 0, "Z1" & rdata)
Exit Sub
End If

'pluto:2.9.0 'Se crea un Torneo
If Left$(rdata, 3) = "TO2" Then
rdata = Right$(rdata, Len(rdata) - 3)
Call CrearTorneo(rdata)
Call EnviarTorneo(UserIndex)
Exit Sub
End If
'pluto:2.9.0 'Se participa Torneo
If Left$(rdata, 3) = "TO3" Then
'rdata = Right$(rdata, Len(rdata) - 3)
Call ParticipaTorneo(UserList(UserIndex).Name)
Exit Sub
End If


Exit Sub
ErrorComandoPj:
 Call LogError("TCP1. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).Name & "UI:" & UserIndex & " N: " & Err.number & " D: " & Err.Description)
 Call CloseSocket(UserIndex)


End Sub
