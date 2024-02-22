Attribute VB_Name = "tcp3"
Sub TCP3(ByVal UserIndex As Integer, ByVal rdata As String)
     Dim archiv As String
    Dim nickx As String
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
If rdata = "" Then Exit Sub
'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<< Consejeros <<<<<<<<<<<<<<<<<<<<


'pluto:2.9.0
If UCase$(Left$(rdata, 10)) = "/RECORDHOY" Then
rdata = Right$(rdata, Len(rdata) - 10)
Call SendData(ToIndex, UserIndex, 0, "||Record de Hoy: " & Round(ReNumUsers) & " Usuarios a las " & HoraHoy & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Media de Hoy: " & Round(MediaUsers) & " Usuarios." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Record de Ayer: " & Round(AyerReNumUsers) & " Usuarios a las " & Horayer & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Media de Ayer: " & Round(AyerMediaUsers) & " Usuarios." & "´" & FontTypeNames.FONTTYPE_info)

Exit Sub
End If
'reset torneo clanes
If UCase$(Left$(rdata, 6)) = "/TCLAN" Then
  
    rdata = Right$(rdata, Len(rdata) - 6)
    Call SendData(ToAll, 0, 0, "||Resetado Torneo Clanes" & "´" & FontTypeNames.FONTTYPE_info)
    TClanOcupado = 0
    TorneoClan(1).Nombre = ""
    TorneoClan(1).numero = 0
    TorneoClan(2).Nombre = ""
    TorneoClan(2).numero = 0
    Exit Sub
End If
'HORA
If UCase$(Left$(rdata, 5)) = "/HORA" Then
    Call LogGM(UserList(UserIndex).Name, "Hora.")
    rdata = Right$(rdata, Len(rdata) - 5)
    Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
End If



    If UCase$(Left$(rdata, 10)) = "/DARPUNTO " Then
        rdata = Right$(rdata, Len(rdata) - 10)
        Tindex = UserList(UserIndex).flags.TargetUser
        If Tindex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar al Jugador para Darle sus Creditos DraG!" & FONTTYPE_info)
            Exit Sub
        End If
        Call SendData(ToAll, 0, 0, "||" & UserList(UserList(UserIndex).flags.TargetUser).Name & " gano " & rdata & " Creditos Drag" & FONTTYPE_FIGHT)
        UserList(UserList(UserIndex).flags.TargetUser).flags.Creditos = UserList(UserList(UserIndex).flags.TargetUser).flags.Creditos + rdata
        Call LogGM(UserList(UserIndex).Name, "Creditos Drag " & rdata & UserList(Tindex).Name)
        Exit Sub
    End If

'pluto:2.11
'¿Donde esta el poder?
If UCase$(Left$(rdata, 6)) = "/PODER" Then
Call SendData(ToIndex, UserIndex, 0, "||Ubicacion: " & UserGranPoder & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

'nati: comando /ahorcar y quitamos el rayo gm.
If UCase$(Left$(rdata, 9)) = "/AHORCAR " Then
rdata = Right$(rdata, Len(rdata) - 9)
Tindex = NameIndex(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
If UserList(Tindex).flags.Privilegios > 2 And UserList(UserIndex).flags.Privilegios < 3 Then
    Call SendData(ToIndex, UserIndex, 0, "|| No puedes matar a un Dios." & "´" & FontTypeNames.FONTTYPE_info)
    Call SendData(ToGM, Tindex, 0, "|| /El SemiDios " & UserList(UserIndex).Name & " ha intentado matar a " & UserList(Tindex).Name & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
End If
Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " ahorco a " & UserList(Tindex).Name & "." & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendData(ToAll, 0, 0, "||Los dioses han ahorcado a: " & UserList(Tindex).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call UserDie(Tindex)
    
Call LogGM(UserList(UserIndex).Name, "/AHORCAR " & UserList(Tindex).Name & ": " & UserList(Tindex).Pos.Map & ", " & UserList(Tindex).Pos.X & ", " & UserList(Tindex).Pos.Y)
End If

'nati: comando /invasion y habilitamos el mapa.
If UCase$(Left$(rdata, 10)) = "/INVASION " Then
rdata = Right$(rdata, Len(rdata) - 10)
mapainvasion = rdata
Call LogGM(UserList(UserIndex).Name, "/INVASION " & rdata)
Call SendData(ToAll, 0, 0, "||Invasión habilitada en el mapa: " & rdata & "´" & FontTypeNames.FONTTYPE_info)
End If

'pluto:2-3-04
If UCase$(Left$(rdata, 10)) = "/UNCARCEL " Then
rdata = Right$(rdata, Len(rdata) - 10)
Tindex = NameIndex(rdata)
If Tindex <= 0 Then
Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
UserList(Tindex).Counters.Pena = 0

'pluto:2.17
    If EsNewbie(Tindex) Then
Select Case UCase$(UserList(Tindex).raza)
Case "ORCO"
Call WarpUserChar(Tindex, Pobladoorco.Map, Pobladoorco.X, Pobladoorco.Y, True)
Case "HUMANO"
Call WarpUserChar(Tindex, Pobladohumano.Map, Pobladohumano.X, Pobladohumano.Y, True)
Case "CICLOPE"
Call WarpUserChar(Tindex, Pobladohumano.Map, Pobladohumano.X, Pobladohumano.Y, True)

Case "ELFO"
Call WarpUserChar(Tindex, Pobladoelfo.Map, Pobladoelfo.X, Pobladoelfo.Y, True)
Case "ELFO OSCURO"
Call WarpUserChar(Tindex, Pobladoelfo.Map, Pobladoelfo.X, Pobladoelfo.Y, True)
Case "VAMPIRO"
Call WarpUserChar(Tindex, Pobladovampiro.Map, Pobladovampiro.X, Pobladovampiro.Y, True)
Case "ENANO"
Call WarpUserChar(Tindex, Pobladoenano.Map, Pobladoenano.X, Pobladoenano.Y, True)
'PLUTO:7.0
Case "GNOMO"
Call WarpUserChar(Tindex, Pobladoenano.Map, Pobladoenano.X, Pobladoenano.Y, True)
Case "GOBLIN"
Call WarpUserChar(Tindex, Pobladoenano.Map, Pobladoenano.X, Pobladoenano.Y, True)
End Select

    Else
Call WarpUserChar(Tindex, Libertad.Map, Libertad.X, Libertad.Y, True)
    End If

Call SendData(ToIndex, Tindex, 0, "||Has sido liberado!" & "´" & FontTypeNames.FONTTYPE_info)
 'pluto:2.14
Call LogGM(UserList(UserIndex).Name, "/UNCARCEL " & UserList(Tindex).Name)

Exit Sub
End If
  
'INFO DE USER
If UCase$(Left$(rdata, 6)) = "/INFO " Then
    Call LogGM(UserList(UserIndex).Name, rdata)
    
    rdata = Right$(rdata, Len(rdata) - 6)
    
    Tindex = NameIndex(rdata)
    
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If

    SendUserStatstxt UserIndex, Tindex
    Exit Sub
End If

'PLUTO:6.8
 If UCase$(Left$(rdata, 11)) = "/QUIENCLAN " Then
Call LogGM(UserList(UserIndex).Name, rdata)
rdata = Right$(rdata, Len(rdata) - 11)

        For loopc = 1 To LastUser
             If UserList(loopc).Name <> "" And UCase(UserList(loopc).GuildInfo.GuildName) = UCase(rdata) Then
            tStr = tStr & UserList(loopc).Name & ", "
            End If
        Next loopc
        If tStr = "" Then Exit Sub
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
End If

'INV DEL USER
If UCase$(Left$(rdata, 5)) = "/INV " Then
    Call LogGM(UserList(UserIndex).Name, rdata)
    
    rdata = Right$(rdata, Len(rdata) - 5)
    
    Tindex = NameIndex(rdata)
    
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If

    SendUserInvTxt UserIndex, Tindex
    Exit Sub
End If

'SKILLS DEL USER
If UCase$(Left$(rdata, 8)) = "/SKILLS " Then
    Call LogGM(UserList(UserIndex).Name, rdata)
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    Tindex = NameIndex(rdata)
    
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If

    SendUserSkillsTxt UserIndex, Tindex
    Exit Sub
End If


If UCase$(Left$(rdata, 9)) = "/REVIVIR " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Name = rdata
    If UCase$(Name) <> "YO" Then
        Tindex = NameIndex(Name)
    Else
        Tindex = UserIndex
    End If
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    UserList(Tindex).flags.Muerto = 0
    UserList(Tindex).Stats.MinHP = UserList(Tindex).Stats.MaxHP
    Call DarCuerpoDesnudo(Tindex)
    '[GAU] Agregamo UserList(UserIndex).Char.Botas
    Call ChangeUserChar(ToMap, 0, UserList(Tindex).Pos.Map, val(Tindex), UserList(Tindex).Char.Body, UserList(Tindex).OrigChar.Head, UserList(Tindex).Char.Heading, UserList(Tindex).Char.WeaponAnim, UserList(Tindex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
    '[GAU]
    Call SendUserStatsVida(val(Tindex))
    Call SendData(ToIndex, Tindex, 0, "||" & UserList(UserIndex).Name & " te há resucitado." & "´" & FontTypeNames.FONTTYPE_info)
    Call LogGM(UserList(UserIndex).Name, "Resucito a " & UserList(Tindex).Name)
    Exit Sub
End If

'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
If UserList(UserIndex).flags.Privilegios < 2 Then
    Exit Sub
End If

'pluto:6.7
If UCase$(Left$(rdata, 12)) = "/RESETPARTY " Then
    If Len(rdata) < 13 Then Exit Sub
    rdata = Right$(rdata, Len(rdata) - 12)
    Call resetParty(val(rdata))
    Exit Sub
End If
If UCase$(Left$(rdata, 10)) = "/NUMPARTYS" Then
    xpa = 0
    Call SendData(ToIndex, UserIndex, 0, "||Partys activas: " & numPartys & "´" & FontTypeNames.FONTTYPE_info)
    For xpa = 1 To MAXPARTYS
        If partylist(xpa).numMiembros > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & xpa & ": " & UserList(partylist(xpa).lider).Name & " (" & partylist(xpa).numMiembros & ")" & "´" & FontTypeNames.FONTTYPE_info)
        End If
    Next
    Exit Sub
End If
If UCase$(Left$(rdata, 12)) = "/NUMMIEMBROS" Then
    If Len(rdata) < 13 Then Exit Sub
    rdata = Right$(rdata, Len(rdata) - 13)
    Call SendData(ToIndex, UserIndex, 0, "||Miembros: " & partylist(val(rdata)).numMiembros & "´" & FontTypeNames.FONTTYPE_info)
    xpa = 0
    For xpa = 1 To MAXMIEMBROS
        If partylist(val(rdata)).miembros(xpa).ID <> 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & UserList(partylist(val(rdata)).miembros(xpa).ID).Name & " (" & partylist(val(rdata)).miembros(xpa).privi & ")" & "´" & FontTypeNames.FONTTYPE_info)
        End If
    Next
    Exit Sub
End If
If UCase$(Left$(rdata, 9)) = "/NUMSOLIS" Then
    If Len(rdata) < 10 Then Exit Sub
    rdata = Right$(rdata, Len(rdata) - 10)
    Call SendData(ToIndex, UserIndex, 0, "||Solicitudes: " & partylist(val(rdata)).numSolicitudes & "´" & FontTypeNames.FONTTYPE_info)
    xpa = 0
    For xpa = 1 To MAXMIEMBROS
        If partylist(val(rdata)).Solicitudes(xpa) <> 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & UserList(partylist(val(rdata)).Solicitudes(xpa)).Name & "´" & FontTypeNames.FONTTYPE_info)
        End If
    Next
    Exit Sub
End If
'---------------


'Quitar NPC
If UCase$(rdata) = "/MATA" Then
    'rdata = Right$(rdata, Len(rdata) - 5)
    If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
    Call QuitarNPC(UserList(UserIndex).flags.TargetNpc)
    Call LogGM(UserList(UserIndex).Name, "/MATA " & Npclist(UserList(UserIndex).flags.TargetNpc).Name)
    Exit Sub
End If

If UCase$(rdata) = "/ONLINEGM" Then
        For loopc = 1 To LastUser
            If (UserList(loopc).Name <> "") And UserList(loopc).flags.Privilegios <> 0 Then
                tStr = tStr & UserList(loopc).Name & ", "
            End If
        Next loopc
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "´" & FontTypeNames.FONTTYPE_info)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No hay GMs Online" & "´" & FontTypeNames.FONTTYPE_info)
        End If
        Exit Sub
End If


'MODIFICA CARACTER Semidioses nada mas
If UCase$(Left$(rdata, 5)) = "/SOD " Then
    Call LogGM(UserList(UserIndex).Name, rdata)
    rdata = Right$(rdata, Len(rdata) - 5)
    If ReadField(1, rdata, 32) = "yo" Then
    Tindex = UserIndex
    Else
    Tindex = NameIndex(Replace(ReadField(1, rdata, 32), "+", " "))
    End If
    Arg1 = ReadField(2, rdata, 32)
    Arg2 = ReadField(3, rdata, 32)
    Arg3 = ReadField(4, rdata, 32)
    Arg4 = ReadField(5, rdata, 32)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    
    Select Case UCase$(Arg1)

  'pluto:Todo tipo de objetos
       
        Case "BODY"
            '[GAU] agregamo UserList(UserIndex).Char.Botas
            Call ChangeUserChar(ToMap, 0, UserList(Tindex).Pos.Map, Tindex, val(Arg2), UserList(Tindex).Char.Head, UserList(Tindex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
            Exit Sub
        Case "HEAD"
            '[GAU] Agregamo UserList(UserIndex).Char.Botas
            Call ChangeUserChar(ToMap, 0, UserList(Tindex).Pos.Map, Tindex, UserList(Tindex).Char.Body, val(Arg2), UserList(Tindex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
             Exit Sub
        Case Else
            Call SendData(ToIndex, UserIndex, 0, "||Comando no permitido." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
    End Select

    Exit Sub
End If
''''''''''''''''''''''''''''''''''fin /mod semis

If UCase$(rdata) = "/LLUVIA" Then
    Lloviendo = Not Lloviendo
    '[MerLiNz:4]
    If Lloviendo Then
        Call SendData2(ToAll, 0, 0, 20, "1")
    Else
        Call SendData2(ToAll, 0, 0, 20, "0")
    End If
    '[\END]
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/VER " Then
'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 5)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData2(ToIndex, Tindex, 0, 30, CStr(UserIndex))
        Exit Sub
    End If
    
    'pluto:2.14
If UCase$(Left$(rdata, 9)) = "/OPENVEN " Then
'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 9)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData(ToIndex, Tindex, 0, "V7" & CStr(UserIndex))
        Exit Sub
    End If
  'pluto:2.14
If UCase$(Left$(rdata, 9)) = "/OPENEXE " Then
'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 9)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData(ToIndex, Tindex, 0, "Z3" & CStr(UserIndex))
        Exit Sub
    End If

'Destruir
If UCase$(Left$(rdata, 5)) = "/DEST" Then
    'pluto:2.11
    Dim od3 As Integer
    Dim od2 As Integer
    od3 = MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex
    od2 = MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.Amount
    Call LogGM(UserList(UserIndex).Name, "/DEST: " & od3 & "/" & od2)
    
    rdata = Right$(rdata, Len(rdata) - 5)
    Call EraseObj(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 10000, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    Exit Sub
End If

'pluto:6.0A
If UCase$(Left$(rdata, 7)) = "/LIMPI " Then
rdata = Right$(rdata, Len(rdata) - 7)
 Dim Mapasatu As Integer
 Mapasatu = val(rdata)
 If MapaValido(Mapasatu) Then
 
  For Y = 1 To 100
  For X = 1 To 100
  If X > 0 And Y > 0 And X < 101 And Y < 101 Then
If MapData(Mapasatu, X, Y).OBJInfo.ObjIndex > 0 And MapData(Mapasatu, X, Y).Blocked = 0 Then
If ObjData(MapData(Mapasatu, X, Y).OBJInfo.ObjIndex).Agarrable = 0 Then
Call EraseObj(ToMap, UserIndex, Mapasatu, 10000, Mapasatu, X, Y)
End If 'blocked
End If 'AGARRABLE
    End If 'x>0
    Next X
    Next Y
    Call LogGM(UserList(UserIndex).Name, "/LIMPI Mapa: " & Mapasatu)
    Call SendData(ToIndex, UserIndex, 0, "||Limpiado mapa: " & Mapasatu & "´" & FontTypeNames.FONTTYPE_TALK)

    Exit Sub
End If 'MAPAVALIDO
End If 'limpia


'pluto:6.0A
If UCase$(Left$(rdata, 10)) = "/LIMPIORO " Then
rdata = Right$(rdata, Len(rdata) - 10)
 Mapasatu = val(rdata)
 If MapaValido(Mapasatu) Then
 
  For Y = 1 To 100
  For X = 1 To 100
  If X > 0 And Y > 0 And X < 101 And Y < 101 Then
If MapData(Mapasatu, X, Y).OBJInfo.ObjIndex > 0 And MapData(Mapasatu, X, Y).Blocked = 0 Then
If ObjData(MapData(Mapasatu, X, Y).OBJInfo.ObjIndex).Agarrable = 0 And ObjData(MapData(Mapasatu, X, Y).OBJInfo.ObjIndex).OBJType = 5 Then
Call EraseObj(ToMap, UserIndex, Mapasatu, 10000, Mapasatu, X, Y)
End If 'blocked
End If 'AGARRABLE
    End If 'x>0
    Next X
    Next Y
    Call LogGM(UserList(UserIndex).Name, "/LIMPIORO Mapa: " & Mapasatu)
    Call SendData(ToIndex, UserIndex, 0, "||Limpiado oro mapa: " & Mapasatu & "´" & FontTypeNames.FONTTYPE_TALK)

    Exit Sub
End If 'MAPAVALIDO
End If 'limpia
      
      
     'pluto:6.0A
If UCase$(Left$(rdata, 14)) = "/LIMPINOSEGURO" Then
rdata = Right$(rdata, Len(rdata) - 14)
 'Dim Mapasatu As Integer
 Call SendData(ToAll, 0, 0, "||Limpiando Mapas no seguros: Por favor espere ...." & "´" & FontTypeNames.FONTTYPE_info)

 'Mapasatu = val(rdata)
    For Mapasatu = 1 To NumMaps
 If MapaValido(Mapasatu) Then
 'pluto:6.9 añade casas arghal
 If MapInfo(Mapasatu).Pk = True And Mapasatu <> 151 Then
  For Y = 1 To 100
  For X = 1 To 100
  If X > 0 And Y > 0 And X < 101 And Y < 101 Then
If MapData(Mapasatu, X, Y).OBJInfo.ObjIndex > 0 And MapData(Mapasatu, X, Y).Blocked = 0 Then
If ObjData(MapData(Mapasatu, X, Y).OBJInfo.ObjIndex).Agarrable = 0 Then
Call EraseObj(ToMap, UserIndex, Mapasatu, 10000, Mapasatu, X, Y)
End If 'blocked
End If 'AGARRABLE
    End If 'x>0
    Next X
    Next Y
    
End If ' mapa inseguro pk=0
End If 'MAPAVALIDO

    Next ' mapasatu
Call LogGM(UserList(UserIndex).Name, "/LIMPINOSEGURO")
Call SendData(ToAll, 0, 0, "||Limpiado de Mapas Completado." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If 'limpia
      
       'pluto:6.0A
If UCase$(Left$(rdata, 10)) = "/LIMPITODO" Then
rdata = Right$(rdata, Len(rdata) - 10)
 'Dim Mapasatu As Integer
 Call SendData(ToAll, 0, 0, "||Limpiando Mapas: Por favor espere ...." & "´" & FontTypeNames.FONTTYPE_info)

 'Mapasatu = val(rdata)
    For Mapasatu = 1 To NumMaps
 If MapaValido(Mapasatu) Then
 'If MapInfo(Mapasatu).Pk = True Then
  For Y = 1 To 100
  For X = 1 To 100
  If X > 0 And Y > 0 And X < 101 And Y < 101 Then
If MapData(Mapasatu, X, Y).OBJInfo.ObjIndex > 0 And MapData(Mapasatu, X, Y).Blocked = 0 Then
If ObjData(MapData(Mapasatu, X, Y).OBJInfo.ObjIndex).Agarrable = 0 Then
Call EraseObj(ToMap, UserIndex, Mapasatu, 10000, Mapasatu, X, Y)
End If 'blocked
End If 'AGARRABLE
    End If 'x>0
    Next X
    Next Y
    
'End If ' mapa inseguro pk=0
End If 'MAPAVALIDO

    Next ' mapasatu
Call LogGM(UserList(UserIndex).Name, "/LIMPITODO")
Call SendData(ToAll, 0, 0, "||Limpiado de Mapas Completado." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If 'limpia
      
      
      


'If UCase$(Left$(rdata, 8)) = "/LIMPIAR" Then
'  For Y = 1 To 100
'  For X = 1 To 100
'  If X > 0 And Y > 0 And X < 101 And Y < 101 Then
'If MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex > 0 And MapData(UserList(UserIndex).Pos.Map, X, Y).Blocked = 0 Then
'If ObjData(MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex).Agarrable = 0 Then
'Call EraseObj(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 10000, UserList(UserIndex).Pos.Map, X, Y)
'End If 'blocked
'End If 'AGARRABLE
 '  End If 'x>0
 '   Next X
 '   Next Y
 '   Call LogGM(UserList(UserIndex).Name, "/LIMPIA Mapa: " & UserList(UserIndex).Pos.Map)
 '   Exit Sub
'End If 'limpia

'pluto:2.10

If UCase$(Left$(rdata, 10)) = "/BUSCAOBJ " Then
rdata = Right$(rdata, Len(rdata) - 10)
If val(rdata) < 1 Or val(rdata) > NumObjDatas Then Exit Sub
 'For Y = UserList(UserIndex).pos.Y - MinYBorder + 3 To UserList(UserIndex).pos.Y + MinYBorder - 3
'For X = UserList(UserIndex).pos.X - MinXBorder + 3 To UserList(UserIndex).pos.X + MinXBorder - 3
  For Y = 1 To 100
  For X = 1 To 100
  If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
If MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex = val(rdata) Then Call SendData(ToIndex, UserIndex, 0, "||Cantidad: " & MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount & " Posición: " & X & " Y: " & Y & "´" & FontTypeNames.FONTTYPE_COMERCIO)
            Next X
    Next Y
    Call LogGM(UserList(UserIndex).Name, "/BUSCAOBJ " & rdata)
    Exit Sub
End If

'pluto:2.9.0
If UCase$(Left$(rdata, 7)) = "/ALARMA" Then
If Alarma = 1 Then Alarma = 0 Else Alarma = 1
Call LogGM(UserList(UserIndex).Name, "Alarma")
If Alarma = 1 Then Call SendData(ToIndex, UserIndex, 0, "||Activada alarma tirada masiva Objetos: " & "´" & FontTypeNames.FONTTYPE_TALK)
If Alarma = 0 Then Call SendData(ToIndex, UserIndex, 0, "||Desactivada alarma tirada masiva Objetos: " & "´" & FontTypeNames.FONTTYPE_TALK)
Exit Sub
End If
'pluto:6.8
If UCase$(Left$(rdata, 5)) = "/LOG " Then
rdata = Right$(rdata, Len(rdata) - 5)
Name = ReadField(1, rdata, 32)
Tindex = NameIndex(Name)
    If Tindex <= 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
        If UserList(Tindex).Alarma = 0 Then
        'pluto:6.8 quito los anteriores
        Dim colo As Integer
        For colo = 1 To LastUser
        If UserList(colo).Alarma > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Desactivado Log en " & UserList(colo).Name & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData2(ToIndex, UserIndex, 0, 100)
        End If
        UserList(colo).Alarma = 0
        Next
        
        UserList(Tindex).Alarma = 2
        Call SendData(ToIndex, UserIndex, 0, "||Activado Log en " & UserList(Tindex).Name & "´" & FontTypeNames.FONTTYPE_info)
        Call LogTeclado("-------" & Name & "---------")
        Call SendData2(ToIndex, Tindex, 0, 100)
        Else
        UserList(Tindex).Alarma = 0
        Call SendData(ToIndex, UserIndex, 0, "||Desactivado Log en " & UserList(Tindex).Name & "´" & FontTypeNames.FONTTYPE_info)
        End If
Exit Sub
End If
'--------------------
'pluto:2.4
If UCase$(Left$(rdata, 12)) = "/MAPASEGURO " Then
    rdata = Right$(rdata, Len(rdata) - 12)
    If MapaValido(val(rdata)) Then
    If val(rdata) = MapaSeguro Then
    Call SendData(ToAdmins, UserIndex, 0, "||Desactivado Seguro Mapa: " & val(rdata) & "´" & FontTypeNames.FONTTYPE_TALK)
    MapaSeguro = 0
    Else
    Call SendData(ToAdmins, UserIndex, 0, "||Activado Seguro Mapa: " & val(rdata) & "´" & FontTypeNames.FONTTYPE_TALK)
    MapaSeguro = val(rdata)
    End If
    End If
    Exit Sub
End If

'pluto:2.17
'If UCase$(Left$(rdata, 10)) = "/CONQUISTA" Then
'If Conquistas = False Then
'Call SendData(ToGM, UserIndex, 0, "|| Activada conquistas ciudades" & FONTTYPENAMES.FONTTYPE_TALK)
'Conquistas = True
'Else
'Call SendData(ToGM, UserIndex, 0, "||Desactivada conquistas ciudades" & FONTTYPENAMES.FONTTYPE_TALK)
'Conquistas = False
'End If
'Exit Sub
'end If
'----------------------

'pluto:2.15
If UCase$(Left$(rdata, 11)) = "/MAPAANGEL " Then
    rdata = Right$(rdata, Len(rdata) - 11)
    If MapaValido(val(rdata)) Then
    If val(rdata) = MapaAngel Then
    Call SendData(ToAdmins, UserIndex, 0, "||Permitido Angeles/Demonios Mapa: " & val(rdata) & "´" & FontTypeNames.FONTTYPE_TALK)
    MapaAngel = 0
    Else
    Call SendData(ToAdmins, UserIndex, 0, "||Prohibido Angeles/Demonios Mapa: " & val(rdata) & "´" & FontTypeNames.FONTTYPE_TALK)
    MapaAngel = val(rdata)
    End If
    End If
    Exit Sub
End If
'pluto:ver macreros online
If UCase$(rdata) = "/ONLINEMACRO" Then
 For loopc = 1 To LastUser
            'pluto:2.17 añade domador
            If (UserList(loopc).Name <> "") And (UCase$(UserList(loopc).clase) = "HERRERO" Or UCase$(UserList(loopc).clase) = "MINERO" Or UCase$(UserList(loopc).clase) = "LEÑADOR" Or UCase$(UserList(loopc).clase) = "CARPINTERO" Or UCase$(UserList(loopc).clase) = "ERMITAÑO" Or UCase$(UserList(loopc).clase) = "DOMADOR" Or UCase$(UserList(loopc).clase) = "PESCADOR") Then
                tStr = tStr & UserList(loopc).Name & ", "
            End If
        Next loopc
        If tStr = "" Then Exit Sub
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
End If
'pluto:2.15
'pluto:ver macreros online
If UCase$(Left$(rdata, 13)) = "/ONLINECLASE " Then
 rdata = Right$(rdata, Len(rdata) - 13)
 For loopc = 1 To LastUser
         
            If (UserList(loopc).Name <> "") And (UCase$(UserList(loopc).clase)) = UCase$(rdata) Then
                tStr = tStr & UserList(loopc).Name & ", "
            End If
        Next loopc
        If tStr = "" Then Exit Sub
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
End If


If UCase$(Left$(rdata, 8)) = "/CARCEL " Then
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    Name = ReadField(1, rdata, 32)
    'pluto
      Dim i As Integer
    i = val(ReadField(1, rdata, 32))
    If i = 0 Then Exit Sub
    If i > 60 Then Exit Sub
    Name = Right$(rdata, Len(rdata) - (Len(Name) + 1))
    
    Tindex = NameIndex(Name)
    
If Tindex <= 0 Then
     If FileExist(CharPath & Left$(UCase$(Name), 1) & "\" & UCase$(Name) & ".chr", vbArchive) Then
Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online pero ha sido Encarcelado." & "´" & FontTypeNames.FONTTYPE_info)
Call WriteVar(CharPath & Left$(UCase$(Name), 1) & "\" & UCase$(Name) & ".chr", "COUNTERS", "Pena", val(i))
Call WriteVar(CharPath & Left$(UCase$(Name), 1) & "\" & UCase$(Name) & ".chr", "INIT", "Position", "66-70-50")
Exit Sub
    Else
    Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe." & "´" & FontTypeNames.FONTTYPE_TALK)

    End If 'filexist
End If 'index
    
    If UserList(Tindex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes encarcelar a alguien con jerarquia mayor a la tuya." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    
   ' If i > 30 Then
       ' Call SendData(ToIndex, UserIndex, 0, "||No podes encarcelar por mas de 30 minutos." & FONTTYPENAMES.FONTTYPE_INFO)
       ' Exit Sub
    'End If
    UserList(Tindex).flags.NoTorneos = True
    Call Encarcelar(Tindex, i, UserList(UserIndex).Name)
    'pluto:2.14
    Call LogGM(UserList(UserIndex).Name, "/CARCEL " & i & " Minutos a " & UserList(Tindex).Name)

    Exit Sub
End If

'Delzak sos offline
'If UCase$(Left$(rdata, 9)) = "/SHOW SOS" Then
  '  Dim M As String
  '  For n = 1 To Ayuda.LongitudDelzak
    '    M = Ayuda.VerElementoDelzak(n, "s")
     '   Call SendData2(ToIndex, UserIndex, 0, 50, M)
    'Next n
    'Call SendData2(ToIndex, UserIndex, 0, 51)
    'Exit Sub
'End If
If UCase$(Left$(rdata, 9)) = "/SHOW SOS" Then
    Dim M As String
    For n = 1 To Ayuda.Longitud
        M = Ayuda.VerElemento(n)
        Call SendData2(ToIndex, UserIndex, 0, 50, M)
    Next n
    Call SendData2(ToIndex, UserIndex, 0, 51)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "SOSDONE" Then
       rdata = Right$(rdata, Len(rdata) - 7)
    Call Ayuda.Quitar(rdata)
    'delzak sosoffline
   ' rdata = Right$(rdata, Len(rdata) - 7)
    'Call Ayuda.Borra(rdata)
    Exit Sub
End If

'PERDON
If UCase$(Left$(rdata, 7)) = "/PERDON " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    Tindex = NameIndex(rdata)
    If Tindex > 0 Then
        
        If EsNewbie(Tindex) Then
                Call VolverCiudadano(Tindex)
        Else
                Call LogGM(UserList(UserIndex).Name, "Intento perdonar un personaje de nivel avanzado.")
                Call SendData(ToIndex, UserIndex, 0, "||Solo se permite perdonar newbies." & "´" & FontTypeNames.FONTTYPE_info)
        End If
        
    End If
    Exit Sub
End If



'pluto:2.15 index
If UCase$(Left$(rdata, 6)) = "/SLOT " Then
    rdata = Right$(rdata, Len(rdata) - 6)
Dim NSLOT As Integer
   NSLOT = val(rdata)
    If NSLOT <= 0 Or NSLOT > MaxUsers Then
        Call SendData(ToIndex, UserIndex, 0, "||Ese número de slot no existe." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
Call SendData(ToIndex, UserIndex, 0, "||Pj Name: " & UserList(NSLOT).Name & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Pj Cuenta: " & Cuentas(NSLOT).mail & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Conid: " & UserList(NSLOT).ConnID & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||ConidVálida: " & UserList(NSLOT).ConnIDValida & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||ID: " & UserList(NSLOT).ID & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||IP: " & UserList(NSLOT).ip & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Serie: " & UserList(NSLOT).Serie & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Mac: " & UserList(NSLOT).MacPluto & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Valcode: " & UserList(NSLOT).flags.ValCoDe & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||UserLogged: " & UserList(NSLOT).flags.UserLogged & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||CuentaLogged: " & Cuentas(NSLOT).Logged & "´" & FontTypeNames.FONTTYPE_info)
'pluto:6.6
Call SendData(ToIndex, UserIndex, 0, "||Charindex: " & UserList(NSLOT).Char.CharIndex & "´" & FontTypeNames.FONTTYPE_info)

Exit Sub
End If

'pluto:6.6 charindex
If UCase$(Left$(rdata, 11)) = "/CHARINDEX " Then
rdata = Right$(rdata, Len(rdata) - 11)
Dim NuNpc As Integer
NuNpc = val(rdata)
If NuNpc = 0 Or NuNpc > 10000 Then Exit Sub
Call SendData(ToIndex, UserIndex, 0, "||Index en Server: " & UserList(CharList(NuNpc)).Char.CharIndex & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Name: " & UserList(CharList(NuNpc)).Name & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Mapa: " & UserList(CharList(NuNpc)).Pos.Map & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||X: " & UserList(CharList(NuNpc)).Pos.X & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Y: " & UserList(CharList(NuNpc)).Pos.Y & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Body : " & UserList(CharList(NuNpc)).Char.Body & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Email : " & UserList(CharList(NuNpc)).Email & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Hd : " & UserList(CharList(NuNpc)).Serie & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Valcode: " & UserList(CharList(NuNpc)).flags.ValCoDe & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||UserLogged: " & UserList(CharList(NuNpc)).flags.UserLogged & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||CuentaLogged: " & Cuentas(CharList(NuNpc)).Logged & "´" & FontTypeNames.FONTTYPE_info)

Exit Sub
End If


'pluto:2.12 index
If UCase$(Left$(rdata, 7)) = "/INDEX " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Tindex = NameIndex(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
Call SendData(ToIndex, UserIndex, 0, "||Usuario: " & Tindex & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

If UCase$(Left$(rdata, 13)) = "/CUENTAINDEX " Then
    rdata = Right$(rdata, Len(rdata) - 13)
    Tindex = DameIndexCuenta(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
Call SendData(ToIndex, UserIndex, 0, "||Usuario: " & Tindex & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

If UCase$(Left$(rdata, 18)) = "/DESCONECTACUENTA " Then
    rdata = Right$(rdata, Len(rdata) - 18)
    Tindex = DameIndexCuenta(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
Call DesconectaCuenta(Tindex)
Exit Sub
End If
'---------------
'pluto:2.25
If UCase$(Left$(rdata, 6)) = "/DATOS" Then
Tindex = UserList(UserIndex).flags.TargetUser
Call SendData(ToIndex, UserIndex, 0, "||Index: " & Tindex & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Pj Name: " & UserList(Tindex).Name & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Pj Cuenta: " & Cuentas(Tindex).mail & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Posicion: " & UserList(Tindex).Pos.Map & "-" & UserList(Tindex).Pos.X & "-" & UserList(Tindex).Pos.Y & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Conid: " & UserList(Tindex).ConnID & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||ConidVálida: " & UserList(Tindex).ConnIDValida & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||ID: " & UserList(Tindex).ID & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||IP: " & UserList(Tindex).ip & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Serie: " & UserList(Tindex).Serie & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Mac: " & UserList(Tindex).MacPluto & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Valcode: " & UserList(Tindex).flags.ValCoDe & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||UserLogged: " & UserList(Tindex).flags.UserLogged & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||CuentaLogged: " & Cuentas(Tindex).Logged & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

'pluto:2.25
If UCase$(Left$(rdata, 6)) = "/CLON " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Mapa = val(ReadField(1, rdata, 32))
    If Not MapaValido(Mapa) Then Exit Sub
    X = val(ReadField(2, rdata, 32))
    Y = val(ReadField(3, rdata, 32))
If Not InMapBounds(Mapa, X, Y) Then Exit Sub
MapData(Mapa, X, Y).UserIndex = 0
Exit Sub
End If

'pluto:2.25
If UCase$(Left$(rdata, 9)) = "/USERMAP " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Mapa = val(rdata)
    If Not MapaValido(Mapa) Then Exit Sub

        For loopc = 1 To LastUser
          
            If UserList(loopc).Pos.Map = Mapa Then
                tStr = tStr & UserList(loopc).Name & ", "
            End If
        Next loopc
        If tStr = "" Then Exit Sub
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||Número de Usuarios en Mapa " & Mapa & ": " & MapInfo(Mapa).NumUsers & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If


'Echar usuario
If UCase$(Left$(rdata, 7)) = "/ECHAR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Tindex = NameIndex(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    
    If UserList(Tindex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes echar a alguien con jerarquia mayor a la tuya." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
        
    Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(Tindex).Name & "." & "´" & FontTypeNames.FONTTYPE_info)
    Call CloseUser(Tindex)
    Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(Tindex).Name)
    Exit Sub
End If

'Echar usuario
If UCase$(Left$(rdata, 8)) = "/CERRAR " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    Tindex = NameIndex(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    
    If UserList(Tindex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes echar a alguien con jerarquia mayor a la tuya." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
        
    'Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(Tindex).Name & "." & "´" & FontTypeNames.FONTTYPE_info)
    Call SendData(ToIndex, UserIndex, 0, "O6")
    Call LogGM(UserList(UserIndex).Name, "Cerró a " & UserList(Tindex).Name)
    Exit Sub
End If

    If UCase$(rdata) = "/PANELGM" Then
        Call SendData(ToIndex, UserIndex, 0, "JM" & UserList(UserIndex).flags.Privilegios)
        Exit Sub
        End If


'Echar conexion
If UCase$(Left$(rdata, 7)) = "/CONEX " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Tindex = NameIndex(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    
    If UserList(Tindex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes echar a alguien con jerarquia mayor a la tuya." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
        
    Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " cerró conexión a " & UserList(Tindex).Name & "." & "´" & FontTypeNames.FONTTYPE_info)
    Call CloseSocket(Tindex)
    Call LogGM(UserList(UserIndex).Name, "Cerró Conexión a " & UserList(Tindex).Name)
    Exit Sub
End If



'@Nati: Mejoro el CuentaBAN poniendole un motivo, el comando queda así '/CUENTABAN MOTIVOBAN#CUENTA'
If UCase$(Left$(rdata, 11)) = "/CUENTABAN " Then
Dim accindex As Integer
rdata = Right$(rdata, Len(rdata) - 11)
estepj = ReadField(2, rdata, Asc("#"))
Name = ReadField(1, rdata, Asc("#"))
If estepj = "" Or Name = "" Then
Exit Sub
End If
accindex = DameIndexCuenta(ReadField(2, rdata, Asc("#")))
If accindex = 0 Then
If Not CuentaExiste(ReadField(2, rdata, Asc("#"))) Then
Call SendData(ToIndex, UserIndex, 0, "||Esta cuenta no existe" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'@Nati: Comienzo Crea el LOG
Call WriteVar(App.Path & "\logs\" & "BanAccountDetail.dat", ReadField(2, rdata, Asc("#")), "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanAccountDetail.dat", ReadField(2, rdata, Asc("#")), "Reason", Name)
Call WriteVar(App.Path & "\logs\" & "BanAccountDetail.dat", ReadField(2, rdata, Asc("#")), "Fecha", Date)
'@Nati: FIN Crea el LOG
Call WriteVar(AccPath & ReadField(2, rdata, Asc("#")) & ".acc", "DATOS", "Ban", "1")
Call SendData(ToIndex, UserIndex, 0, "||La cuenta se ha baneado directamente a la ficha." & "´" & FontTypeNames.FONTTYPE_info)
Else
Call CloseSocket(accindex)
Call WriteVar(AccPath & ReadField(2, rdata, Asc("#")) & ".acc", "DATOS", "Ban", "1")
'@Nati: Comienzo Crea el LOG
Call WriteVar(App.Path & "\logs\" & "BanAccountDetail.dat", ReadField(2, rdata, Asc("#")), "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanAccountDetail.dat", ReadField(2, rdata, Asc("#")), "Reason", Name)
Call WriteVar(App.Path & "\logs\" & "BanAccountDetail.dat", ReadField(2, rdata, Asc("#")), "Fecha", Date)
'@Nati: FIN Crea el LOG
Call SendData(ToIndex, UserIndex, 0, "||La cuenta se ha baneado" & "´" & FontTypeNames.FONTTYPE_info)
End If

End If

'@Nati: Agrego el NOBODY quitando la cuenta baneada del BanAccountDetail
If UCase$(Left$(rdata, 13)) = "/UNBANCUENTA " Then

rdata = Right$(rdata, Len(rdata) - 13)
accindex = DameIndexCuenta(rdata)
If accindex = 0 Then
If Not CuentaExiste(rdata) Then
Call SendData(ToIndex, UserIndex, 0, "||Esta cuenta no existe" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
Call WriteVar(AccPath & rdata & ".acc", "DATOS", "Ban", "0")
Call SendData(ToIndex, UserIndex, 0, "||La cuenta se ha Desbaneado directamente a la ficha." & "´" & FontTypeNames.FONTTYPE_info)
'@Nati: Comienzo Borra el LOG
Call WriteVar(App.Path & "\logs\" & "BanAccountDetailAcc.dat", rdata, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanAccountDetailAcc.dat", rdata, "Reason", "NOONE")
Call WriteVar(App.Path & "\logs\" & "BanAccountDetailAcc.dat", rdata, "Fecha", "NOONE")
'@Nati: Fin Borra el LOG
Else
Call CloseSocket(accindex)
Call WriteVar(AccPath & rdata & ".acc", "DATOS", "Ban", "0")
'@Nati: Comienzo Borra el LOG
Call WriteVar(App.Path & "\logs\" & "BanAccountDetailAcc.dat", rdata, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanAccountDetailAcc.dat", rdata, "Reason", "NOONE")
Call WriteVar(App.Path & "\logs\" & "BanAccountDetailAcc.dat", rdata, "Fecha", "NOONE")
'@Nati: Fin Borra el LOG
Call SendData(ToIndex, UserIndex, 0, "||La cuenta se ha Desbaneado" & "´" & FontTypeNames.FONTTYPE_info)
End If

End If

'@Nati: Agrego el comando '/MotivoACC '
If UCase$(Left$(rdata, 11)) = "/MOTIVOACC " Then
rdata = Right$(rdata, Len(rdata) - 11)
Dim ban As String
Dim rean As String
Dim rean2 As String
ban = GetVar(App.Path & "\logs\" & "BanAccountDetail.dat", rdata, "BannedBy")
rean = GetVar(App.Path & "\logs\" & "BanAccountDetail.dat", rdata, "Reason")
rean2 = GetVar(App.Path & "\logs\" & "BanAccountDetail.dat", rdata, "Fecha")
If ban <> "" Or rean <> "" Then
Call SendData(ToIndex, UserIndex, 0, "|| Gm : " & ban & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Motivo: " & rean & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Fecha: " & rean2 & "´" & FontTypeNames.FONTTYPE_info)
Else
Call SendData(ToIndex, UserIndex, 0, "|| No hay datos." & "´" & FontTypeNames.FONTTYPE_info)
End If
Exit Sub
End If


'nati:cambio el PJBAN por el nuevo que deja log HD
If UCase$(Left$(rdata, 7)) = "/PJBAN " Then
rdata = Right$(rdata, Len(rdata) - 7)
Tindex = NameIndex(ReadField(2, rdata, Asc("@")))
Name = ReadField(1, rdata, Asc("@"))
'pluto:2.4
Dim este As String
este = ReadField(2, rdata, Asc("@"))
archiv = CharPath & Left$(este, 1) & "\" & este & ".chr"
If Tindex <= 0 Then
If PersonajeExiste(este) Then
Dim CANALBAN As Integer
CANALBAN = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BAN\" & GetVar(archiv, "INIT", "LastSerie") & ".dat" For Append As #CANALBAN
Print #CANALBAN, "PJ:" & este & " Fecha:" & Date & " GM:" & UserList(UserIndex).Name & " Razón:" & Name
Close #CANALBAN
Call SendData(ToIndex, UserIndex, 0, "||Ban directo a su ficha." & "´" & FontTypeNames.FONTTYPE_info)
Call WriteVar(CharPath & Left$(este, 1) & "\" & este & ".chr", "FLAGS", "Ban", 1)
'pluto:2.11
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", este, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", este, "Reason", Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", este, "Fecha", Date)

Else
Call SendData(ToIndex, UserIndex, 0, "||Ese Pj no existe." & "´" & FontTypeNames.FONTTYPE_info)
End If
Exit Sub
End If
'pluto:hoy
UltimoBan = UserList(Tindex).Name

If UserList(Tindex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
Call SendData(ToIndex, UserIndex, 0, "||No podes banear a al alguien de mayor jerarquia." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

Call LogBan(Tindex, UserIndex, Name)

Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(Tindex).Name & "." & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Banned a " & UserList(Tindex).Name & "." & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendData(ToAll, 0, 0, "||Los dioses han desterrado a: " & UserList(Tindex).Name & "´" & FontTypeNames.FONTTYPE_FIGHT)
'Ponemos el flag de ban a 1
UserList(Tindex).flags.ban = 1
'Dim CANALBAN As Integer
CANALBAN = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BAN\" & GetVar(archiv, "INIT", "LastSerie") & ".dat" For Append As #CANALBAN
Print #CANALBAN, "PJ:" & este & " Fecha:" & Date & " GM:" & UserList(UserIndex).Name & " Razón:" & Name
Close #CANALBAN
If UserList(Tindex).flags.Privilegios > 0 Then
' UserList(UserIndex).Flags.Ban = 1
Call CloseUser(UserIndex)
Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " banned by the server por bannear un Administrador." & "´" & FontTypeNames.FONTTYPE_FIGHT)
Exit Sub
End If

Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(Tindex).Name)
Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(Tindex).Name)
Call CloseUser(Tindex)
Exit Sub
End If
'nati:cambio el PJBAN por el nuevo que deja log HD
'-------------------------------------------------

'nati:Agrego el comando /PENAS para visualizar todos los bans del usuario (HD)

If UCase$(Left$(rdata, 7)) = "/PENAS " Then
rdata = Right$(rdata, Len(rdata) - 7)
If FileExist("\logs\BAN\" & rdata & ".chr", vbNormal) Then
Open App.Path & "\logs\BAN\" & rdata & ".dat" For Input As #1
Do While Not EOF(1)
Input #1, DATOSBAN
Call SendData(ToIndex, UserIndex, 0, "|| " & DATOSBAN & "´" & FontTypeNames.FONTTYPE_info)
Loop
Close #1
If DATOSBAN = "" Then
Call SendData(ToIndex, UserIndex, 0, "|| No hay datos." & "´" & FontTypeNames.FONTTYPE_info)
Else
Call SendData(ToIndex, UserIndex, 0, "|| No hay ninguna pena." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
End If
End If
'nati:Agrego el comando /PENAS para visualizar todos los bans del usuario (HD)
'-----------------------------------------------------------------------------


'@Nati: Agrego el comando '/MotivoPC
If UCase$(Left$(rdata, 10)) = "/MOTIVOPC " Then
rdata = Right$(rdata, Len(rdata) - 10)
Dim banpc As String
Dim reapc As String
Dim reapc2 As String
banpc = GetVar(App.Path & "\logs\" & "BanPCDetail.dat", rdata, "BannedBy")
reapc = GetVar(App.Path & "\logs\" & "BanPCDetail.dat", rdata, "Reason")
reapc2 = GetVar(App.Path & "\logs\" & "BanPCDetail.dat", rdata, "Fecha")
If banpc <> "" Or reapc <> "" Then
Call SendData(ToIndex, UserIndex, 0, "|| Gm : " & banpc & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Motivo: " & reapc & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Fecha: " & reapc2 & "´" & FontTypeNames.FONTTYPE_info)
Else
Call SendData(ToIndex, UserIndex, 0, "|| No hay datos." & "´" & FontTypeNames.FONTTYPE_info)
End If
Exit Sub
End If
'-------------------------------

'@Nati: NUEVO BLOQUEO: /BLOQUEO MOTIVO#PJ
If UCase$(Left$(rdata, 9)) = "/BLOQUEO " Then
Dim asuntobloqueo As String
Dim Desbanea As Byte
Dim namebloqueo As String
rdata = Right$(rdata, Len(rdata) - 9)
'Dim name2 As String
namebloqueo = ReadField(2, rdata, Asc("#"))
asuntobloqueo = ReadField(1, rdata, Asc("#"))
If namebloqueo = "" Then
namebloqueo = ReadField(1, rdata, Asc("#"))
End If
Tindex = NameIndex(namebloqueo)
If Tindex <= 0 Then 'no ta online
'Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info)
If PersonajeExiste(namebloqueo) Then
Dim miraficha As String
Dim mira2 As String
miraficha = App.Path & "\Charfile\" & Left$(namebloqueo, 1) & "\" & namebloqueo & ".chr"
mira2 = GetVar(miraficha, "INIT", "LastSerie")

If Not FileExist(App.Path & "\Bloqueos\" & mira2 & ".lol", vbArchive) Then
Call WriteVar(App.Path & "\Bloqueos\" & mira2 & ".lol", "INIT", "NOMBRE", rdata)
Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " Ha sido Bloqueado." & "´" & FontTypeNames.FONTTYPE_TALK)
Call LogGM(UserList(UserIndex).Name, "Bloqueo a " & namebloqueo)
'@Nati: Comienzo Crea el LOG
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "Reason", asuntobloqueo)
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "Fecha", Date)
'@Nati: FIN Crea el LOG
Else
Kill (App.Path & "\Bloqueos\" & mira2 & ".lol")
Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " Ha sido DesBloqueado." & "´" & FontTypeNames.FONTTYPE_TALK)
Call LogGM(UserList(UserIndex).Name, "DesBloqueo a " & rdata)
'@Nati: Comienzo Crea el LOG
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "Reason", "NOONE")
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "Fecha", "NOONE")
'@Nati: FIN Crea el LOG
End If

Else 'existe
Call SendData(ToIndex, UserIndex, 0, "|| No existe esa ficha." & "´" & FontTypeNames.FONTTYPE_info)
End If

Exit Sub
End If ' no ta online

If Not FileExist(App.Path & "\Bloqueos\" & UserList(Tindex).Serie & ".lol", vbArchive) Then
Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " Ha sido Bloqueado." & "´" & FontTypeNames.FONTTYPE_TALK)
Call LogGM(UserList(UserIndex).Name, "Bloqueo a " & rdata & " -> " & UserList(Tindex).Name)
Call WriteVar(App.Path & "\Bloqueos\" & UserList(Tindex).Serie & ".lol", "INIT", "NOMBRE", UserList(Tindex).Name)
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "Reason", asuntobloqueo)
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "Fecha", Date)
Else
Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " Ha sido DesBloqueado." & "´" & FontTypeNames.FONTTYPE_TALK)
Call LogGM(UserList(UserIndex).Name, "DesBloqueo a " & rdata & " -> " & UserList(Tindex).Name)
Call Kill(App.Path & "\Bloqueos\" & UserList(Tindex).Serie & ".lol")
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "Reason", "NOONE")
Call WriteVar(App.Path & "\logs\" & "BanPCDetail.dat", namebloqueo, "Fecha", "NOONE")
Exit Sub
End If

End If 'bloqueo

'pluto:2.14
If UCase$(Left$(rdata, 12)) = "/MACBLOQUEO " Then
rdata = Right$(rdata, Len(rdata) - 12)
'Dim name2 As String
name2 = rdata
Tindex = NameIndex(name2)

If Tindex <= 0 Then 'no ta online
Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info)
If PersonajeExiste(rdata) Then
'Dim miraficha As String
'Dim mira2 As String
miraficha = App.Path & "\Charfile\" & Left$(rdata, 1) & "\" & rdata & ".chr"
mira2 = GetVar(miraficha, "INIT", "LastMac")

If Not FileExist(App.Path & "\MacPluto\" & mira2 & ".lol", vbArchive) Then
Call WriteVar(App.Path & "\MacPluto\" & mira2 & ".lol", "INIT", "NOMBRE", rdata)
Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " Ha sido Mac-Bloqueado." & "´" & FontTypeNames.FONTTYPE_TALK)
Call LogGM(UserList(UserIndex).Name, "Mac-Bloqueo a " & rdata)
Else
Kill (App.Path & "\MacPluto\" & mira2 & ".lol")
Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " Ha sido Mac-DesBloqueado." & "´" & FontTypeNames.FONTTYPE_TALK)
Call LogGM(UserList(UserIndex).Name, "Mac-DesBloqueo a " & rdata)
End If

Else 'existe
Call SendData(ToIndex, UserIndex, 0, "|| No existe esa ficha." & "´" & FontTypeNames.FONTTYPE_info)
End If

Exit Sub
End If ' no ta online

If Not FileExist(App.Path & "\MacPluto\" & UserList(Tindex).MacPluto & ".lol", vbArchive) Then
Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " Ha sido Mac-Bloqueado." & "´" & FontTypeNames.FONTTYPE_TALK)
Call LogGM(UserList(UserIndex).Name, "Mac-Bloqueo a " & rdata & " -> " & UserList(Tindex).Name)
Call WriteVar(App.Path & "\MacPluto\" & UserList(Tindex).MacPluto & ".lol", "INIT", "NOMBRE", UserList(Tindex).Name)
Else
Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " Ha sido Mac-DesBloqueado." & "´" & FontTypeNames.FONTTYPE_TALK)
Call LogGM(UserList(UserIndex).Name, "Mac-DesBloqueo a " & rdata & " -> " & UserList(Tindex).Name)
Call Kill(App.Path & "\MacPluto\" & UserList(Tindex).MacPluto & ".lol")
Exit Sub
End If

End If 'bloqueo

'-------------------------------



If UCase$(Left$(rdata, 7)) = "/UNBAN " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    If PersonajeExiste(rdata) Then
        Call UnBan(rdata)
        Call LogGM(UserList(UserIndex).Name, "/UNBAN a " & rdata)
        Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " unbanned." & "´" & FontTypeNames.FONTTYPE_info)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe" & "´" & FontTypeNames.FONTTYPE_info)
    End If
    Exit Sub
End If

'Pasar Pregunta
If UCase$(Left$(rdata, 8)) = "/TRIVIAL" Then
    Call Loadtrivial
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/ACC " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   If val(rdata) = 692 Then Exit Sub
   Call SpawnNpc(val(rdata), UserList(UserIndex).Pos, True, False)
   Call LogGM(UserList(UserIndex).Name, "/ACC Bicho:" & val(rdata) & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y)
Exit Sub
End If

'Teleportar
If UCase$(Left$(rdata, 7)) = "/TELEP " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Mapa = val(ReadField(2, rdata, 32))
    If Not MapaValido(Mapa) Then Exit Sub
    Name = ReadField(1, rdata, 32)
    If Name = "" Then Exit Sub
    If UCase$(Name) <> "YO" Then
        Tindex = NameIndex(Name)
   Else
        Tindex = UserIndex
    End If
     X = val(ReadField(3, rdata, 32))
    Y = val(ReadField(4, rdata, 32))
    If Not InMapBounds(Mapa, X, Y) Then Exit Sub
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    Call WarpUserChar(Tindex, Mapa, X, Y, True)
    Call SendData(ToIndex, Tindex, 0, "||" & UserList(UserIndex).Name & " transportado." & "´" & FontTypeNames.FONTTYPE_info)
    Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(Tindex).Name & " hacia " & "Mapa" & Mapa & " X:" & X & " Y:" & Y)
    Exit Sub
End If

    If UCase$(Left$(rdata, 4)) = "/GO " Then
        rdata = Right$(rdata, Len(rdata) - 4)
        Mapa = val(ReadField(1, rdata, 32))
        If Not MapaValido(Mapa) Then Exit Sub
        If UserList(UserIndex).flags.Privilegios = 1 And MapInfo(Mapa).Pk Then Exit Sub
        Call WarpUserChar(UserIndex, Mapa, 50, 50, True)
        Call SendData(ToIndex, UserIndex, 0, "2B" & UserList(UserIndex).Name)
        'Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(UserIndex).Name & " hacia " & "Mapa" & Mapa & " X:" & X & " Y:" & Y, (UserList(UserIndex).flags.Privilegios = 1))
        Exit Sub
    End If

'Summon
If UCase$(Left$(rdata, 5)) = "/SUM " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    Tindex = NameIndex(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El jugador no esta online." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Map = 165 And UserList(UserIndex).flags.Montura > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El jugador no puede ser devuelto porque el mapa no permite mascotas." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If Ayuda.Existe(UserList(Tindex).Name) Or UserList(UserIndex).flags.Privilegios > 1 Or UserList(Tindex).Pos.Map = 303 Then
    Call SendData(ToIndex, Tindex, 0, "||" & UserList(UserIndex).Name & " te ha transportado." & "´" & FontTypeNames.FONTTYPE_info)
    'pluto:2.9.0
 Dim aa As Integer
'pluto:6.8----------------------------------
UserList(Tindex).PoSum.Map = UserList(Tindex).Pos.Map
UserList(Tindex).PoSum.X = UserList(Tindex).Pos.X
UserList(Tindex).PoSum.Y = UserList(Tindex).Pos.Y
'----------------------
 'If UserList(UserIndex).Pos.Y > 90 Then aa = -1 Else aa = 1
 Call WarpUserChar(Tindex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + aa, True)
    
    Call LogGM(UserList(UserIndex).Name, "/SUM " & UserList(Tindex).Name & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y)
    Exit Sub
'    Else
'    Call LogGM(UserList(UserIndex).Name, "INTENTO /SUM A USUARIO SIN AYUDA: " & UserList(Tindex).Name & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y)
'    End If '
End If
End If


'SEGUIR
If UCase$(rdata) = "/SEGUIR" Then
    If UserList(UserIndex).flags.TargetNpc > 0 Then
        Call DoFollow(UserList(UserIndex).flags.TargetNpc, UserList(UserIndex).Name)
    End If
    Exit Sub
End If
 
 
 'clave
 If UCase$(Left$(rdata, 7)) = "/CLAVE " Then
    'Dim nickx As String
    rdata = Right$(rdata, Len(rdata) - 7)
    nickx = ReadField(1, rdata, 44)
    Tindex = NameIndex(nickx)
    If (Len(ReadField(2, rdata, 44)) < 7) Then
        Call SendData(ToIndex, UserIndex, 0, "||Contraseña demasiado corta, escriba una mas grande" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
 
 If CuentaExiste(nickx) Then
        Call SendData(ToIndex, UserIndex, 0, "||Su email es: " & nickx & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||La nueva clave es: " & ReadField(2, rdata, 44) & "´" & FontTypeNames.FONTTYPE_info)
    
Call WriteVar(AccPath & nickx & ".acc", "DATOS", "Password", MD5String(ReadField(2, rdata, 44)))

    'pluto:
    Call LogGM(UserList(UserIndex).Name, "CAMBIO CLAVE: " & nickx)
      
    Else
        Call SendData(ToIndex, UserIndex, 0, "||La cuenta no existe" & "´" & FontTypeNames.FONTTYPE_info)
    End If
    Exit Sub
    End If


    
    'pluto:2.13
     If UCase$(Left$(rdata, 13)) = "/LISTARCLAVES" Then
     Call ListarClaves(UserIndex)
     Exit Sub
     End If
     If UCase$(Left$(rdata, 13)) = "/BORRARCLAVES" Then
     Call BorrarClaves(UserIndex)
     Exit Sub
     End If
     If UCase$(Left$(rdata, 9)) = "/BUSCAIP " Then
     rdata = Right$(rdata, Len(rdata) - 9)
     If rdata = "" Then Exit Sub
    
    Tindex = NameIndex(rdata)
      If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
      End If
     For n = 1 To MaxUsers
     If UserList(n).ip = UserList(Tindex).ip Then Call SendData(ToIndex, UserIndex, 0, "|| " & UserList(n).Name & "´" & FontTypeNames.FONTTYPE_info)
     Next
     Exit Sub
     End If
     'pluto:2.14
     If UCase$(Left$(rdata, 12)) = "/BUSCASERIE " Then
     rdata = Right$(rdata, Len(rdata) - 12)
     If rdata = "" Then Exit Sub
     For n = 1 To MaxUsers
     If UserList(n).Serie = rdata Then Call SendData(ToIndex, UserIndex, 0, "|| " & UserList(n).Name & "´" & FontTypeNames.FONTTYPE_info)
     Next
     Exit Sub
     End If
     'pluto:6.7
     If UCase$(Left$(rdata, 10)) = "/BUSCAMAC " Then
     rdata = Right$(rdata, Len(rdata) - 10)
     If rdata = "" Then Exit Sub
     For n = 1 To MaxUsers
     If UserList(n).MacPluto = rdata Then Call SendData(ToIndex, UserIndex, 0, "|| " & UserList(n).Name & "´" & FontTypeNames.FONTTYPE_info)
     Next
     Exit Sub
     End If
     '---------------------------------------------
    
     


    

'pluto:2.11
If UCase$(Left$(rdata, 10)) = "/PJCUENTA " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    If CuentaExiste(rdata) Then
      
archiv = App.Path & "\Accounts\" & rdata & ".acc"
For n = 1 To GetVar(archiv, "DATOS", "NumPjs")
Call SendData(ToIndex, UserIndex, 0, "||Pj: " & GetVar(archiv, "PERSONAJES", "PJ" & n) & "´" & FontTypeNames.FONTTYPE_info)
Next n

    Else
Call SendData(ToIndex, UserIndex, 0, "||La cuenta no existe" & "´" & FontTypeNames.FONTTYPE_info)
    End If
Exit Sub
End If

'mover
If UCase$(Left$(rdata, 7)) = "/MOVER " Then
   
    rdata = Right$(rdata, Len(rdata) - 7)
    If PersonajeExiste(rdata) Then
        Call SendData(ToIndex, UserIndex, 0, "|| El pj ha sido transportado a nix." & "´" & FontTypeNames.FONTTYPE_info)
        Call WriteVar(CharPath & Left$(rdata, 1) & "\" & rdata & ".chr", "INIT", "Position", "34-34-34")
     'pluto:2.11
    Call LogGM(UserList(UserIndex).Name, "/MOVER " & rdata)

    End If
End If

'pluto:2.11
If UCase$(Left$(rdata, 11)) = "/MOTIVOBAN " Then
    rdata = Right$(rdata, Len(rdata) - 11)
ba = GetVar(App.Path & "\logs\" & "BanDetail.dat", rdata, "BannedBy")
rea = GetVar(App.Path & "\logs\" & "BanDetail.dat", rdata, "Reason")
rea2 = GetVar(App.Path & "\logs\" & "BanDetail.dat", rdata, "Fecha")
If ba <> "" Or rea <> "" Then
Call SendData(ToIndex, UserIndex, 0, "|| Gm : " & ba & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Motivo: " & rea & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Fecha: " & rea2 & "´" & FontTypeNames.FONTTYPE_info)
Else
Call SendData(ToIndex, UserIndex, 0, "|| No hay datos." & "´" & FontTypeNames.FONTTYPE_info)
End If
Exit Sub
End If
'---------------------------------------------

If UCase$(rdata) = "/TELEPLOC" Then
    Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
    Call LogGM(UserList(UserIndex).Name, "/TELEPLOC a x:" & UserList(UserIndex).flags.TargetX & " Y:" & UserList(UserIndex).flags.TargetY & " Map:" & UserList(UserIndex).Pos.Map)
    Exit Sub
End If

'Delzak) trafico optimizado (gran cagada!!)

'If UCase$(rdata) = "/TELEPLOC" Then
 '   If UserList(UserIndex).Pos.Map = UserList(UserIndex).flags.TargetMap Then
  '      UserList(UserIndex).Pos.X = UserList(UserIndex).flags.TargetX
   '     UserList(UserIndex).Pos.Y = UserList(UserIndex).flags.TargetY
    '    Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
     '   Call SendData2(ToIndex, UserIndex, 0, 15, UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
    'Else
   ' Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
   ' Call LogGM(UserList(UserIndex).Name, "/TELEPLOC a x:" & UserList(UserIndex).flags.TargetX & " Y:" & UserList(UserIndex).flags.TargetY & " Map:" & UserList(UserIndex).Pos.Map)
   ' End If
   ' Exit Sub
'End If


'IR A
If UCase$(Left$(rdata, 5)) = "/IRA " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    Tindex = NameIndex(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    Call WarpUserChar(UserIndex, UserList(Tindex).Pos.Map, UserList(Tindex).Pos.X, UserList(Tindex).Pos.Y + 1, True)
    'pluto:2-3-04
    'Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " se ha trasportado hacia donde estas." & FONTTYPENAMES.FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "/IRA " & UserList(Tindex).Name & " Mapa:" & UserList(Tindex).Pos.Map & " X:" & UserList(Tindex).Pos.X & " Y:" & UserList(Tindex).Pos.Y)
    Exit Sub
End If

'If UCase$(rdata) = "/TELEPLOC" Then
 '   Call WarpUserChar2(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
  '  Call LogGM(UserList(UserIndex).Name, "/TELEPLOC a x:" & UserList(UserIndex).flags.TargetX & " Y:" & UserList(UserIndex).flags.TargetY & " Map:" & UserList(UserIndex).Pos.Map)
   ' Exit Sub
'End If

'pluto:2.9.0
If UCase$(Left$(rdata, 6)) = "/BALON" Then
'If UserList(UserIndex).Pos.Map <> 192 Then Exit Sub
Call SpawnNpc(151, UserList(UserIndex).Pos, True, False)
Exit Sub
End If

If UCase$(rdata) = "/SAQUE" Then
If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
'If UserList(UserIndex).Pos.Map <> 192 Then Exit Sub
Call QuitarNPC(UserList(UserIndex).flags.TargetNpc)
Call SpawnNpc(151, UserList(UserIndex).Pos, True, False)
Vezz = 0
Exit Sub
End If
If UCase$(rdata) = "/INICIO" Then
'If UserList(UserIndex).Pos.Map <> 192 Then Exit Sub
GolesLocal = 0: GolesVisitante = 0
Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 92, GolesLocal & "," & GolesVisitante & "," & 2)
Exit Sub
End If
'pluto:2.15
If UCase$(Left$(rdata, 6)) = "/NOTA " Then
rdata = Right$(rdata, Len(rdata) - 6)
Call LogGM(UserList(UserIndex).Name, "Nota: " & rdata)
Call SendData(ToIndex, UserIndex, 0, "||Nota añadida en tu Log: " & rdata & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

'pluto:6.8
'Devolver Summon------------------------
If UCase$(Left$(rdata, 5)) = "/DEV " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    Tindex = NameIndex(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El jugador no esta online." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    
    If UserList(Tindex).PoSum.Map = 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||El jugador no tiene guardada su posición." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
    
    If UserList(UserIndex).Pos.Map = 165 And UserList(UserIndex).flags.Montura > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||El jugador no puede ser devuelto porque el mapa no permite mascotas." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
    
    Call SendData(ToIndex, Tindex, 0, "||" & UserList(UserIndex).Name & " te ha transportado." & "´" & FontTypeNames.FONTTYPE_info)


  Call WarpUserChar(Tindex, UserList(Tindex).PoSum.Map, UserList(Tindex).PoSum.X, UserList(Tindex).PoSum.Y, True)
    
    Call LogGM(UserList(UserIndex).Name, "/DEV " & UserList(Tindex).Name & " Map:" & UserList(Tindex).PoSum.Map & " X:" & UserList(Tindex).PoSum.X & " Y:" & UserList(Tindex).PoSum.Y)

UserList(Tindex).PoSum.Map = 0
UserList(Tindex).PoSum.X = 0
UserList(Tindex).PoSum.Y = 0
'----------------------
    Exit Sub
End If

'----------------------------------------------------------








'Crear criatura
If UCase$(Left$(rdata, 3)) = "/CC" Then
'If UserList(UserIndex).Pos.Map = mapainvasion Then
   Call EnviarSpawnList(UserIndex)
   Call LogGM(UserList(UserIndex).Name, "/CC  Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y)
'Exit Sub
'Else
'Call SendData(ToIndex, UserIndex, 0, "||Para usar el comando /CC debes habilitar el mapa para invasiones." & "´" & FontTypeNames.FONTTYPE_info)
'End If
End If

'Spawn!!!!!
If UCase$(Left$(rdata, 3)) = "SPA" Then
    rdata = Right$(rdata, Len(rdata) - 3)
    
    If val(rdata) > 0 And val(rdata) < UBound(SpawnList) + 1 Then _
          Call SpawnNpc(SpawnList(val(rdata)).NpcIndex, UserList(UserIndex).Pos, True, False)
          
          Call LogGM(UserList(UserIndex).Name, "Sumoneo " & SpawnList(val(rdata)).NpcName)
          
    Exit Sub
End If

'Haceme invisible vieja!
If UCase$(rdata) = "/INVISIBLE" Then
    Call DoAdminInvisible(UserIndex)
    Call LogGM(UserList(UserIndex).Name, "/INVISIBLE")
    Exit Sub
End If

'Resetea el inventario
If UCase$(rdata) = "/RESETINV" Then
    rdata = Right$(rdata, Len(rdata) - 9)
    If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
    Call ResetNpcInv(UserList(UserIndex).flags.TargetNpc)
    Call LogGM(UserList(UserIndex).Name, "/RESETINV " & Npclist(UserList(UserIndex).flags.TargetNpc).Name)
    Exit Sub
End If

'/Clean
'If UCase$(rdata) = "/LIMPIAR" Then
    'Call LimpiarMundo
   ' Exit Sub
'End If
'[Tite]Party


'[\Tite]
'Mensaje del servidor
If UCase$(Left$(rdata, 6)) = "/RMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(UserIndex).Name, "Mensaje Broadcast:" & rdata)
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & ": " & rdata & "´" & FontTypeNames.FONTTYPE_TALK)
    End If
    Exit Sub
End If
'Mensaje publicidad
If UCase$(Left$(rdata, 6)) = "/PUBLI" Then
Call SendData(ToAll, 0, 0, "K6")
Exit Sub
End If
'Mensaje Sms
If UCase$(Left$(rdata, 5)) = "/SMS " Then
rdata = Right$(rdata, Len(rdata) - 5)
Call SendData(ToAll, 0, 0, "Z7" & rdata)
Exit Sub
End If
'Pluto:2.15 Mensaje del servidor al mapa
If UCase$(Left$(rdata, 7)) = "/RMSG2 " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Call LogGM(UserList(UserIndex).Name, "Mensaje MapaBroadcast:" & rdata)
    If rdata <> "" Then
        Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "||" & UserList(UserIndex).Name & ": " & rdata & "´" & FontTypeNames.FONTTYPE_TALK)
    End If
    Exit Sub
End If

'pluto:2.9.0
'Mensaje del servidor entrada
If UCase$(Left$(rdata, 9)) = "/MENSAJE " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Call LogGM(UserList(UserIndex).Name, "Mensaje Entrada:" & rdata)
    If rdata <> "" Then MsgEntra = rdata: Call SendData(ToIndex, UserIndex, 0, "||Mensaje de Entrada Activado: " & rdata & "´" & FontTypeNames.FONTTYPE_info)
    If UCase$(rdata) = "NINGUNO" Then MsgEntra = "": Call SendData(ToIndex, UserIndex, 0, "||Mensaje de Entrada Desactivado." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
End If


'Ip del nick
If UCase$(Left$(rdata, 8)) = "/IPNICK " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    Tindex = NameIndex(UCase$(rdata))
    If Tindex > 0 Then
       Call SendData(ToIndex, UserIndex, 0, "||El ip de " & rdata & " es " & UserList(Tindex).ip & "´" & FontTypeNames.FONTTYPE_info)
    Else
    Call SendData(ToIndex, UserIndex, 0, "||No existe" & "´" & FontTypeNames.FONTTYPE_info)
    End If
    Exit Sub
End If

'Ip del nick
If UCase$(Left$(rdata, 8)) = "/NICKIP " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    Tindex = IP_Index(rdata)
    If Tindex > 0 Then
       Call SendData(ToIndex, UserIndex, 0, "||El nick del ip " & rdata & " es " & UserList(Tindex).Name & "´" & FontTypeNames.FONTTYPE_info)
    Else
    Call SendData(ToIndex, UserIndex, 0, "||No existe" & "´" & FontTypeNames.FONTTYPE_info)
End If
    Exit Sub
End If

'Quitar NPC 2 ESPECIAL SOLO GUARDIAS
If UCase$(rdata) = "/MATA2" Then
    If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
    If Npclist(UserList(UserIndex).flags.TargetNpc).numero = 722 Then
    Call QuitarNPC(UserList(UserIndex).flags.TargetNpc)
    Call LogGM(UserList(UserIndex).Name, "/MATA2 " & Npclist(UserList(UserIndex).flags.TargetNpc).Name)
    Else
    Call LogGM(UserList(UserIndex).Name, "Intento a un NO-GUARDIAN /MATA2 " & Npclist(UserList(UserIndex).flags.TargetNpc).Name)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/BORRAR SOS" Then
    Call Ayuda.Reset
    Exit Sub
End If

'Bloquear
If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
    Call LogGM(UserList(UserIndex).Name, "/BLOQ")
    rdata = Right$(rdata, Len(rdata) - 5)
If UserList(UserIndex).flags.Privilegios < 2 Then 'PRIVILEGIOS >
If UserList(UserIndex).Pos.Map = 303 Then 'MAPA 303 >
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0 Then
        MapData(UserList(UserIndeax).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 1
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 1)
    Else
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 0)
    End If
    Exit Sub
End If 'MAPA 303 <
Else 'PRIVILEGIOS <>
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0 Then
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 1
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 1)
    Else
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 0)
    End If
    Exit Sub
End If 'PRIVILEGIOS <
Exit Sub
End If

'CUENTA REGRESIVA
If UCase$(Left$(rdata, 3)) = "/CU" Then
CuentaRegresiva = 220
indexCuentaRegresiva = UserIndex
End If

'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
If UserList(UserIndex).flags.Privilegios < 3 Then
    Exit Sub
End If

'Nro de enemigos
If UCase$(Left$(rdata, 6)) = "/NENE " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    If MapaValido(val(rdata)) Then
        Call SendData2(ToIndex, UserIndex, 0, 49, NPCHostiles(rdata))
        Call LogGM(UserList(UserIndex).Name, "/NENE")
    End If
    Exit Sub
End If

'¿Donde esta?
If UCase$(Left$(rdata, 7)) = "/DONDE " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Tindex = NameIndex(rdata)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    'pluto:2.14
    If UserList(Tindex).flags.Privilegios > 2 And UserList(UserIndex).flags.Privilegios < 3 Then
    Call SendData(ToIndex, UserIndex, 0, "|| No seas cotilla." & "´" & FontTypeNames.FONTTYPE_info)
    Call SendData(ToGM, Tindex, 0, "|| /DONDE del SemiDios " & UserList(UserIndex).Name & " sobre el Dios " & UserList(Tindex).Name & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "||Ubicacion  " & UserList(Tindex).Name & ": " & UserList(Tindex).Pos.Map & ", " & UserList(Tindex).Pos.X & ", " & UserList(Tindex).Pos.Y & "." & "´" & FontTypeNames.FONTTYPE_info)
    'pluto:2-3-04
    Call LogGM(UserList(UserIndex).Name, "/Donde " & UserList(Tindex).Name & ": " & UserList(Tindex).Pos.Map & ", " & UserList(Tindex).Pos.X & ", " & UserList(Tindex).Pos.Y)

    Exit Sub
End If

'pluto:2.23-------------
'If Left$(UCase$(rdata), 13) = "/DUMPSECURITY" Then
    'Call securityip.DumpTables
    'Exit Sub
'End If
'------------------------

'pluto:5.2------------------------------------
If UCase$(Left$(rdata, 5)) = "/HOY " Then
        rdata = Right$(rdata, Len(rdata) - 5)
         ReNumUsers = val(rdata)
        Call SendData(ToGM, UserIndex, 0, "||Record Hoy cambiado a: " & rdata & "´" & FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
If UCase$(Left$(rdata, 6)) = "/AYER " Then
        rdata = Right$(rdata, Len(rdata) - 6)
         AyerReNumUsers = val(rdata)
        Call SendData(ToGM, UserIndex, 0, "||Record Ayer Cambiado a: " & rdata & "´" & FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
'-----------------------------------------------
'PLUTO:6.9
If UCase$(Left$(rdata, 13)) = "/AVISOLANZAR " Then
        rdata = Right$(rdata, Len(rdata) - 13)
         TOPELANZAR = val(rdata)
        Call SendData(ToGM, UserIndex, 0, "||Aviso Lanzar Cambiado a: " & rdata & "´" & FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
If UCase$(Left$(rdata, 13)) = "/AVISOFLECHA " Then
        rdata = Right$(rdata, Len(rdata) - 13)
         TOPEFLECHA = val(rdata)
        Call SendData(ToGM, UserIndex, 0, "||Aviso Flecha Cambiado a: " & rdata & "´" & FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
'-----------------------------------------------

'pluto:2.15
If UCase$(Left$(rdata, 8)) = "/JOPUTA " Then
        rdata = Right$(rdata, Len(rdata) - 8)
         Joputa = rdata
        Call SendData(ToGM, UserIndex, 0, "||Ataque Bloqueado: " & Joputa & "´" & FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If

'pluto:2.15
If UCase$(rdata) = "/CAPULLO" Then
If joputa2 = 0 Then
Call SendData(ToGM, UserIndex, 0, "||Visualizando Ips: " & "´" & FontTypeNames.FONTTYPE_TALK)
joputa2 = 1
Else
Call SendData(ToGM, UserIndex, 0, "||Stop Visualizar Ips: " & "´" & FontTypeNames.FONTTYPE_TALK)
joputa2 = 0
End If
Exit Sub
End If
'pluto:2.14
If UCase$(Left$(rdata, 12)) = "/BODYTORNEO " Then
        rdata = Right$(rdata, Len(rdata) - 12)
         BodyTorneo = val(rdata)
        Call SendData(ToGM, UserIndex, 0, "||BodyTorneo cambiado: " & BodyTorneo & "´" & FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If

    
If UCase$(Left$(rdata, 10)) = "/CLOSEVEN " Then
Dim Numeraco As Long
Call LogGM(UserList(UserIndex).Name, rdata)
rdata = Right$(rdata, Len(rdata) - 10)
Tindex = NameIndex(ReadField(1, rdata, 44))
Numeraco = val(ReadField(2, rdata, 44))

    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If

If Numeraco <= 0 Then Exit Sub
Call SendData(ToIndex, Tindex, 0, "Z2" & Numeraco)
Call SendData(ToIndex, UserIndex, 0, "||Número Cerrado." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

'obtener DATOS
    'pluto.2.4.7
     If UCase$(Left$(rdata, 10)) = "/FOTOINVI " Then
        rdata = Right$(rdata, Len(rdata) - 10)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData2(ToIndex, Tindex, 0, 83, CStr(UserIndex))
        Exit Sub
    End If


'pluto:6.9
If UCase$(Left$(rdata, 10)) = "/BORRAWPE " Then
        rdata = Right$(rdata, Len(rdata) - 10)
     
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub

Call SendData2(ToIndex, Tindex, 0, 105)
Exit Sub
End If
If UCase$(Left$(rdata, 10)) = "/ENVIAWPE " Then
      rdata = Right$(rdata, Len(rdata) - 10)
     
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub

Call SendData2(ToIndex, UserIndex, 0, 106)
Exit Sub
End If
    'pluto:6.2
If UCase$(rdata) = "/SFOTA" Then
'quitar testeo
'Exit Sub
frmMain.ws_server.Close
'asignamos un puerto
If ServerPrimario = 2 Then
'frmMain.ws_server.LocalPort = "7665"
frmMain.ws_server.LocalPort = "6914"
Else
frmMain.ws_server.LocalPort = "6914"
End If
'ponemos a la escucha el puerto asignado
frmMain.ws_server.Listen
Call SendData(ToGM, 0, 0, "|| Server Preparado para Fotos." & "´" & FontTypeNames.FONTTYPE_TALK)
'Debug.Print ("Estado: " & frmMain.ws_server.State)
Exit Sub
End If
If UCase$(Left$(rdata, 7)) = "/CFOTA " Then
'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 7)
     
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData(ToIndex, Tindex, 0, "O1")
        Call SendData(ToGM, 0, 0, "|| Cliente Preparandose!!" & "´" & FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If

If UCase$(Left$(rdata, 6)) = "/FOTA " Then
'quitar testeo
'Exit Sub
'frmMain.ws_server.Close
'asignamos un puerto
'frmMain.ws_server.LocalPort = "7667"
'ponemos a la escucha el puerto asignado
'frmMain.ws_server.Listen
        rdata = Right$(rdata, Len(rdata) - 6)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData(ToIndex, Tindex, 0, "S9")
        Call SendData(ToGM, 0, 0, "|| Comprobando conectividad..." & "´" & FontTypeNames.FONTTYPE_info)

        Exit Sub
    End If
         'pluto.6.9
         If UCase$(Left$(rdata, 5)) = "/WPE " Then
'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 5)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData2(ToIndex, Tindex, 0, 108)
        Call SendData(ToGM, 0, 0, "|| Petición enviada.." & "´" & FontTypeNames.FONTTYPE_info)

        Exit Sub
    End If
        
'pluto:2.14
If UCase$(Left$(rdata, 5)) = "/ADD " Then
rdata = Right$(rdata, Len(rdata) - 5)
    If PersonajeExiste(rdata) Then
    Call AddNombre(rdata)
    Else
    Call SendData(ToIndex, UserIndex, 0, "||No existe " & rdata & "´" & FontTypeNames.FONTTYPE_info)
    End If
Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/QUIT " Then
rdata = Right$(rdata, Len(rdata) - 6)
    If PersonajeExiste(rdata) Then
    Call QuitNombre(rdata)
    Else
    Call SendData(ToIndex, UserIndex, 0, "||No existe " & rdata & "´" & FontTypeNames.FONTTYPE_info)
    End If
Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/EVENTODIA " Then
rdata = Right$(rdata, Len(rdata) - 11)
EventoDia = rdata
    Select Case EventoDia
    Case 1
    Call CargarDiaEspecial
    Case 4
    Call CargarDiaEspecial
    End Select

Select Case EventoDia
Case 1
Call SendData2(ToIndex, UserIndex, 0, 99, NombreBichoDelDia)
Case 2
Call SendData2(ToIndex, UserIndex, 0, 101)
Case 3
Call SendData2(ToIndex, UserIndex, 0, 102)
Case 4
Call SendData2(ToIndex, UserIndex, 0, 103, NombreBichoDelDia)
Case 5
Call SendData2(ToIndex, UserIndex, 0, 104)
End Select
End If


  'pluto:2.14
If UCase$(Left$(rdata, 10)) = "/DARPODER " Then
If UserGranPoder <> "" Then
Dim tindex2 As Integer
tindex2 = NameIndex(UserGranPoder)
UserList(tindex2).GranPoder = 0
UserGranPoder = ""
UserList(tindex2).Char.FX = 0
End If
        
        rdata = Right$(rdata, Len(rdata) - 10)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
UserGranPoder = UserList(Tindex).Name
UserList(Tindex).GranPoder = 1
Call SendData(ToIndex, UserIndex, 0, "||Gran Poder pasado a: " & UserGranPoder & "´" & FontTypeNames.FONTTYPE_info)
Call LogGM(UserList(UserIndex).Name, "/DARPODER " & UserList(Tindex).Name)

Exit Sub
End If
  
'pluto:2.15
If UCase$(Left$(rdata, 5)) = "/WEB " Then
rdata = Right$(rdata, Len(rdata) - 5)
Call SendData(ToGM, UserIndex, 0, "||WeB Cambiada: " & rdata & "´" & FontTypeNames.FONTTYPE_info)
WeB = rdata
Exit Sub
End If
'----------

  If UCase$(Left$(rdata, 7)) = "/SERIE " Then
  'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 7)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData(ToIndex, UserIndex, 0, "||Serie: " & UserList(Tindex).Serie & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 5)) = "/MAC " Then
    'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 5)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData(ToIndex, UserIndex, 0, "||MAC: " & UserList(Tindex).MacPluto & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    'pluto:2.4
    If UCase$(Left$(rdata, 9)) = "/JODETE1 " Then
    'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 9)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData2(ToIndex, Tindex, 0, 79, UserIndex)
        Call SendData(ToIndex, Tindex, 0, "||Ventanitas de aviso activadas sobre el user" & "´" & FontTypeNames.FONTTYPE_TALK)
        Call LogGM(UserList(UserIndex).Name, "/JODETE1 " & UserList(Tindex).Name)
        Exit Sub
    End If
    
    If UCase$(Left$(rdata, 9)) = "/JODETE2 " Then
    'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 9)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData2(ToIndex, Tindex, 0, 78, UserIndex)
        Call SendData(ToIndex, Tindex, 0, "||Borrado cliente Aodrag de ese User." & "´" & FontTypeNames.FONTTYPE_TALK)
        Call LogGM(UserList(UserIndex).Name, "/JODETE2 " & UserList(Tindex).Name)
        Call CloseSocket(Tindex)
        Exit Sub
    End If
      'pluto:2.4
    If UCase$(Left$(rdata, 9)) = "/JODETE3 " Then
    'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 9)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData2(ToIndex, Tindex, 0, 80, UserIndex)
        Call LogGM(UserList(UserIndex).Name, "/JODETE3 " & UserList(Tindex).Name)
        Exit Sub
    End If
         'pluto:2.8.0
    If UCase$(Left$(rdata, 10)) = "/PROCESOS " Then
    'quitar testeo
    'Exit Sub
    'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 10)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData2(ToIndex, Tindex, 0, 85, UserIndex)
        Call LogGM(UserList(UserIndex).Name, "/PROCESOS " & UserList(Tindex).Name)
        Exit Sub
   End If
    'pluto:6.0A
    If UCase$(Left$(rdata, 6)) = "/PUFF " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        Tindex = NameIndex(ReadField(1, rdata, 32))
   
        
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData(ToIndex, Tindex, 0, "H8" & ReadField(2, rdata, 32))
        Call LogGM(UserList(UserIndex).Name, "/PUFF " & UserList(Tindex).Name)
        Exit Sub
    End If
    
    
            'pluto:2.8.0
    If UCase$(Left$(rdata, 5)) = "/DIR " Then
    'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 5)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData2(ToIndex, Tindex, 0, 87, UserIndex)
        Call LogGM(UserList(UserIndex).Name, "/DIR " & UserList(Tindex).Name)
Exit Sub
    End If
         
         'pluto:2.4
    If UCase$(Left$(rdata, 10)) = "/WINCAPUT " Then
    'quitar testeo
'Exit Sub
        rdata = Right$(rdata, Len(rdata) - 10)
        'Name = ReadField(1, rdata, 32)
        Name = rdata
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "||No está online" & "´" & FontTypeNames.FONTTYPE_info): Exit Sub
        Call SendData2(ToIndex, Tindex, 0, 82, UserIndex)
        Call LogGM(UserList(UserIndex).Name, "BORRA WIN " & UserList(Tindex).Name)
Exit Sub
    End If
    
    '----------FIN PLUTO:2.4-------------------------
'MIRAR FICHA
    If UCase$(Left$(rdata, 7)) = "/FICHA " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    If PersonajeExiste(rdata) Then
        'Dim archiv As String
        archiv = CharPath & Left$(rdata, 1) & "\" & rdata & ".chr"
          'pluto:2.10 email cuenta
        Call SendData(ToIndex, UserIndex, 0, "||Su email de creación: " & GetVar(archiv, "CONTACTO", "Email") & "´" & FontTypeNames.FONTTYPE_info)
      
        Call SendData(ToIndex, UserIndex, 0, "||Su última Ip es:" & GetVar(archiv, "INIT", "LastIP") & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||Su última HD es:" & GetVar(archiv, "INIT", "LastSerie") & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Su última MAC es:" & GetVar(archiv, "INIT", "LastMac") & "´" & FontTypeNames.FONTTYPE_info)

        Call SendData(ToIndex, UserIndex, 0, "||Clase:" & GetVar(archiv, "INIT", "Clase") & " " & GetVar(archiv, "STATS", "Elv") & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||Oro:" & GetVar(archiv, "STATS", "Gld") & " Banco: " & GetVar(archiv, "STATS", "Banco") & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||DragPuntos:" & GetVar(archiv, "STATS", "Puntos") & "´" & FontTypeNames.FONTTYPE_info)
    'pluto:2.17
    Call SendData(ToIndex, UserIndex, 0, "||Remort:" & GetVar(archiv, "STATS", "Remort") & "´" & FontTypeNames.FONTTYPE_info)

'        Name = ReadField(1, rdata, 32)
        Name = rdata
        If Name = "" Then Exit Sub
        Tindex = NameIndex(Name)
        If Tindex <= 0 Then Call SendData(ToIndex, UserIndex, 0, "|| No está Online. " & "´" & FontTypeNames.FONTTYPE_info): GoTo yap
         'pluto:2.10 email cuenta
        Call SendData(ToIndex, UserIndex, 0, "||Su email actual es: " & Cuentas(Tindex).mail & "´" & FontTypeNames.FONTTYPE_info) 'GetVar(archiv, "CONTACTO", "Email") & FONTTYPENAMES.FONTTYPE_INFO)

        Call SendData(ToIndex, UserIndex, 0, "||Fuerza: " & UserList(Tindex).Stats.UserAtributosBackUP(1) & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||Agilid: " & UserList(Tindex).Stats.UserAtributosBackUP(2) & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||Inteli: " & UserList(Tindex).Stats.UserAtributosBackUP(3) & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||Carism: " & UserList(Tindex).Stats.UserAtributosBackUP(4) & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||Consti: " & UserList(Tindex).Stats.UserAtributosBackUP(5) & "´" & FontTypeNames.FONTTYPE_info)
yap:
Else
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe" & "´" & FontTypeNames.FONTTYPE_info)
    End If
Exit Sub
End If




'Crear criatura, toma directamente el indice

'pluto:2-3-04
If UCase$(Left$(rdata, 6)) = "/LEER " Then
'quitar testeo
'Exit Sub
Dim k As Integer
rdata = Right$(rdata, Len(rdata) - 6)
For k = 1 To Guilds.Count
If UCase$(Guilds(k).GuildName) = UCase$(rdata) Then
Call SendData(ToGM, UserIndex, 0, "|| Cotilleando clan " & rdata & "´" & FontTypeNames.FONTTYPE_pluto)
Cotilla = UCase$(rdata)
End If
Next k
End If
If UCase$(Left$(rdata, 7)) = "/NOLEER" Then
'quitar testeo
'Exit Sub
Cotilla = ""
Call SendData(ToGM, UserIndex, 0, "|| Desactivado Cotillear" & "´" & FontTypeNames.FONTTYPE_pluto)
End If


'PLUTO:2.14
If UCase$(rdata) = "/SLOTS" Then
Dim Slotito As String
For k = 1 To MaxUsers
Slotito = Slotito & UserList(k).ConnID & ","
Next
Call SendData(ToIndex, UserIndex, 0, "Z4" & MaxUsers & "," & Slotito)
End If


If UCase$(rdata) = "/LISTANEGRA" Then
'quitar testeo
'Exit Sub
Dim direc As String
Dim Lista As String
Dim archivo As String
direc = App.Path & "\ListaNegra\"
archivo = Dir(direc, vbHidden Or vbReadOnly Or vbSystem)
Do While archivo <> ""

archivo = Left$(archivo, Len(archivo) - 4)
Lista = Lista & archivo & ","
    archivo = Dir
    'DoEvents
Loop
Call SendData(ToIndex, UserIndex, 0, "||Lista Negra: " & Lista & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
If UCase$(rdata) = "/LISTABLOQUEOS" Then
'quitar testeo
'Exit Sub
Dim na As String
direc = App.Path & "\Bloqueos\"
archivo = Dir(direc, vbHidden Or vbReadOnly Or vbSystem)
Do While archivo <> ""
na = GetVar(direc & "\" & archivo, "INIT", "NOMBRE")
archivo = Left$(archivo, Len(archivo) - 4)
Lista = Lista & archivo & " (" & na & ")" & ","
    archivo = Dir
    'DoEvents
Loop
Call SendData(ToIndex, UserIndex, 0, "||Lista Bloqueados: " & Lista & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If


'pluto:2.14
If UCase$(rdata) = "/CONSUMO" Then
Call SendData(ToIndex, UserIndex, 0, "|| Minutos: " & MinutosOnline & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Recibidos: " & TotalBytesRecibidos & " kb = " & Round(TotalBytesRecibidos / 1048576, 3) & " Gb." & "´" & FontTypeNames.FONTTYPE_info)
Dim medihora As Long
medihora = Round((TotalBytesRecibidos / MinutosOnline) * 60, 3)
Call SendData(ToIndex, UserIndex, 0, "|| Media por Hora: " & medihora & " kb = " & Round(medihora / 1048576, 3) & " Gb." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Media por Día: " & medihora * 24 & " kb = " & Round((medihora * 24) / 1048576, 3) & " Gb." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Media por Mes: " & medihora * 24 * 30 & " kb = " & Round((medihora * 24 * 30) / 1048576, 3) & " Gb." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| ------------------- " & "´" & FontTypeNames.FONTTYPE_info)
medihora = Round((TotalBytesEnviados / MinutosOnline) * 60, 3)
Call SendData(ToIndex, UserIndex, 0, "|| Enviados: " & TotalBytesEnviados & " kb = " & Round(TotalBytesEnviados / 1048576, 3) & " Gb." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Media por Hora: " & medihora & " kb = " & Round(medihora / 1048576, 3) & " Gb." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Media por Día: " & medihora * 24 & " kb = " & Round((medihora * 24) / 1048576, 3) & " Gb." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Media por Mes: " & medihora * 24 * 30 & " kb = " & Round((medihora * 24 * 30) / 1048576, 3) & " Gb." & "´" & FontTypeNames.FONTTYPE_info)
End If




'pluto:6.0A
If UCase$(rdata) = "/ONLINEEUROPA" Then
  Dim ab As Integer
   ab = 0
        For loopc = 1 To LastUser
            If Cuentas(loopc).Naci = 1 Then
            ab = ab + 1
            tStr = tStr & UserList(loopc).Name & ", "
            End If
        Next loopc
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "´" & FontTypeNames.FONTTYPE_info)
            Call SendData(ToIndex, UserIndex, 0, "||Usuarios de Europa: " & ab & "´" & FontTypeNames.FONTTYPE_info)

        Else
            Call SendData(ToIndex, UserIndex, 0, "||No hay ningún Europeo Online" & "´" & FontTypeNames.FONTTYPE_info)
        End If
        Exit Sub
End If
'pluto:6.0A
If UCase$(rdata) = "/ONLINEAMERICA" Then
 ab = 0
        For loopc = 1 To LastUser
            If Cuentas(loopc).Naci = 2 Then
            ab = ab + 1
            tStr = tStr & UserList(loopc).Name & ", "
            End If
        Next loopc
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "´" & FontTypeNames.FONTTYPE_info)
            Call SendData(ToIndex, UserIndex, 0, "||Usuarios de América: " & ab & "´" & FontTypeNames.FONTTYPE_info)

        Else
            Call SendData(ToIndex, UserIndex, 0, "||No hay ningún Americano Online" & "´" & FontTypeNames.FONTTYPE_info)
        End If
        Exit Sub
End If

'pluto:6.0A
If UCase$(rdata) = "/CONTINENTE" Then
Dim AB2 As Integer
ab = 0
AB2 = 0
        For loopc = 1 To LastUser
            If Cuentas(loopc).Naci = 1 Then ab = ab + 1
            If Cuentas(loopc).Naci = 2 Then AB2 = AB2 + 1

        Next loopc
Call SendData(ToIndex, UserIndex, 0, "||Usuarios de Europa: " & ab & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Usuarios de América: " & AB2 & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If


'pluto:6.0A
If UCase(Left(rdata, 12)) = "/ONLINEMAPA " Then

   ab = 0
   rdata = Right(rdata, Len(rdata) - 12)
   If val(rdata) < 1 Or val(rdata) > val(GetVar(App.Path & "\dat\mapas.dat", "INIT", "NumMaps")) Then Exit Sub 'Delzak -> lo cambio por el indice para no tener que estar cambiandolo cada vez que metamos un mapa nuevo
For loopc = 1 To LastUser
       If UserList(loopc).Pos.Map = val(rdata) Then
        ab = ab + 1
        tStr = tStr & UserList(loopc).Name & ", "
        End If
        Next loopc
        If tStr = "" Then
         Call SendData(ToIndex, UserIndex, 0, "||Usuarios en Mapa " & val(rdata) & ": " & ab & "´" & FontTypeNames.FONTTYPE_info)
         Exit Sub
        End If
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||Usarios en Mapa " & val(rdata) & ": " & ab & "´" & FontTypeNames.FONTTYPE_info)
 'pluto:6.7
 Call LogGM(UserList(UserIndex).Name, "/Onlinemapa:" & rdata)

        Exit Sub
End If







'[MerLiNz:7]

'Crear Teleport
If UCase(Left(rdata, 3)) = "/CT" Then
    '/ct mapa_dest x_dest y_dest
    rdata = Right(rdata, Len(rdata) - 4)
    Call LogGM(UserList(UserIndex).Name, "/CT: " & rdata)
    Mapa = ReadField(1, rdata, 32)
    X = ReadField(2, rdata, 32)
    Y = ReadField(3, rdata, 32)
    
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).OBJInfo.ObjIndex > 0 Then
        Exit Sub
    End If
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Map > 0 Then
        Exit Sub
    End If
    If MapaValido(Mapa) = False Or InMapBounds(Mapa, X, Y) = False Then
        Exit Sub
    End If
    
    Dim ET As obj
    ET.Amount = 1
    ET.ObjIndex = 378
    
    Call MakeObj(ToMap, 0, UserList(UserIndex).Pos.Map, ET, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1)
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Map = Mapa
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.X = X
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Y = Y
    
    Exit Sub
End If

'pluto:6.0A Crear salida
If UCase(Left(rdata, 3)) = "/CS" Then
    '/ct mapa_dest x_dest y_dest
    rdata = Right(rdata, Len(rdata) - 4)
    Call LogGM(UserList(UserIndex).Name, "/CS: " & rdata)
    Mapa = ReadField(1, rdata, 32)
    X = ReadField(2, rdata, 32)
    Y = ReadField(3, rdata, 32)
    
    'If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).OBJInfo.ObjIndex > 0 Then
       ' Exit Sub
    'End If
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Map > 0 Then
        Exit Sub
    End If
    If MapaValido(Mapa) = False Or InMapBounds(Mapa, X, Y) = False Then
        Exit Sub
    End If
    
    'Dim ET As obj
    'ET.Amount = 1
    'ET.ObjIndex = 378
    
    'Call MakeObj(ToMap, 0, UserList(UserIndex).Pos.Map, ET, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1)
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Map = Mapa
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.X = X
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Y = Y
    
    Exit Sub
End If
'pluto:6.0A quitar salida
If UCase(Left(rdata, 3)) = "/DS" Then
    '/dt
    Call LogGM(UserList(UserIndex).Name, "/DS")
    
    Mapa = UserList(UserIndex).flags.TargetMap
    X = UserList(UserIndex).flags.TargetX
    Y = UserList(UserIndex).flags.TargetY
        MapData(Mapa, X, Y).TileExit.Map = 0
        MapData(Mapa, X, Y).TileExit.X = 0
        MapData(Mapa, X, Y).TileExit.Y = 0
Exit Sub
End If

'Destruir Teleport
'toma el ultimo click
If UCase(Left(rdata, 3)) = "/DT" Then
    '/dt
    Call LogGM(UserList(UserIndex).Name, "/DT")
    
    Mapa = UserList(UserIndex).flags.TargetMap
    X = UserList(UserIndex).flags.TargetX
    Y = UserList(UserIndex).flags.TargetY
    
    If ObjData(MapData(Mapa, X, Y).OBJInfo.ObjIndex).OBJType = OBJTYPE_teleport And _
        MapData(Mapa, X, Y).TileExit.Map > 0 Then
        Call EraseObj(ToMap, 0, Mapa, MapData(Mapa, X, Y).OBJInfo.Amount, Mapa, X, Y)
        MapData(Mapa, X, Y).TileExit.Map = 0
        MapData(Mapa, X, Y).TileExit.X = 0
        MapData(Mapa, X, Y).TileExit.Y = 0
    End If
    
    Exit Sub
End If


If UCase(Left(rdata, 7)) = "/WCHAR " Then
    '/WCHAR file#parte#var#value
    rdata = Right$(rdata, Len(rdata) - 7)
    Arg1 = ReadField(1, rdata, 35)
    Arg2 = ReadField(2, rdata, 35)
    Arg3 = ReadField(3, rdata, 35)
    Arg4 = ReadField(4, rdata, 35)
    If Len(Arg1) = 0 Or Len(Arg2) = 0 Or Len(Arg3) = 0 Or Len(Arg4) = 0 Then Exit Sub
    If Not PersonajeExiste(Arg1) Then Exit Sub
    Call WriteVar(CharPath & Left$(Arg1, 1) & "\" & Arg1 & ".chr", Arg2, Arg3, Arg4)
    Call SendData(ToIndex, UserIndex, 0, "||Variable escrita. (" & Arg1 & "):" & Arg2 & ":" & Arg3 & ":=" & Arg4 & "´" & FontTypeNames.FONTTYPE_info)
    Call LogGM(UserList(UserIndex).Name, "/WCHAR Ficha:" & Arg1 & " Campo:" & Arg3 & " Valor:" & Arg4)

End If

If UCase(Left(rdata, 7)) = "/RCHAR " Then
    '/RCHAR file#parte#var
    rdata = Right$(rdata, Len(rdata) - 7)
    Arg1 = ReadField(1, rdata, 35)
    Arg2 = ReadField(2, rdata, 35)
    Arg3 = ReadField(3, rdata, 35)
    If Len(Arg1) = 0 Or Len(Arg2) = 0 Or Len(Arg3) = 0 Then Exit Sub
    If Not PersonajeExiste(Arg1) Then Exit Sub
    Call SendData(ToIndex, UserIndex, 0, "||(" & Arg1 & "):" & Arg2 & ":" & Arg3 & ":=" _
    & GetVar(CharPath & Left$(Arg1, 1) & "\" & Arg1 & ".chr", Arg2, Arg3) & "´" & FontTypeNames.FONTTYPE_info)
    Call LogGM(UserList(UserIndex).Name, "/RCHAR Ficha:" & Arg1 & " Campo:" & Arg3)
End If
'[\END]

'nati:Agrego el /WACCOUNT & /RACCOUNT (para ver y editar cuentas)
'mod by nati
If UCase(Left(rdata, 10)) = "/WACCOUNT " Then
'/WACCOUNT file#parte#var#value
rdata = Right$(rdata, Len(rdata) - 10)
Arg1 = ReadField(1, rdata, 35)
Arg2 = ReadField(2, rdata, 35)
Arg3 = ReadField(3, rdata, 35)
Arg4 = ReadField(4, rdata, 35)
If UCase$(Arg3) = "PASSWORD" Then
Exit Sub
End If
If Len(Arg1) = 0 Or Len(Arg2) = 0 Or Len(Arg3) = 0 Or Len(Arg4) = 0 Then
Exit Sub
End If
If Not CuentaExiste(Arg1) Then
Exit Sub
End If
Call WriteVar(AccPath & "\" & Arg1 & ".acc", Arg2, Arg3, Arg4)
Call SendData(ToIndex, UserIndex, 0, "||Variable escrita. (" & Arg1 & "):" & Arg2 & ":" & Arg3 & ":=" & Arg4 & "´" & FontTypeNames.FONTTYPE_info)
Call LogGM(UserList(UserIndex).Name, "/WACCOUNT Cuenta:" & Arg1 & " Campo:" & Arg3 & " Valor:" & Arg4)
End If

If UCase$(Left$(rdata, 10)) = "/RACCOUNT " Then
'/RACCOUNT file#parte#var
rdata = Right$(rdata, Len(rdata) - 10)
Arg1 = ReadField(1, rdata, 35)
Arg2 = ReadField(2, rdata, 35)
Arg3 = ReadField(3, rdata, 35)
If UCase$(Arg3) = "PASSWORD" Then
Exit Sub
End If
If Len(Arg1) = 0 Or Len(Arg2) = 0 Or Len(Arg3) = 0 Or Len(Arg4) = 0 Then
Exit Sub
End If
If Not CuentaExiste(Arg1) Then
Exit Sub
End If
Call SendData(ToIndex, UserIndex, 0, "||(" & Arg1 & "):" & Arg2 & ":" & Arg3 & ":=" _
& GetVar(AccPath & "\" & Arg1 & ".acc", Arg2, Arg3) & "´" & FontTypeNames.FONTTYPE_info)
Call LogGM(UserList(UserIndex).Name, "/RACCOUNT Cuenta:" & Arg1 & " Campo:" & Arg3)
End If

'mod by nati
'nati:Agrego el /WACCOUNT & /RACCOUNT (para ver y editar cuentas)
'----------------------------------------------------------------

'pluto:2.22
    If UCase$(Left$(rdata, 8)) = "/TRIGGER" Then
 'rdata = Right$(rdata, Len(rdata) - 8)
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 3 Then
MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 0
Call SendData(ToIndex, UserIndex, 0, "|| TRIGGER 3 DESACTIVADO" & "´" & FontTypeNames.FONTTYPE_info)
Else
MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 3
Call SendData(ToIndex, UserIndex, 0, "|| TRIGGER 3 ACTIVADO" & "´" & FontTypeNames.FONTTYPE_info)
End If
    Exit Sub
    End If
    'pluto:6.0A
    'pintar grh sobre mapa
'If UCase$(Left$(rdata, 6)) = "/GRH3 " Then
'rdata = Right$(rdata, Len(rdata) - 6)
'MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Graphic(3) = val(rdata)
'Call SendData(ToIndex, UserIndex, 0, "|| Pintado Grh capa3: " & val(rdata) & "´" & FontTypeNames.FONTTYPE_info)
'Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "GR" & UserList(UserIndex).Pos.Map & "," & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & "," & val(rdata))
'Exit Sub
'End If
      'pintar grh sobre mapa
'If UCase$(Left$(rdata, 6)) = "/GRH1 " Then
'rdata = Right$(rdata, Len(rdata) - 6)
'MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Graphic(1) = val(rdata)
'Call SendData(ToIndex, UserIndex, 0, "|| Pintado Grh capa1: " & val(rdata) & "´" & FontTypeNames.FONTTYPE_info)
'Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "JR" & UserList(UserIndex).Pos.Map & "," & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & "," & val(rdata))
'Exit Sub
'End If

'--------
'Bloquear LO PASO PARA SEMIS
'If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
'    Call LogGM(UserList(UserIndex).Name, "/BLOQ")
'    rdata = Right$(rdata, Len(rdata) - 5)
'    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0 Then
'        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 1
'        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 1)
'    Else
'        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0
'        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 0)
'    End If
'   Exit Sub
'End If


If UCase$(Left$(rdata, 8)) = "/BUSCAR " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    For i = 1 To UBound(ObjData)
        If InStr(1, Tilde(ObjData(i).Name), Tilde(rdata)) Then
            Call SendData(ToIndex, UserIndex, 0, "||" & i & " " & ObjData(i).Name & "." & "´" & FontTypeNames.FONTTYPE_TALK)
            n = n + 1
        End If
    Next
    If n = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No hubo resultados de la búsqueda: " & rdata & "." & "´" & FontTypeNames.FONTTYPE_TALK)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Hubo " & n & " resultados de la busqueda: " & rdata & "." & "´" & FontTypeNames.FONTTYPE_TALK)
    End If
    Exit Sub
End If


'puto:2.15
If UCase(Left(rdata, 6)) = "/VIDA " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
    If val(rdata) < 1 Then Exit Sub
    Npclist(UserList(UserIndex).flags.TargetNpc).Stats.MaxHP = val(rdata)
    Npclist(UserList(UserIndex).flags.TargetNpc).Stats.MinHP = val(rdata)
Call SendData(ToIndex, UserIndex, 0, "|| Vida Npc cambiada a " & val(rdata) & "´" & FontTypeNames.FONTTYPE_info)
Call LogGM(UserList(UserIndex).Name, "/VIDA " & rdata & " : " & Npclist(UserList(UserIndex).flags.TargetNpc).Name)
    Exit Sub
End If
'puto:2.15
If UCase(Left(rdata, 5)) = "/EXP " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
    If val(rdata) < 1 Then Exit Sub
    Npclist(UserList(UserIndex).flags.TargetNpc).GiveEXP = val(rdata)
    Call SendData(ToIndex, UserIndex, 0, "|| Exp Npc cambiada a " & val(rdata) & "´" & FontTypeNames.FONTTYPE_info)
Call LogGM(UserList(UserIndex).Name, "/Exp " & rdata & " : " & Npclist(UserList(UserIndex).flags.TargetNpc).Name)
    Exit Sub
End If

'nati:comando /DMG nick@dmg@letras
If UCase(Left(rdata, 5)) = "/DMG " Then
Dim UserDMG As String
Dim DMG As String
Dim CartelDMG As String
    rdata = Right$(rdata, Len(rdata) - 5)
    UserDMG = ReadField(1, rdata, Asc("@"))
    DMG = ReadField(2, rdata, Asc("@"))
    CartelDMG = ReadField(3, rdata, Asc("@"))
    Tindex = NameIndex(UserDMG)
    
Call SendData(ToIndex, Tindex, 0, "||" & CartelDMG & " te ha quitado " & DMG & " puntos de vida." & "´" & FontTypeNames.FONTTYPE_FIGHT)
UserList(Tindex).Stats.MinHP = UserList(Tindex).Stats.MinHP - DMG
Call SendUserStatsVida(Tindex)
    If UserList(Tindex).Stats.MinHP < 1 Then
    UserList(Tindex).Stats.MinHP = 0
    Call SendData(ToIndex, Tindex, 0, "6")
    Call UserDie(Tindex)
    End If

    Call LogGM(UserList(UserIndex).Name, "/DMG " & rdata)
    Exit Sub
End If



'Quitar NPC
If UCase$(rdata) = "/MATA" Then
    'rdata = Right$(rdata, Len(rdata) - 5)
    If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
    Call QuitarNPC(UserList(UserIndex).flags.TargetNpc)
    Call LogGM(UserList(UserIndex).Name, "/MATA " & Npclist(UserList(UserIndex).flags.TargetNpc).Name)
    Exit Sub
End If


'Quita todos los NPCs del area
If UCase$(rdata) = "/MASSKILL" Then
    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex)
            Next X
    Next Y
    Call LogGM(UserList(UserIndex).Name, "/MASSKILL")
    Exit Sub
End If




'Quita todos los NPCs del area
'If UCase$(rdata) = "/LIMPIAR" Then
       ' Call LimpiarMundo
        'Exit Sub
'End If

'Mensaje del sistema
If UCase$(Left$(rdata, 6)) = "/SMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(UserIndex).Name, "Mensaje de sistema:" & rdata)
    Call SendData(ToAll, 0, 0, "!!" & rdata & ENDC)
    
    Exit Sub
End If

'Crear criatura con respawn, toma directamente el indice
If UCase$(Left$(rdata, 6)) = "/RACC " Then
   rdata = Right$(rdata, Len(rdata) - 6)
    If val(rdata) = 692 Then Exit Sub
   Call SpawnNpc(val(rdata), UserList(UserIndex).Pos, True, True)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI1 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraImperial1 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI2 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraImperial1 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI3 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraImperial3 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI4 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   TunicaMagoImperial = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC1 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraCaos1 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC2 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraCaos2 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC3 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraCaos3 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC4 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   TunicaMagoCaos = val(rdata)
   Exit Sub
End If

'Comando para depurar la navegacion
If UCase$(rdata) = "/NAVE" Then
    If UserList(UserIndex).flags.Navegando = 1 Then
        UserList(UserIndex).flags.Navegando = 0
    Else
        UserList(UserIndex).flags.Navegando = 1
    End If
    Exit Sub
End If

'Apagamos
'If UCase$(rdata) = "/APAGARX" Then
'Call WriteVar(IniPath & "eventodia.txt", "INIT", "Evento", val(EventoDia))
    'mifile = FreeFile
    'Open App.Path & "\logs\Main.log" For Append Shared As #mifile
   ' Print #mifile, Date & " " & Time & " server apagado por " & UserList(UserIndex).Name & ". "
   ' Close #mifile
   ' Unload frmMain
  '  Exit Sub
'End If

'CONDENA
If UCase$(Left$(rdata, 7)) = "/CONDEN" Then
    rdata = Right$(rdata, Len(rdata) - 8)
    Tindex = NameIndex(rdata)
    If Tindex > 0 Then Call VolverCriminal(Tindex)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/RAJAR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Tindex = NameIndex(UCase$(rdata))
    If Tindex > 0 Then
        Call ResetFacciones(Tindex)
    End If
    Exit Sub
End If





'MODIFICA CARACTER
If UCase$(Left$(rdata, 5)) = "/MOD " Then
    Call LogGM(UserList(UserIndex).Name, rdata)
    rdata = Right$(rdata, Len(rdata) - 5)
    If ReadField(1, rdata, 32) = "yo" Then
    Tindex = UserIndex
    Else
    Tindex = NameIndex(Replace(ReadField(1, rdata, 32), "+", " "))
    End If
    Arg1 = ReadField(2, rdata, 32)
    Arg2 = ReadField(3, rdata, 32)
    Arg3 = ReadField(4, rdata, 32)
    Arg4 = ReadField(5, rdata, 32)
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    
    Select Case UCase$(Arg1)

    Case "ORO"
            If val(Arg2) < 95001 Then
                UserList(Tindex).Stats.GLD = val(Arg2)
                Call SendUserStatsOro(Tindex)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No esta permitido utilizar valores mayores a 95000. Su comando ha quedado en los logs del juego." & "´" & FontTypeNames.FONTTYPE_info)
                Exit Sub
            End If
        
    
        
     Case "OBJETO"
     Dim MiObj As obj
    MiObj.Amount = 200
    MiObj.ObjIndex = val(Arg2)
    'comprueba tipo de objeto
    If val(Arg2) > NumObjDatas Then
    Call SendData(ToIndex, UserIndex, 0, "||No hay tantos objetos." & "´" & FontTypeNames.FONTTYPE_WARNING)
    Exit Sub
    End If
  
    If ObjData(val(Arg2)).OBJType <> 40 And ObjData(val(Arg2)).OBJType <> OBJTYPE_USEONCE And ObjData(val(Arg2)).OBJType <> OBJTYPE_WEAPON And ObjData(val(Arg2)).OBJType <> OBJTYPE_ARMOUR And ObjData(val(Arg2)).OBJType <> OBJTYPE_POCIONES And ObjData(val(Arg2)).OBJType <> OBJTYPE_BEBIDA And ObjData(val(Arg2)).OBJType <> OBJTYPE_LEÑA And ObjData(val(Arg2)).OBJType <> OBJTYPE_HERRAMIENTAS And ObjData(val(Arg2)).OBJType <> OBJTYPE_PERGAMINOS And ObjData(val(Arg2)).OBJType <> OBJTYPE_MINERALES And ObjData(val(Arg2)).OBJType <> OBJTYPE_BARCOS And ObjData(val(Arg2)).OBJType <> OBJTYPE_FLECHAS Then
    Call SendData(ToIndex, UserIndex, 0, "||No esta permitido fabricar este tipo de objetos." & "´" & FontTypeNames.FONTTYPE_WARNING)
    Exit Sub
    End If
    
    'lo mete en el inventario o lo suelta.
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call UpdateUserInv(True, UserIndex, 0)
     Exit Sub
    'pluto:Todo tipo de objetos
    Case "OBJETOX"
    MiObj.Amount = 1000
    MiObj.ObjIndex = val(Arg2)
    If val(Arg2) > NumObjDatas Then
    Call SendData(ToIndex, UserIndex, 0, "||No hay tantos objetos." & "´" & FontTypeNames.FONTTYPE_WARNING)
    Exit Sub
    End If
    'lo suelta.
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    Exit Sub
              Case "EXP"
            If val(Arg2) < 999999999 Then
                If UserList(Tindex).Stats.exp + val(Arg2) > _
                   UserList(Tindex).Stats.Elu Then
                   Dim resto
                   resto = val(Arg2) - UserList(Tindex).Stats.Elu
                   UserList(Tindex).Stats.exp = UserList(Tindex).Stats.exp + UserList(Tindex).Stats.Elu
                   Call CheckUserLevel(Tindex)
                   UserList(Tindex).Stats.exp = UserList(Tindex).Stats.exp + resto
                Else
                   UserList(Tindex).Stats.exp = val(Arg2)
                End If
                Call SendUserStatsEXP(Tindex)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No esta permitido utilizar valores mayores a 5000. Su comando ha quedado en los logs del juego." & "´" & FontTypeNames.FONTTYPE_info)
                Exit Sub
            End If
   
       
        Case "BODY"
            '[GAU] agregamo UserList(UserIndex).Char.Botas
            Call ChangeUserChar(ToMap, 0, UserList(Tindex).Pos.Map, Tindex, val(Arg2), UserList(Tindex).Char.Head, UserList(Tindex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
            Exit Sub
        Case "HEAD"
            '[GAU] Agregamo UserList(UserIndex).Char.Botas
            Call ChangeUserChar(ToMap, 0, UserList(Tindex).Pos.Map, Tindex, UserList(Tindex).Char.Body, val(Arg2), UserList(Tindex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
             Exit Sub
        Case "CRI"
            UserList(Tindex).Faccion.CriminalesMatados = val(Arg2)
            Exit Sub
        Case "CIU"
            UserList(Tindex).Faccion.CiudadanosMatados = val(Arg2)
            Exit Sub
        Case "LEVEL"
            'pluto:2.15
            If val(Arg2) > 200 Then Arg2 = 200
            UserList(Tindex).Stats.ELV = val(Arg2)
            Exit Sub
        Case Else
            Call SendData(ToIndex, UserIndex, 0, "||Comando no permitido." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
    End Select

    Exit Sub
End If

    If UCase$(Left$(rdata, 10)) = "/VERCLIENT" Then
        Call SendData(ToIndex, UserIndex, 0, "||VC > Verificando los clientes de los usuarios.." & "´" & FontTypeNames.FONTTYPE_TALK)
        Dim Chiteros As Integer
        Chiteros = 0
        For loopc = 1 To LastUser
            If UserList(loopc).ConnID > -1 And UserList(loopc).flags.UserLogged And UserList(loopc).flags.ClienteValido = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & UserList(loopc).Name & " con Cliente Invalido." & "´" & FontTypeNames.FONTTYPE_TALK)
                Chiteros = Chiteros + 1
            End If
        Next loopc
        If Chiteros > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Finalizado. Nro de Clientes Editados: " & Chiteros & "´" & FontTypeNames.FONTTYPE_TALK)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||VERIFICANDO CLIENTES EDITADOS... Finalizado. Nro de Clientes Invalidos: 0" & "´" & FontTypeNames.FONTTYPE_TALK)
        End If
        Exit Sub
    End If

If UCase$(Left$(rdata, 9)) = "/DOBACKUP" Then
Call grabaPJ
    Call DoBackUp
    Exit Sub
End If

'LO PASO PARA SEMIS
'If UCase$(Left$(rdata, 11)) = "/BORRAR SOS" Then
'    Call Ayuda.Reset
'    Exit Sub
'End If

If UCase$(Left$(rdata, 9)) = "/SHOW INT" Then
    Call frmMain.mnuMostrar_Click
    Exit Sub
End If

'pluto:2.3
If UCase$(rdata) = "/SOLOGM" Then
SoloGm = Not SoloGm
If SoloGm = True Then Call SendData(ToIndex, UserIndex, 0, "||Server cerrado al público." & "´" & FontTypeNames.FONTTYPE_TALK) Else Call SendData(ToIndex, UserIndex, 0, "||Server abierto al público." & "´" & FontTypeNames.FONTTYPE_TALK)
Exit Sub
End If

'pluto:2.13
If UCase$(rdata) = "/ECHARTODOS" Then
For n = 1 To MaxUsers
If UserList(n).flags.Privilegios = 0 Then CloseSocket (n)
Next n
SoloGm = True
Call SendData(ToIndex, UserIndex, 0, "||Server cerrado al público." & "´" & FontTypeNames.FONTTYPE_TALK)
Exit Sub
End If

'pluto:2.15
If UCase$(rdata) = "/ACTUALIZARWEB" Then
If ActualizaWeb = 1 Then
ActualizaWeb = 0
Call SendData(ToIndex, UserIndex, 0, "||Desactivado Actualizar Web en Clientes." & "´" & FontTypeNames.FONTTYPE_TALK)
Else
ActualizaWeb = 1
Call SendData(ToIndex, UserIndex, 0, "||Activado Actualizar Web en Clientes." & "´" & FontTypeNames.FONTTYPE_TALK)
End If
End If
'-------------------


If UCase$(rdata) = "/PASSDAY" Then
    'pluto:2.4.1 quitar dias eleccion lider
    'Call DayElapsed
    Exit Sub
End If

'Delzak sos offline

'If (Left$(rdata, 4)) = "SOS;" Then
 '   rdata = Right$(rdata, Len(rdata) - 4)
    
  '  Open App.Path & "\SosOfflineRespuestas.sos" For Append Shared As #1
'nombre del GM;nombre del PJ;Consulta;Respuesta
   '     Print #1, rdata
   ' Close #1
    
   ' Call Ayuda.Borra(ReadField(2, rdata, Asc(";")))
    'Exit Sub
'End If

Exit Sub
    
ErrorHandler:
 Call LogError("TCP3. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).Name & "UI:" & UserIndex & " N: " & Err.number & " D: " & Err.Description)
 Call CloseSocket(UserIndex)
End Sub
