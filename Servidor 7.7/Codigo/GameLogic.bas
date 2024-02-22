Attribute VB_Name = "Extra"
Option Explicit

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie

Exit Function
fallo:
Call LogError("ESNEWBIE" & Err.number & " D: " & Err.Description)

End Function
Public Sub ControlaSalidas(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo errhandler

Dim nPos As WorldPos
Dim FxFlag As Boolean
'Controla las salidas
If InMapBounds(Map, X, Y) Then
'pluto.6.5
  'DoEvents
     
 'pluto:6.0A
If UserList(UserIndex).Pos.Map = 274 And UserList(UserIndex).Pos.X = 42 And UserList(UserIndex).Pos.Y = 46 Then
    If UserList(UserIndex).flags.Pitag = 1 Then
MapData(Map, X, Y).TileExit.Map = 274
MapData(Map, X, Y).TileExit.X = 49
MapData(Map, X, Y).TileExit.Y = 33
    Else
MapData(Map, X, Y).TileExit.Map = 28
MapData(Map, X, Y).TileExit.X = 46
MapData(Map, X, Y).TileExit.Y = 86
    End If
End If
    
    
    
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        FxFlag = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = OBJTYPE_teleport
    End If
    
If MapData(Map, X, Y).TileExit.Map > 0 Then


'pluto:2.12
If UserList(UserIndex).Pos.Map = MapaTorneo2 And MapInfo(UserList(UserIndex).Pos.Map).NumUsers > 1 And UserList(UserIndex).Torneo2 < 10 Then
        
        Call SendData(ToIndex, UserIndex, 0, "||No puedes salir hasta que consigas 10 victorias." & "´" & FontTypeNames.FONTTYPE_info)
        'Call ClosestLegalPos(UserList(UserIndex).Pos, nPos, 0)
        'If nPos.X <> 0 And nPos.Y <> 0 Then
        Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

        'End If
     
        Exit Sub
End If


'pluto:hoy
If Map > 177 And Map < 183 Then Call SendData(ToIndex, UserIndex, 0, "TW" & 135)
If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Terreno) = "CASA" And (UserList(UserIndex).Stats.GLD < 30000 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 0 Or EsNewbie(UserIndex) Or UserList(UserIndex).NroMacotas > 0 Or UserList(UserIndex).flags.Montura > 0) Then
   
        'No llevas oro a la casa
        Call SendData(ToIndex, UserIndex, 0, "||Los espíritus no te dejan entrar si tienes menos de 30000 Monedas, eres Newbie, llevas mascotas o estás Desnudo." & "´" & FontTypeNames.FONTTYPE_info)
        'Call ClosestLegalPos(UserList(UserIndex).Pos, nPos, 0)
        'If nPos.X <> 0 And nPos.Y <> 0 Then
        Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

        'End If
     
        Exit Sub
 End If
 'pluto:6.0A añado caballero
 If (MapData(Map, X, Y).TileExit.Map = mapi Or MapData(Map, X, Y).TileExit.Map = 250) And (UserList(UserIndex).NroMacotas > 0 Or UserList(UserIndex).flags.Montura > 0) Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes acceder a esta sala con mascotas." & "´" & FontTypeNames.FONTTYPE_info)
       ' Call ClosestLegalPos(UserList(UserIndex).Pos, nPos, 0)
        'If nPos.X <> 0 And nPos.Y <> 0 Then
        'Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
        Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

        'End If
     Exit Sub
 End If
     
     'pluto:2.17 mapa conquistas
'If MapInfo(MapData(Map, X, Y).TileExit.Map).Terreno = "CONQUISTA" And UserList(UserIndex).Faccion.ArmadaReal = 0 And UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
       'Call SendData(ToIndex, UserIndex, 0, "||No estás en ninguna Armada." & FONTTYPENAMES.FONTTYPE_INFO)
        'Call ClosestLegalPos(UserList(UserIndex).Pos, nPos, 0)
        'If nPos.X <> 0 And nPos.Y <> 0 Then
        
    'Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
       ' End If
     'exit sub
     'Exit Sub
 'End If
 '-----------------
   'pluto:2-3-04
If MapInfo(MapData(Map, X, Y).TileExit.Map).StartPos.Map = 178 And MapInfo(MapData(Map, X, Y).TileExit.Map).StartPos.Y = 93 And UserList(UserIndex).Stats.ELV < 30 Then
        Call SendData(ToIndex, UserIndex, 0, "||Necesitas ser Level 30 para acceder a la Pirámide." & "´" & FontTypeNames.FONTTYPE_info)
        'Call ClosestLegalPos(UserList(UserIndex).Pos, nPos, 0)
        'If nPos.X <> 0 And nPos.Y <> 0 Then
        Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

        'End If
     Exit Sub
 End If
 
'pluto:6.8---------------------------------------

'If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Terreno) = "CASTILLO" And UserList(UserIndex).Stats.PClan < 0 Then
        'No pclan
'        Call SendData(ToIndex, UserIndex, 0, "||No tienes Puntos Clan!!" & "´" & FontTypeNames.FONTTYPE_info)
'              Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

     
     
'        Exit Sub
' End If
 '------------------------------------------------
 
 
 
 'pluto:2.17-----------------------
    'If MapInfo(MapData(Map, X, Y).TileExit.Map).Terreno <> "ALDEA" And EsNewbie(UserIndex) And UserList(UserIndex).Remort = 0 Then
'Call SendData(ToIndex, UserIndex, 0, "Z8")
'Call ClosestLegalPos(UserList(UserIndex).Pos, nPos, 0)
'If nPos.X <> 0 And nPos.Y <> 0 Then
'Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

'exit sub
'End If
    'End If
'--------------------------------
        '¿Es mapa de newbies?
        If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "SI" Then
            '¿El usuario es un newbie?
            If EsNewbie(UserIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos, 0)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                        End If
                    End If
                End If
            Else 'No es newbie
                Call SendData(ToIndex, UserIndex, 0, "||Mapa exclusivo para newbies." & "´" & FontTypeNames.FONTTYPE_info)
                
                Call ClosestLegalPos(UserList(UserIndex).Pos, nPos, 0)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                End If
            End If
        Else 'No es un mapa de newbies
            If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                If FxFlag Then
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                Else
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                End If
            Else
                Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos, 0)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                    If FxFlag Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                    End If
                End If
            End If
        End If
    End If
    
End If

Exit Sub

errhandler:
Call LogError("Error en ControlaSalidas ->Nom: " & UserList(UserIndex).Name & " POS:" & UserList(UserIndex).Pos.Map & " - " & UserList(UserIndex).Pos.X & " - " & UserList(UserIndex).Pos.Y & " N: " & Err.number & " D: " & Err.Description)


End Sub
Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)










End Sub


Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
On Error GoTo fallo
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Or Map = 0 Then
    InMapBounds = False
Else
    InMapBounds = True
End If
Exit Function
fallo:
Call LogError("INMAPBOUNDS" & Err.number & " D: " & Err.Description)


End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, agua As Byte)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************
On Error GoTo fallo
Dim Notfound As Boolean
Dim loopc As Integer
Dim tX As Integer
Dim tY As Integer
Dim pagua As Boolean

If agua = 1 Then pagua = True Else pagua = False
nPos.Map = Pos.Map
nop:
Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, pagua)

    
    If loopc > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - loopc To Pos.Y + loopc
        For tX = Pos.X - loopc To Pos.X + loopc
            'pluto:2.17 añade exits
            If tX > 99 Or tY > 99 Or tX < 1 Or tY < 1 Then GoTo nuu
            If LegalPos(nPos.Map, tX, tY, pagua) Then 'and MapData(nPos.Map, tX, tY).TileExit.Map = 0
                
          
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                tX = Pos.X + loopc
                tY = Pos.Y + loopc
            Notfound = False
            Exit Sub
            End If
nuu:
        Next tX
    Next tY
    loopc = loopc + 1
Loop



If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

Exit Sub
fallo:
Call LogError("CLOSESTLEGALPOS" & Err.number & " D: " & nPos.Map & "-" & tX & "-" & tY & pagua)

End Sub

Function NameIndex(ByVal Name As String) As Integer
On Error GoTo fallo
Dim UserIndex As Integer
'¿Nombre valido?
If Name = "" Then
    NameIndex = 0
    Exit Function
End If
UserIndex = 1
If (Right$(Name, 1) <> "$") Then
    GoTo prim
Else
    Name = Left$(Name, Len(Name) - 1)
End If

Do Until UCase$(UserList(UserIndex).Name) = UCase$(Name)
    UserIndex = UserIndex + 1
    If UserIndex > MaxUsers Then
        UserIndex = 0
        Exit Do
    End If
Loop
GoTo final
prim:

Do Until UCase$(Left$(UserList(UserIndex).Name, Len(Name))) = UCase$(Name)
    UserIndex = UserIndex + 1
    If UserIndex > MaxUsers Then
        UserIndex = 0
        Exit Do
    End If
Loop
final:
NameIndex = UserIndex

Exit Function
fallo:
Call LogError("NAMEINDEX" & Err.number & " D: " & Err.Description)

End Function


Function IP_Index(ByVal inIP As String) As Integer
On Error GoTo local_errHand

Dim UserIndex As Integer
'¿Nombre valido?
If inIP = "" Then
    IP_Index = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UserList(UserIndex).ip = inIP
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        IP_Index = 0
        Exit Function
    End If
    
Loop
 
IP_Index = UserIndex
Exit Function
local_errHand:
    IP_Index = UserIndex
Call LogError("IP INDEX" & Err.number & " D: " & Err.Description)

End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
On Error GoTo fallo
Dim loopc As Integer
For loopc = 1 To MaxUsers
    If UserList(loopc).flags.UserLogged = True Then
        If UserList(loopc).ip = UserIP And UserIndex <> loopc Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next loopc
CheckForSameIP = False
Exit Function
fallo:
Call LogError("CHECKFORSAMEIP" & Err.number & " D: " & Err.Description)


End Function

Function CheckForSameName(ByVal UserIndex As Integer, ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
On Error GoTo fallo
Dim loopc As Integer
For loopc = 1 To MaxUsers
    If UserList(loopc).flags.UserLogged Then
        If UCase$(UserList(loopc).Name) = UCase$(Name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next loopc
CheckForSameName = False

Exit Function
fallo:
Call LogError("CHECKFORSAMENAME" & Err.number & " D: " & Err.Description)

End Function

Sub HeadtoPos(Head As Byte, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
On Error GoTo fallo
Dim X As Integer
Dim Y As Integer
Dim tempVar As Single
Dim nx As Integer
Dim nY As Integer

X = Pos.X
Y = Pos.Y

If Head = NORTH Then
    nx = X
    nY = Y - 1
End If

If Head = SOUTH Then
    nx = X
    nY = Y + 1
End If

If Head = EAST Then
    nx = X + 1
    nY = Y
End If

If Head = WEST Then
    nx = X - 1
    nY = Y
End If

'Devuelve valores
Pos.X = nx
Pos.Y = nY
Exit Sub
fallo:
Call LogError("HEADTOPOS" & Err.number & " D: " & Err.Description)

End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua = False) As Boolean
'¿Es un mapa valido?
On Error GoTo fallo


If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            LegalPos = False
Else

  If Not PuedeAgua Then
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).UserIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                    (Not HayAgua(Map, X, Y))
  Else
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).UserIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) 'And _

                    '(HayAgua(Map, x, Y))
  End If


'MsgBox (LegalPos)
End If

Exit Function
fallo:
Call LogError("LEGALPOS" & Err.number & " D: " & Err.Description)


End Function



Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean
On Error GoTo fallo
If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
    LegalPosNPC = False
Else
Dim a As Integer
 a = AguaValida + 1
 If AguaValida = 0 Or AguaValida = 11 Then
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> a) And _
     (MapData(Map, X, Y).UserIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).trigger <> POSINVALIDA) _
     And Not HayAgua(Map, X, Y)
 Else
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).UserIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).trigger <> POSINVALIDA)
 End If
 
End If
Exit Function
fallo:
Call LogError("LEGALPOSNPC" & Err.number & " D: " & Err.Description)


End Function

Sub SendHelp(ByVal index As Integer)
On Error GoTo fallo
Dim NumHelpLines As Integer
Dim loopc As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For loopc = 1 To NumHelpLines
    Call SendData(ToIndex, index, 0, "||" & GetVar(DatPath & "Help.dat", "Help", "Line" & loopc) & "´" & FontTypeNames.FONTTYPE_info)
Next loopc

Exit Sub
fallo:
Call LogError("SENDHELP" & Err.number & " D: " & Err.Description)

End Sub
'pluto:hoy
Public Sub Gusano(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim daño As Integer
Dim lado As Integer
daño = RandomNumber(5, 20)
lado = RandomNumber(35, 36)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & lado & "," & 1)
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 121)
UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
Call SendData(ToIndex, UserIndex, 0, "|| ¡¡ Un Gusano te causa " & daño & " de daño !!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendUserStatsVida(UserIndex)
If UserList(UserIndex).Stats.MinHP <= 0 Then Call UserDie(UserIndex)
Exit Sub
fallo:
Call LogError("GUSANO" & Err.number & " D: " & Err.Description)

End Sub
'pluto:hoy
Public Sub Trampa(ByVal UserIndex As Integer, Tipotrampa As Integer)
On Error GoTo fallo
Dim daño As Integer
daño = RandomNumber(5, 20)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Tipotrampa & "," & 1)
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 120)
UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
Call SendData(ToIndex, UserIndex, 0, "|| ¡¡ Una trampa te causa " & daño & " de daño !!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendUserStatsVida(UserIndex)
If UserList(UserIndex).Stats.MinHP <= 0 Then Call UserDie(UserIndex)
Exit Sub
fallo:
Call LogError("TRAMPA " & Err.number & " D: " & Err.Description)

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo fallo
If Npclist(NpcIndex).NroExpresiones > 0 Then
    Dim randomi
    randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "6°" & Npclist(NpcIndex).Expresiones(randomi) & "°" & Npclist(NpcIndex).Char.CharIndex)
End If
  Exit Sub
fallo:
Call LogError("EXPRESAR " & Err.number & " D: " & Err.Description)
                  
End Sub
Sub MirarDerecho(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo fallo
Dim TempCharIndex As Integer
Dim foundchar As Integer
'¿Posicion valida?
If InMapBounds(Map, X, Y) Then
    '¿Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).UserIndex
            foundchar = 1
        End If
        If MapData(Map, X, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
            foundchar = 2
        End If
    End If
    '¿Es un personaje?
    If foundchar = 0 Then
        If MapData(Map, X, Y).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).UserIndex
            foundchar = 1
        End If
        If MapData(Map, X, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).NpcIndex
            foundchar = 2
        End If
    End If
If foundchar = 1 Then
Dim genero1 As Byte
'pluto:6.0A
    If UserList(TempCharIndex).flags.Privilegios > 0 Then Exit Sub
    Dim UrlClan As String
    If UserList(TempCharIndex).GuildInfo.GuildName = "" Then
        UrlClan = 1
            Else
        Dim TotalClanes As Integer
        Dim NumGuild As Integer
        Dim RevisoGuild As String
        Dim Emblema As String
        TotalClanes = GetVar(App.Path & "\Guilds\" & "GuildsInfo.inf", "Init", "NroGuilds")
        For NumGuild = 1 To TotalClanes
        RevisoGuild = GetVar(App.Path & "\Guilds\" & "GuildsInfo.inf", "GUILD" & NumGuild, "GuildName")
        If RevisoGuild = UserList(TempCharIndex).GuildInfo.GuildName Then
        Exit For
        End If
        Next
        Dim oGuild As cGuild
        Set oGuild = FetchGuild(UserList(TempCharIndex).GuildInfo.GuildName)
        UrlClan = oGuild.Emblema
        If UrlClan = "" Then UrlClan = 1
    End If
    If UCase$(UserList(TempCharIndex).Genero) = "HOMBRE" Then genero1 = 1 Else genero1 = 2
    Call SendData(ToIndex, UserIndex, 0, "K1" & UserList(TempCharIndex).Name & "," & UserList(TempCharIndex).Hogar & "," & UserList(TempCharIndex).clase & "," & UserList(TempCharIndex).raza & "," & UserList(TempCharIndex).Remort & "," & genero1 & "," & UserList(TempCharIndex).Nhijos & "," & UserList(TempCharIndex).Hijo(1) & "," & UserList(TempCharIndex).Hijo(2) & "," & UserList(TempCharIndex).Hijo(3) & "," & UserList(TempCharIndex).Hijo(4) & "," & UserList(TempCharIndex).Hijo(5) & "," & UserList(TempCharIndex).Padre & "," & UserList(TempCharIndex).Madre & "," & UserList(TempCharIndex).Esposa & "," & UserList(TempCharIndex).Amor & "," & UserList(TempCharIndex).Embarazada & ";" & UrlClan)
    End If
End If

Exit Sub
fallo:
Call LogError("MirarDerecho" & Err.number & " D: " & Err.Description)
End Sub


Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

On Error GoTo fallo
'Responde al click del usuario sobre el mapa
Dim foundchar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim clickpiso As WorldPos
'¿Posicion valida?
If InMapBounds(Map, X, Y) Then
    UserList(UserIndex).flags.TargetMap = Map
    UserList(UserIndex).flags.TargetX = X
    UserList(UserIndex).flags.TargetY = Y

    '¿Es un obj?
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre

'pluto:hoy
If UserList(UserIndex).Pos.Map = 180 And ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = 57 Then
If MapData(Map, X, Y).OBJInfo.ObjIndex = ResEgipto Then
       Call WarpUserChar(UserIndex, 182, 47, 50, True)
       'pluto:2-3-04
        Call SendData(ToMap, 0, 0, "TW" & 138)
        Call LoadEgipto
        Else:  Call WarpUserChar(UserIndex, 181, 42, 66, True)
End If
End If

'pluto:2-3-04
If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = 58 Then
If UltimoBan <> "" Then Call SendData(ToIndex, UserIndex, 0, "||Este es el destino al que se vió reducido " & UltimoBan & " por las grandes fechorías que cometió, siendo aquí colgado en público para su verguenza." & "´" & FontTypeNames.FONTTYPE_FIGHT)
If UltimoBan = "" Then Call SendData(ToIndex, UserIndex, 0, "||En estos momentos no hay ningún delincuente ahorcado." & "´" & FontTypeNames.FONTTYPE_FIGHT)
End If

'pluto:2.4
Dim Tipo As Integer
Tipo = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType
If Tipo = 44 Or Tipo = 45 Then
Call SendData2(ToIndex, UserIndex, 0, 77, Tipo)
End If

'pluto:2-3-04 momia faraón
If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = 59 Then
    If Tesoromomia = 0 Then
        Call SpawnNpc(611, UserList(UserIndex).Pos, True, False)
        Call SendData(ToAll, 0, 0, "||¡¡¡ Han invocado a la grandisima Momia Faraón " & " !!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
        Tesoromomia = 1
        
    Else
        Call SendData(ToIndex, UserIndex, 0, "|| ¡¡ Se han llevado el tesoro !!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
End If
'Caballero de la Muerte
If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = 70 Then
    If tesorocaballero = 0 Then
        Call SpawnNpc(726, UserList(UserIndex).Pos, True, False)
        Call SendData(ToAll, 0, 0, "||¡¡¡ Han invocado al Gran Caballero De la Muerte " & " !!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
        tesorocaballero = 1
    Else
        Call SendData(ToIndex, UserIndex, 0, "|| ¡¡ El Caballero ya ha sido desterrado !!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
    End If
End If

Dim ab As String
'pluto:2.10
If UserList(UserIndex).flags.Privilegios > 1 Then
Call SendData(ToIndex, UserIndex, 0, "||Objeto Numero: " & MapData(Map, X, Y).OBJInfo.ObjIndex & "´" & FontTypeNames.FONTTYPE_info)
End If

If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Peso > 0 Then ab = " " & (ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Peso * MapData(Map, X, Y).OBJInfo.Amount) & " Kg" Else ab = ""
If (MapData(Map, X, Y).OBJInfo.Amount > 1 And ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType <> 4) Then
Call SendData(ToIndex, UserIndex, 0, "||" & MapData(Map, X, Y).OBJInfo.Amount & " " & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Name & ab & "´" & FontTypeNames.FONTTYPE_info)
Else
Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Name & ab & "´" & FontTypeNames.FONTTYPE_info)
End If

'pluto:6.2----------------------
    If MapData(Map, X, Y).UserIndex > 0 And UserList(UserIndex).flags.Muerto = 0 Then
    If UserList(MapData(Map, X, Y).UserIndex).flags.Muerto = 1 And MapData(Map, X, Y).UserIndex <> UserIndex Then
clickpiso.Map = Map
clickpiso.X = X
clickpiso.Y = Y
    If Distancia(UserList(UserIndex).Pos, clickpiso) < 2 Then
    Call GetObjFantasma(UserIndex, X, Y)
    End If
    End If
    End If
'----------------------------------


        ' Then Call SendData(ToIndex, UserIndex, 0, "||" &  & " Kg" & FONTTYPENAMES.FONTTYPE_INFO)
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
        UserList(UserIndex).flags.TargetObjMap = Map
        UserList(UserIndex).flags.TargetObjX = X
        UserList(UserIndex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType = OBJTYPE_PUERTAS Then
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType = OBJTYPE_PUERTAS Then
            'Informa el nombre
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType = OBJTYPE_PUERTAS Then
            'Informa el nombre
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    'pluto:2.15 yoyita
    If UserList(UserIndex).flags.TargetObj > 0 Then
    If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = OBJTYPE_PUERTAS Or ObjData(UserList(UserIndex).flags.TargetObj).OBJType = OBJTYPE_CARTELES Or ObjData(UserList(UserIndex).flags.TargetObj).OBJType = OBJTYPE_FOROS Or ObjData(UserList(UserIndex).flags.TargetObj).OBJType = OBJTYPE_LEÑA Then
    Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    End If
    End If
    '-------
    
    '¿Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).UserIndex
            foundchar = 1
        End If
        If MapData(Map, X, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
            foundchar = 2
        End If
    End If
    '¿Es un personaje?
    If foundchar = 0 Then
        If MapData(Map, X, Y).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).UserIndex
            foundchar = 1
        End If
        If MapData(Map, X, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).NpcIndex
            foundchar = 2
        End If
    End If
    
    
    'Reaccion al personaje
    If foundchar = 1 Then '  ¿Encontro un Usuario?
            
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Then

            If EsNewbie(TempCharIndex) Then
                Stat = " <NEWBIE>"
            End If
            
            If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                Stat = Stat & " <Ejercito real> " & "<" & TituloReal(TempCharIndex) & ">"
            ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                Stat = Stat & " <Fuerzas del caos> " & "<" & TituloCaos(TempCharIndex) & ">"
            'legion
            ElseIf UserList(TempCharIndex).Faccion.ArmadaReal = 2 Then
                Stat = Stat & " <La Legión> " & "<" & Titulolegion(TempCharIndex) & ">"
                End If
            
            If UserList(TempCharIndex).GuildInfo.GuildName <> "" Then
                'pluto:2.4
                'Nati: Ahora sera de un titulo segun sus puntos aportados al clan.
              Dim a As String
            a = " (Soldado)"
             If UserList(TempCharIndex).Stats.PClan >= 100 Then a = " (Teniente)"
             If UserList(TempCharIndex).Stats.PClan >= 250 Then a = " (Capitán)"
             If UserList(TempCharIndex).Stats.PClan >= 500 Then a = " (General)"
             If UserList(TempCharIndex).Stats.PClan >= 1000 Then a = " (Comandante)"
             If UserList(TempCharIndex).Stats.PClan >= 1500 Then a = " (SubLider)"
             If UserList(TempCharIndex).GuildInfo.GuildPoints >= 5000 Then a = " (Lider)"
            'If UserList(TempCharIndex).GuildInfo.GuildPoints >= 1000 Then a = " (Teniente)"
            'If UserList(TempCharIndex).GuildInfo.GuildPoints >= 2000 Then a = " (Capitán)"
            'If UserList(TempCharIndex).GuildInfo.GuildPoints >= 3000 Then a = " (General)"
            'If UserList(TempCharIndex).GuildInfo.GuildPoints >= 4000 Then a = " (SubLider)"
            'If UserList(TempCharIndex).GuildInfo.GuildPoints >= 5000 Then a = " (Lider)"

               Stat = Stat & " <" & UserList(TempCharIndex).GuildInfo.GuildName & a & ">"
              '-----------------fin pluto:2.4-------------------
              
              End If
            
            If Len(UserList(TempCharIndex).Desc) > 1 Then
                Stat = "||Ves a " & UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).Desc
            Else
                'Call SendData(ToIndex, UserIndex, 0, "||Ves a " & UserList(TempCharIndex).Name & Stat)
                Stat = "||Ves a " & UserList(TempCharIndex).Name & Stat
            End If
            
            If UserList(TempCharIndex).Remort = 1 Then Stat = Stat & " *" & UserList(TempCharIndex).Remorted & "*"
            'LEGION

            If Criminal(TempCharIndex) And UserList(TempCharIndex).flags.Privilegios = 0 Then
                'Stat = Stat & " <CRIMINAL> ~255~0~0~1~0"
                Stat = Stat & "´" & FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
            ElseIf UserList(TempCharIndex).Faccion.ArmadaReal = 2 And UserList(TempCharIndex).flags.Privilegios = 0 Then
                'Stat = Stat & " <LEGIONARIO> ~0~255~0~1~0 "
                Stat = Stat & "´" & FontTypeNames.FONTTYPE_CONSEJOVesa
            ElseIf UserList(TempCharIndex).flags.Privilegios = 0 Then
                'Stat = Stat & " <CIUDADANO>~0~0~200~1~0"
                Stat = Stat & "´" & FontTypeNames.FONTTYPE_CONSEJO
            ElseIf UserList(TempCharIndex).flags.Privilegios > 0 Then
                'Stat = Stat & " <GameMaster>~255~255~255~1~0"
                Stat = Stat & "´" & FontTypeNames.FONTTYPE_TALK
            End If
            
            Call SendData(ToIndex, UserIndex, 0, Stat)
            
            FoundSomething = 1
            UserList(UserIndex).flags.TargetUser = TempCharIndex
            UserList(UserIndex).flags.TargetNpc = 0
            UserList(UserIndex).flags.TargetNpcTipo = 0
            'nati: hago que me envie el nombre del usuario
            Call SendData2(ToIndex, UserIndex, 0, 115, UserList(TempCharIndex).Name)
            'nati: hago que me envie el nombre del usuario
       
       End If
       
    End If
    If foundchar = 2 Then '¿Encontro un NPC?

'pluto:6.4
 UserList(UserIndex).flags.TargetUser = 0

 
 'pluto:2.15
If Distancia(UserList(UserIndex).Pos, Npclist(TempCharIndex).Pos) > 10 Then
Call SendData(ToIndex, UserIndex, 0, "L2")
GoTo AI
End If
  
  '¿Esta el user muerto? Si es asi no puede interactuar
        If UserList(UserIndex).flags.Muerto = 1 And Npclist(TempCharIndex).NPCtype <> 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                 Exit Sub
        End If

'-------------------------------------------------
'-------------------------------------------------
'pluto:6.0A pongo selec case a los tipos de npcs
'-------------------------------------------------
'--------------------------------------------------
UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
UserList(UserIndex).flags.TargetNpc = TempCharIndex
'pluto:7.0
If UserList(UserIndex).flags.Privilegios > 0 Then
Call SendData(ToIndex, UserIndex, 0, "|| Número Npc: " & Npclist(TempCharIndex).numero & "´" & FontTypeNames.FONTTYPE_info)
End If

'pluto:6.0A
If UserList(UserIndex).flags.Navegando = 1 And (Npclist(TempCharIndex).Comercia > 0 Or Npclist(TempCharIndex).NPCtype = 1 Or Npclist(TempCharIndex).NPCtype = 4) Then
Call SendData(ToIndex, UserIndex, 0, "||¡¡Deja de Navegar!!" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If


'PLUTO:7.0
'entrega misiones
If UserList(UserIndex).Mision.estado = 1 And (Npclist(TempCharIndex).numero = UserList(UserIndex).Mision.Entrega) Then 'And TieneObjetos(UserList(UserIndex).Mision.Objeto, UserList(UserIndex).Mision.Cantidad, UserIndex) Then
   'cargar
   If UserList(UserIndex).Mision.Cargada = False Then
   Call CargarQuest(UserIndex)
   End If
   
   If ComprobarObjetivos(UserIndex) = True Then
   
   
   'quitamos objetos
   Dim n As Byte
   Dim Ocan As Integer
   Dim Ocan2 As Byte
   
   For n = 1 To UserList(UserIndex).Mision.NObjetos
   Ocan = val(ReadField(1, UserList(UserIndex).Mision.Objeto(n), 45))
   Ocan2 = val(ReadField(2, UserList(UserIndex).Mision.Objeto(n), 45))
   Call SendData(ToIndex, UserIndex, 0, "|| Has pérdido " & Ocan2 & " " & ObjData(Ocan).Name & "´" & FontTypeNames.FONTTYPE_info)
   Call QuitarObjetos(Ocan, Ocan2, UserIndex)
   Next
   'entregamos recompensa de objetos
    
    Dim MiObj As obj
    For n = 1 To UserList(UserIndex).Mision.NObjetosR
    'nati: agrego el "(ALTOS)NºITEM-CANTIDAD/(ENANOS)NºITEM-CANTIDAD
    'nati:            Miobj.ObjIndex-MiObj.amount& Miobj.Separador & Miobj.ObjIndex1-MiObj.amount1
    MiObj.ObjIndex = val(ReadField(1, UserList(UserIndex).Mision.ObjetoR(n), 45))
    MiObj.Amount = val(ReadField(2, UserList(UserIndex).Mision.ObjetoR(n), 45))
    MiObj.Separador = val(ReadField(3, UserList(UserIndex).Mision.ObjetoR(n), 47))
    MiObj.ObjIndex2 = val(ReadField(2, UserList(UserIndex).Mision.ObjetoR(n), 47))
    MiObj.Amount2 = val(ReadField(3, UserList(UserIndex).Mision.ObjetoR(n), 45))
    'Si en el documento QUEST no hay separador, entregamos Miobj normal.
    If MiObj.ObjIndex2 = "0" And MiObj.Amount2 = "0" Then
    Call SendData(ToIndex, UserIndex, 0, "|| Has obtenido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & "´" & FontTypeNames.FONTTYPE_info)

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
    End If
    'Si es ENANO O GNOMO entregamos Miobj.2
    If UserList(UserIndex).raza = "Enano" Or UserList(UserIndex).raza = "Gnomo" Or UserList(UserIndex).raza = "Goblin" Then
    Call SendData(ToIndex, UserIndex, 0, "|| Has obtenido " & MiObj.Amount2 & " " & ObjData(MiObj.ObjIndex2).Name & "´" & FontTypeNames.FONTTYPE_info)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
    Else
    Call SendData(ToIndex, UserIndex, 0, "|| Has obtenido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
    End If
    Next n
    'entregamos oro
    If UserList(UserIndex).Mision.oro > 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.GLD, UserList(UserIndex).Mision.oro, MAXORO)
    Call SendData(ToIndex, UserIndex, 0, "|| Has obtenido " & UserList(UserIndex).Mision.oro & " Monedas de Oro." & "´" & FontTypeNames.FONTTYPE_info)
    Call SendUserStatsOro(UserIndex)
    End If
    'entregamos exp
    If UserList(UserIndex).Mision.exp > 0 Then
    UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + Int(UserList(UserIndex).Mision.exp)
    Call SendData(ToIndex, UserIndex, 0, "|| Has obtenido " & UserList(UserIndex).Mision.exp & " Puntos de Experiencia." & "´" & FontTypeNames.FONTTYPE_info)
    SendUserStatsEXP (UserIndex)
    CheckUserLevel (UserIndex)
    End If
    'reiniciamos datos mision
  UserList(UserIndex).Mision.numero = UserList(UserIndex).Mision.numero + 1
   ResetMisionCompletada (UserIndex)
    Exit Sub
    Else
    Call SendData(ToIndex, UserIndex, 0, "|| Tienes Objetivos no cumplidos." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
End If
   
   
          ' UserList(UserIndex).Mision.estado = 0
          ' Call SendData(ToIndex, UserIndex, 0, "!!Quest Número " & UserList(UserIndex).Mision.Numero & " : " & " Muy bién, has cumplido una misión!!")
           'Call QuitarObjetos(UserList(UserIndex).Mision.Objeto, UserList(UserIndex).Mision.Cantidad, UserIndex)
           'pluto:6.0A
           'UserList(UserIndex).Stats.Fama = UserList(UserIndex).Stats.Fama + 5
           'pluto:2-3-04
           'Call SendData(ToIndex, UserIndex, 0, "|| Has ganado " & val(Int(UserList(UserIndex).Mision.Numero / 10) + 1) & " DragPuntos." & "´" & FontTypeNames.FONTTYPE_info)
'UserList(UserIndex).Stats.Puntos = UserList(UserIndex).Stats.Puntos + Int(UserList(UserIndex).Mision.Numero / 10) + 1
  'pluto:2.19----------------------------------------------
'Call SendData(ToIndex, UserIndex, 0, "|| Has ganado " & val(Int(UserList(UserIndex).Mision.Numero * 500)) & " Puntos de Experiencia." & "´" & FontTypeNames.FONTTYPE_info)
'UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + Int(UserList(UserIndex).Mision.Numero * 500)
'SendUserStatsEXP (UserIndex)
'CheckUserLevel (UserIndex)
'----------------------------------------------


End If


Select Case Npclist(TempCharIndex).NPCtype

Case 43
Dim ViajeUlla As String
Dim ViajeCaos As String
Dim ViajeDescanso As String
Dim ViajeAtlantis As String
Dim ViajeArghal As String
Dim ViajeEsperanza As String
Dim ViajeNix As String
Dim ViajeRinkel As String
Dim ViajeBander As String
Dim ViajeLindos As String
'¿Esta en Nix?
If UserList(UserIndex).Pos.Map = 34 Then
ViajeUlla = 3000
ViajeCaos = 4000
ViajeDescanso = 6000
ViajeAtlantis = 6500
Call SendData2(ToIndex, UserIndex, 0, 116, "ULLA@" & ViajeUlla)
Call SendData2(ToIndex, UserIndex, 0, 116, "CAOS@" & ViajeCaos)
Call SendData2(ToIndex, UserIndex, 0, 116, "DESCANSO@" & ViajeDescanso)
Call SendData2(ToIndex, UserIndex, 0, 116, "ATLANTIS@" & ViajeAtlantis)
End If
'¿Esta en Ulla?
If UserList(UserIndex).Pos.Map = 1 Then
ViajeNix = 3000
ViajeCaos = 2500
ViajeDescanso = 3500
ViajeBander = 5500
ViajeRinkel = 3500
Call SendData2(ToIndex, UserIndex, 0, 116, "NIX@" & ViajeNix)
Call SendData2(ToIndex, UserIndex, 0, 116, "CAOS@" & ViajeCaos)
Call SendData2(ToIndex, UserIndex, 0, 116, "DESCANSO@" & ViajeDescanso)
Call SendData2(ToIndex, UserIndex, 0, 116, "BANDER@" & ViajeBander)
Call SendData2(ToIndex, UserIndex, 0, 116, "RINKEL@" & ViajeRinkel)
End If
'¿Esta en Descanso?
If UserList(UserIndex).Pos.Map = 81 Then
ViajeUlla = 3500
ViajeNix = 6000
ViajeCaos = 5500
ViajeArghal = 12000
ViajeBander = 3500
Call SendData2(ToIndex, UserIndex, 0, 116, "ULLA@" & ViajeUlla)
Call SendData2(ToIndex, UserIndex, 0, 116, "BANDER@" & ViajeBander)
Call SendData2(ToIndex, UserIndex, 0, 116, "NIX@" & ViajeNix)
Call SendData2(ToIndex, UserIndex, 0, 116, "CAOS@" & ViajeCaos)
Call SendData2(ToIndex, UserIndex, 0, 116, "ARGHAL@" & ViajeArghal)
End If
'¿Esta en Bander?
If UserList(UserIndex).Pos.Map = 59 Then
ViajeUlla = 5500
ViajeDescanso = 3500
ViajeAtlantis = 10000
ViajeArghal = 12000
Call SendData2(ToIndex, UserIndex, 0, 116, "ULLA@" & ViajeUlla)
Call SendData2(ToIndex, UserIndex, 0, 116, "DESCANSO@" & ViajeDescanso)
Call SendData2(ToIndex, UserIndex, 0, 116, "ATLANTIS@" & ViajeAtlantis)
Call SendData2(ToIndex, UserIndex, 0, 116, "ARGHAL@" & ViajeArghal)
End If
'¿Esta en Rinkel?
If UserList(UserIndex).Pos.Map = 20 Then
ViajeUlla = 3500
ViajeLindos = 5500
ViajeAtlantis = 9500
ViajeEsperanza = 12000
Call SendData2(ToIndex, UserIndex, 0, 116, "ULLA@" & ViajeUlla)
Call SendData2(ToIndex, UserIndex, 0, 116, "LINDOS@" & ViajeLindos)
Call SendData2(ToIndex, UserIndex, 0, 116, "ATLANTIS@" & ViajeAtlantis)
Call SendData2(ToIndex, UserIndex, 0, 116, "ESPERANZA@" & ViajeEsperanza)
End If
'¿Esta en Caos?
If UserList(UserIndex).Pos.Map = 170 Then
ViajeNix = 4500
ViajeUlla = 2500
ViajeLindos = 6500
ViajeDescanso = 5500
Call SendData2(ToIndex, UserIndex, 0, 116, "NIX@" & ViajeNix)
Call SendData2(ToIndex, UserIndex, 0, 116, "ULLA@" & ViajeUlla)
Call SendData2(ToIndex, UserIndex, 0, 116, "LINDOS@" & ViajeLindos)
Call SendData2(ToIndex, UserIndex, 0, 116, "DESCANSO@" & ViajeDescanso)
End If
'¿Esta en Arghal?
If UserList(UserIndex).Pos.Map = 151 Then
ViajeDescanso = 12000
ViajeBander = 12000
Call SendData2(ToIndex, UserIndex, 0, 116, "DESCANSO@" & ViajeDescanso)
Call SendData2(ToIndex, UserIndex, 0, 116, "BANDER@" & ViajeBander)
End If
'¿Esta en Atlantis?
If UserList(UserIndex).Pos.Map = 85 Then
ViajeNix = 6500
ViajeBander = 10000
ViajeRinkel = 9500
Call SendData2(ToIndex, UserIndex, 0, 116, "NIX@" & ViajeNix)
Call SendData2(ToIndex, UserIndex, 0, 116, "BANDER@" & ViajeBander)
Call SendData2(ToIndex, UserIndex, 0, 116, "RINKEL@" & ViajeRinkel)
End If
'¿Esta en Lindos?
If UserList(UserIndex).Pos.Map = 63 Then
ViajeCaos = 3500
ViajeEsperanza = 7500
ViajeRinkel = 5500
Call SendData2(ToIndex, UserIndex, 0, 116, "CAOS@" & ViajeCaos)
Call SendData2(ToIndex, UserIndex, 0, 116, "ESPERANZA@" & ViajeEsperanza)
Call SendData2(ToIndex, UserIndex, 0, 116, "RINKEL@" & ViajeRinkel)
End If
'¿Esta en Isla Esperanza?
If UserList(UserIndex).Pos.Map = 111 Then
ViajeLindos = 7500
ViajeRinkel = 12500
Call SendData2(ToIndex, UserIndex, 0, 116, "LINDOS@" & ViajeLindos)
Call SendData2(ToIndex, UserIndex, 0, 116, "RINKEL@" & ViajeRinkel)
End If

'pluto:7.0
Case 62
Call SendData2(ToIndex, UserIndex, 0, 111, UserList(UserIndex).flags.Creditos)

Case 14
If UserList(UserIndex).Mision.estado = 1 Then
    'cargamos datos quest
    If UserList(UserIndex).Mision.Cargada = False Then
    Call CargarQuest(UserIndex)
    UserList(UserIndex).Mision.Cargada = True
    End If

'es otro npcquest
 If UserList(UserIndex).Mision.NpcQuest <> UserList(UserIndex).Mision.Entrega Then
 Call SendData(ToPCArea, UserIndex, Npclist(TempCharIndex).Pos.Map, "||8° Ya tienes una misión activa." & "°" & Npclist(TempCharIndex).Char.CharIndex)
 End If
'es el npc de entrega, pasamos a comprobar objetivos
'Call ComprobarObjetivos(UserIndex)
 
Else 'estado=0
Call iniciarquest(UserIndex)

End If 'estado=1
Exit Sub


Case 1
 'resucitar
      If UserList(UserIndex).flags.Muerto = 0 Then
      Call SendData(ToIndex, UserIndex, 0, "TW" & 181)
      Exit Sub
      End If
  'pluto:2.18
 'If (MapInfo(Npclist(TempCharIndex).Pos.Map).Dueño = 1 And Criminal(UserIndex)) Or (MapInfo(Npclist(TempCharIndex).Pos.Map).Dueño = 2 And Not Criminal(UserIndex)) Then
 'Call SendData(ToIndex, UserIndex, 0, "||6°" & "No puedo resucitarte, tu armada no controla esta ciudad." & "°" & Npclist(TempCharIndex).Char.CharIndex)
'Exit Sub
'End If
  '--------
    'pluto:6.0A----------
    If UserList(UserIndex).flags.Navegando > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Deja de Navegar!!" & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
    'pluto:6.9
    If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
    Call SendData(ToIndex, UserIndex, 0, "L2")
    Exit Sub
    End If
     '-------------------
    Call RevivirUsuario(UserIndex)
    Call SendData(ToIndex, UserIndex, 0, "S3")
    'Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido resucitado!!" & FONTTYPENAMES.FONTTYPE_INFO)
    Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 72 & "," & 1)
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    Call SendUserStatsVida(val(UserIndex))
    'Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido curado!!" & FONTTYPENAMES.FONTTYPE_INFO)
Exit Sub

Case 36
If UserList(UserIndex).flags.Minotauro > 0 Then
Call SendData(ToPCArea, UserIndex, Npclist(TempCharIndex).Pos.Map, "||8° No puedes liberar el Minotauro más veces." & "°" & Npclist(TempCharIndex).Char.CharIndex)
Exit Sub
End If

If EstadoMinotauro = 2 Then
Call SendData(ToPCArea, UserIndex, Npclist(TempCharIndex).Pos.Map, "||8° El Minotauro fué liberado hace poco y debes esperar un tiempo para poder liberarlo de nuevo." & "°" & Npclist(TempCharIndex).Char.CharIndex)
Exit Sub
End If


If Not TieneObjetos(1218, 1, UserIndex) Then
Call SendData(ToPCArea, UserIndex, Npclist(TempCharIndex).Pos.Map, "||8° Necesitas un Hilo de Ariadna que poseen algunas Viudas Negras para liberar el Minotauro." & "°" & Npclist(TempCharIndex).Char.CharIndex)
Exit Sub
Else
Call QuitarObjetos(1218, 1, UserIndex)
End If
    
  
'pluto 6.0a LIBERA
If Minotauro = "" Then
Call SendData(ToPCArea, UserIndex, Npclist(TempCharIndex).Pos.Map, "||8° Aprisa!, El Minotauro ha huido al verte, búscalo y mátalo antes de que lo haga cualquier otro..." & "°" & Npclist(TempCharIndex).Char.CharIndex)
Minotauro = UserList(UserIndex).Name
Dim mapita As Integer
Dim CabalgaPos As WorldPos
Dim ini As Integer
a:
    mapita = RandomNumber(1, 277)
    CabalgaPos.X = RandomNumber(15, 80)
    CabalgaPos.Y = RandomNumber(15, 80)
    CabalgaPos.Map = mapita
    If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a
    ini = SpawnNpc(692, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a:
    'CabalgaPos.Map = RandomNumber(1, 277)
    'mapita = CabalgaPos.Map
    'ini = SpawnNpc(692, CabalgaPos, False, True)
    'If ini = MAXNPCS Then mapita = 1000
   
Call SendData(ToAll, UserIndex, 0, "|| El Minotauro ha sido liberado por " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_PARTY)
 EstadoMinotauro = 1
 MinutosMinotauro = 30

    'Call WriteVar(IniPath & "cabalgar.txt", MiNPC.Name, "Mapa", val(Mapita))

Call LogCasino("Minotauro: " & CabalgaPos.Map & "-" & CabalgaPos.X & "-" & CabalgaPos.Y & " liberado por " & UserList(UserIndex).Name)
 Else 'NO LIBERA
 
    If Minotauro <> UserList(UserIndex).Name Then
    Call SendData(ToPCArea, UserIndex, Npclist(TempCharIndex).Pos.Map, "||8° El Minotauro fué liberado por otro personaje, debes esperar que sea capturado." & "°" & Npclist(TempCharIndex).Char.CharIndex)
    Else
    Call SendData(ToPCArea, UserIndex, Npclist(TempCharIndex).Pos.Map, "||8° Aprisa!, El Minotauro ha huido al verte, búscalo y mátalo antes de que lo haga cualquier otro..." & "°" & Npclist(TempCharIndex).Char.CharIndex)
    End If

End If
Exit Sub

Case 37
'pluto:6.0A
Call SendData(ToIndex, UserIndex, 0, "H6")
Exit Sub

Case 4
'pluto:6.0A
Call SendData(ToIndex, UserIndex, 0, "H1" & "," & UserList(UserIndex).Stats.Banco)
Exit Sub

Case 26
 'pluto:2.22-------------------------------------------
If MapData(277, 36, 70).OBJInfo.ObjIndex > 0 And UserList(UserIndex).Pos.Map = 277 And UserList(UserIndex).Pos.X = 36 And UserList(UserIndex).Pos.Y = 70 Then
Dim nuo As Integer
Dim nuoc As Integer
nuo = MapData(277, 36, 70).OBJInfo.ObjIndex
nuoc = MapData(277, 36, 70).OBJInfo.Amount

If (ObjData(nuo).LingH = 0 And ObjData(nuo).LingP = 0 And ObjData(nuo).LingO = 0) Then
Call SendData(ToIndex, UserIndex, 0, "|| Ese Objeto no lo puedo fundir!!" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Exit Sub
End If

'pluto:6.0A
If ObjData(nuo).LingH * nuoc > 10000 Or ObjData(nuo).LingP * nuoc > 10000 Or ObjData(nuo).LingO * nuoc > 10000 Then
Call SendData(ToIndex, UserIndex, 0, "|| No puedo fundir tantos objetos, por favor suelta menos objetos." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

Dim Esvende As Byte
If ObjData(nuo).Vendible = 0 Then Esvende = 5 Else Esvende = 60
'borramos el objeto del suelo
Call EraseObj(ToMap, 0, 277, 10000, 277, 36, 70)
'pluto:6.0A
Call LogNpcFundidor(UserList(UserIndex).Name & " Funde Obj: " & nuo & " Cant: " & nuoc)

'Dim MiObj As obj

    ' lingotes hierro
    If ObjData(nuo).LingH > 0 Then
MiObj.ObjIndex = 386
MiObj.Amount = Porcentaje(ObjData(nuo).LingH * nuoc, Esvende)
    If MiObj.Amount < 1 Then GoTo P1
Call SendData(ToIndex, UserIndex, 0, "|| Has ganado " & MiObj.Amount & " Lingotes de Hierro" & "´" & FontTypeNames.FONTTYPE_info)
'pluto:6.0A
Call LogNpcFundidor(UserList(UserIndex).Name & " Obtiene LingH : " & MiObj.Amount)

If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
    End If 'hierro
P1:
    ' lingotes plata
    If ObjData(nuo).LingP > 0 Then
MiObj.ObjIndex = 387
MiObj.Amount = Porcentaje(ObjData(nuo).LingP * nuoc, Esvende)
    If MiObj.Amount < 1 Then GoTo P2

Call SendData(ToIndex, UserIndex, 0, "|| Has ganado " & MiObj.Amount & " Lingotes de Plata" & "´" & FontTypeNames.FONTTYPE_info)
'pluto:6.0A
Call LogNpcFundidor(UserList(UserIndex).Name & " Obtiene LingP : " & MiObj.Amount)
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
    End If 'plata
P2:
       ' lingotes oro
If ObjData(nuo).LingO > 0 Then
MiObj.ObjIndex = 388
MiObj.Amount = Porcentaje(ObjData(nuo).LingO * nuoc, Esvende)
    If MiObj.Amount < 1 Then GoTo P3

Call SendData(ToIndex, UserIndex, 0, "|| Has ganado " & MiObj.Amount & " Lingotes de Oro" & "´" & FontTypeNames.FONTTYPE_info)
'pluto:6.0A
Call LogNpcFundidor(UserList(UserIndex).Name & " Obtiene LingO : " & MiObj.Amount)

If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
    End If 'oro
P3:
Call SendData(ToIndex, UserIndex, 0, "||5°" & "He fundido el objeto!!." & "°" & Npclist(TempCharIndex).Char.CharIndex)
End If


Case 31
 'pluto.6.0A-------------
Dim nx As Byte
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then
Call SendData(ToIndex, UserIndex, 0, "||5°" & "No eres Lider de Clan!!." & "°" & Npclist(TempCharIndex).Char.CharIndex)
Exit Sub
End If
'nx = UserList(userindex).GuildRef.Nivel
Call SendData(ToIndex, UserIndex, 0, "||5°" & "Para subir tu clan al Nivel ." & UserList(UserIndex).GuildRef.Nivel + 1 & " escribe /NIVELCLAN." & "°" & Npclist(TempCharIndex).Char.CharIndex)
Exit Sub
 '---------------------
  'pluto.6.0A----------------------------------------------------
 Case 32
 'pluto:6.9
 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 And UserList(UserIndex).flags.Morph = 0 And UserList(UserIndex).flags.Angel = 0 And UserList(UserIndex).flags.Demonio = 0 Then
     Dim Arm As ObjData
     Dim Slot As Byte
     Arm = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
     Slot = UserList(UserIndex).Invent.WeaponEqpSlot
  'comprobamos si tiene la piedra onice
    If TieneObjetos(1170, 1, UserIndex) And Arm.ArmaNpc > 0 Then
    Call QuitarUserInvItem(UserIndex, Slot, 1)
    Call UpdateUserInv(False, UserIndex, Slot)

    Call QuitarObjetos(1170, 1, UserIndex)
   
    MiObj.Amount = 1
    MiObj.ObjIndex = Arm.ArmaNpc
     
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 109)
        Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 128 & "," & 1)

        Call SendData(ToIndex, UserIndex, 0, "||5°" & "Perfecto!! Ha sido un placer hacer negocios contigo." & "°" & Npclist(TempCharIndex).Char.CharIndex)
      
         Exit Sub
    End If

 Else
 Call SendData(ToIndex, UserIndex, 0, "||5°" & "Hola Aodraguero, para mejorarte un Arma necesito que me traigas una Piedra Mágica, que tengas el arma equipada y que no estes transformado." & "°" & Npclist(TempCharIndex).Char.CharIndex)
 Exit Sub
 End If
 Exit Sub
 
'pluto:6.0A
Case 39
If TieneObjetos(414, 100, UserIndex) Then
Call QuitarObjetos(414, 100, UserIndex)
Call AddtoVar(UserList(UserIndex).Stats.GLD, 10000, MAXORO)
Call SendData(ToIndex, UserIndex, 0, "||5°" & "Perfecto!! Ha sido un placer hacer negocios contigo." & "°" & Npclist(TempCharIndex).Char.CharIndex)
Call SendData(ToIndex, UserIndex, 0, "||Has ganado 10000 Oros!!" & "´" & FontTypeNames.FONTTYPE_info)
Call SendUserStatsOro(UserIndex)
Else
Call SendData(ToIndex, UserIndex, 0, "||5°" & "Vuelve cuando tengas 100 Pieles de Lobo o 100 Botas Rotas y te recompensaré con 10000 oros." & "°" & Npclist(TempCharIndex).Char.CharIndex)
End If

If TieneObjetos(887, 100, UserIndex) Then
Call QuitarObjetos(887, 100, UserIndex)
Call AddtoVar(UserList(UserIndex).Stats.GLD, 10000, MAXORO)
Call SendData(ToIndex, UserIndex, 0, "||5°" & "Perfecto!! Ha sido un placer hacer negocios contigo." & "°" & Npclist(TempCharIndex).Char.CharIndex)
Call SendData(ToIndex, UserIndex, 0, "||Has ganado 10000 Oros!!" & "´" & FontTypeNames.FONTTYPE_info)
Call SendUserStatsOro(UserIndex)
Else
Call SendData(ToIndex, UserIndex, 0, "||5°" & "Vuelve cuando tengas 100 Pieles de Lobo o 100 Botas Rotas y te recompensaré con 10000 oros." & "°" & Npclist(TempCharIndex).Char.CharIndex)
End If
Exit Sub

'pluto:6.5
Case 40 'torneo parejas
If MapInfo(291).NumUsers < 3 Then
'comprueba situación pareja
    If MapData(296, 67, 44).UserIndex > 0 And MapData(296, 65, 44).UserIndex > 0 Then
Dim Pareja1 As Integer
Dim Pareja2 As Integer
Dim r10
Dim y10
r10 = RandomNumber(52, 71)
y10 = RandomNumber(44, 59)
Pareja1 = MapData(296, 67, 44).UserIndex
Pareja2 = MapData(296, 65, 44).UserIndex
Call WarpUserChar(Pareja1, 291, r10, y10, True)
Call WarpUserChar(Pareja2, 291, r10 + 1, y10, True)
UserList(Pareja2).flags.ParejaTorneo = Pareja1
UserList(Pareja1).flags.ParejaTorneo = Pareja2
'pluto:6.3---
Call SendData(ToMap, UserIndex, 0, "La Pareja formada por " & Pareja1 & " y " & Pareja2 & " ha entrado a la sala de Torneo Parejas" & "´" & FontTypeNames.FONTTYPE_TALK)
'-------------
    Else
Call SendData(ToIndex, UserIndex, 0, "||5°" & "Colocaros uno a cada lado." & "°" & Npclist(TempCharIndex).Char.CharIndex)

    End If
    
    
Else
Call SendData(ToIndex, UserIndex, 0, "||5°" & "Mapa ocupado, intentalo más tarde." & "°" & Npclist(TempCharIndex).Char.CharIndex)
End If

Exit Sub


'pluto:2.24-----------------------
Case 27
    If TieneObjetos(NumeroObjEvento, CantEntregarObjEvento, UserIndex) Then
    Call QuitarObjetos(NumeroObjEvento, CantEntregarObjEvento, UserIndex)
    Dim DazA As Byte
    Dim ObjGanado As Integer
    DazA = RandomNumber(1, 100)
    
    Select Case DazA
    Case Is < 25
    ObjGanado = ObjRecompensaEventos(1)
    Case 25 To 50
    ObjGanado = ObjRecompensaEventos(2)
    Case 51 To 75
    ObjGanado = ObjRecompensaEventos(3)
    Case 76 To 100
    ObjGanado = ObjRecompensaEventos(4)
    End Select
    
    MiObj.Amount = CantObjRecompensa
    MiObj.ObjIndex = ObjGanado
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If

   Call SendData(ToIndex, UserIndex, 0, "||5°" & "Muy bién!! me quedo con " & CantEntregarObjEvento & " " & ObjData(NumeroObjEvento).Name & " ¿Te gusta la recompensa? si quieres más visitame de nuevo." & "°" & Npclist(TempCharIndex).Char.CharIndex)

    Else ' NO TIENE SUFICIENTE
    Call SendData(ToIndex, UserIndex, 0, "||5°" & "Para recibir tu recompensa necesito que me traigas al menos " & CantEntregarObjEvento & " " & ObjData(NumeroObjEvento).Name & "°" & Npclist(TempCharIndex).Char.CharIndex)
   End If
           
Exit Sub
  'pluto:6.0A
Case 3
UserList(UserIndex).flags.TargetNpc = TempCharIndex
Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNpc)
Exit Sub

Case 30
UserList(UserIndex).flags.TargetNpc = TempCharIndex
Call IniciarBovedaClan(UserIndex)
Exit Sub
'----------click npc niñera---------------------
Case 23
'tiempo embarazo
If UserList(UserIndex).Embarazada >= TimeEmbarazo Then

Dim Tindex As Integer
Tindex = NameIndex(UserList(UserIndex).Esposa)
If Tindex = 0 Then
Call SendData(ToIndex, UserIndex, 0, "||5°" & "Tu pareja debe estar presente!!" & "°" & Npclist(TempCharIndex).Char.CharIndex)
Exit Sub
End If

If Distancia(UserList(UserIndex).Pos, UserList(Tindex).Pos) < 8 Then
'tiene el niño
Call SendData(ToIndex, UserIndex, 0, "Z5")
Exit Sub
Else
Call SendData(ToIndex, UserIndex, 0, "||5°" & "Tu pareja debe estar presente!!" & "°" & Npclist(TempCharIndex).Char.CharIndex)
Exit Sub
End If

Else
Call SendData(ToIndex, UserIndex, 0, "||5°" & "No estás en condiciones de tener un bebé en estos momentos." & "°" & Npclist(TempCharIndex).Char.CharIndex)
End If
Exit Sub
'------------
'pluto:2.4.1
Case 20
If fortaleza <> UserList(UserIndex).GuildInfo.GuildName Then Exit Sub
Dim df As Integer
If UserList(UserIndex).Pos.X > Npclist(TempCharIndex).Pos.X Then df = 74 Else df = 80
Call WarpUserChar(UserIndex, 186, df, 71, True)

'pluto:6.0A
Case 38
If UserList(UserIndex).GuildInfo.GuildName = "" Then
Call SendData(ToIndex, UserIndex, 0, "||6°" & "No te puedo ayudar. No perteneces a ningún Clan." & "°" & Npclist(TempCharIndex).Char.CharIndex)
Exit Sub
End If
If UserList(UserIndex).GuildRef.SalaClan = 0 Then
Call SendData(ToIndex, UserIndex, 0, "||6°" & "No te puedo ayudar. Tú clan no tiene Sala de Clan." & "°" & Npclist(TempCharIndex).Char.CharIndex)
Exit Sub
End If
Call WarpUserChar(UserIndex, UserList(UserIndex).GuildRef.SalaClan, 53, 71, True)
Exit Sub
End Select




'comerciar
 '¿El NPC puede comerciar?
If Npclist(TempCharIndex).Comercia > 0 Then
'if UserList(UserIndex).flags.Comerciando = True then exit sub
'Iniciamos la rutina pa' comerciar.
    'pluto:2.17
    If Npclist(TempCharIndex).TipoItems = 888 And UserList(UserIndex).Faccion.ArmadaReal = 0 Then
       'Call SendData(ToIndex, UserIndex, 0, "S3")
    Call SendData(ToIndex, UserIndex, 0, "||6°" & "Sólo comercio con miembros de la Armada Real." & "°" & Npclist(TempCharIndex).Char.CharIndex)
    Exit Sub
    End If
UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
UserList(UserIndex).flags.TargetNpc = TempCharIndex
Call IniciarCOmercioNPC(UserIndex)
 Exit Sub
End If





 

            If Len(Npclist(TempCharIndex).Desc) > 1 Then
                'pluto:hoy
                If Npclist(TempCharIndex).NPCtype = 15 Then Call SendData(ToPCArea, UserIndex, Npclist(TempCharIndex).Pos.Map, "||8°" & PreTrivial & "°" & Npclist(TempCharIndex).Char.CharIndex)
                If Npclist(TempCharIndex).NPCtype = 16 Then Call SendData(ToIndex, UserIndex, 0, "||8°" & PreEgipto & "°" & Npclist(TempCharIndex).Char.CharIndex)
                If Npclist(TempCharIndex).NPCtype <> 22 And Npclist(TempCharIndex).NPCtype <> 15 And Npclist(TempCharIndex).NPCtype <> 16 Then Call SendData(ToIndex, UserIndex, 0, "||5°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).Char.CharIndex)
            
    
            
            
 'pluto:2.14
 If Npclist(TempCharIndex).NPCtype = 22 Then Call SendData(ToIndex, UserIndex, 0, "||5°" & "Escribe /Torneo son 100 oros.Hay " & MapInfo(194).NumUsers & " Jugadores en la sala y un Bote de " & TorneoBote & " Oros. No se caen los objetos." & "°" & Npclist(TempCharIndex).Char.CharIndex)
 '------------
            


'pluto:2.3
If Npclist(TempCharIndex).NPCtype = 19 Then
Dim userfile As String
userfile = CharPath & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".chr"
For X = 1 To 12
'If GetVar(UserFile, "MONTURA", "NIVEL" & x) > 0 Then
If UserList(UserIndex).Montura.Nivel(X) > 0 Then
If Not TieneObjetos(X + 887, 1, UserIndex) Then

'pluto:2.4.1
'If UserList(UserIndex).Stats.GLD < 1000 Then
'Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente Oro" & FONTTYPENAMES.FONTTYPE_INFO)
'Exit Sub
'End If
'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 1000
'Call SendUserStatsOro(UserIndex)
'Dim Miobj As obj
MiObj.Amount = 1
MiObj.ObjIndex = X + 887
Call LogMascotas("Recupera Cuidadora: " & UserList(UserIndex).Name & " Objeto " & MiObj.ObjIndex)

If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call SendData(ToIndex, UserIndex, 0, "||No tienes sitio en el inventario" & "´" & FontTypeNames.FONTTYPE_info)
End If
End If
End If
Next X
End If
 'FIN PLUTO:2.3
            Else
                If Npclist(TempCharIndex).MaestroUser > 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & "´" & FontTypeNames.FONTTYPE_info)


'pluto:2.4

If Npclist(TempCharIndex).NPCtype = 60 Then

Dim xx As Integer
xx = UserList(UserIndex).flags.ClaseMontura
If xx = 0 Then GoTo q:
Call SendData(ToIndex, UserIndex, 0, "|| Nombre: " & UserList(UserIndex).Montura.Nombre(xx) & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Nivel: " & UserList(UserIndex).Montura.Nivel(xx) & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Exp: " & UserList(UserIndex).Montura.exp(xx) & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Elu: " & UserList(UserIndex).Montura.Elu(xx) & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Vida: " & Npclist(TempCharIndex).Stats.MinHP & " / " & UserList(UserIndex).Montura.Vida(xx) & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "|| Golpe: " & UserList(UserIndex).Montura.Golpe(xx) & "´" & FontTypeNames.FONTTYPE_info)
End If
q:
Else
'pluto:6.8
    If UserList(UserIndex).flags.Privilegios > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(TempCharIndex).Name & " con " & Npclist(TempCharIndex).Stats.MinHP & " de vida." & "´" & FontTypeNames.FONTTYPE_info)
    Else
    Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(TempCharIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
    End If
'Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(TempCharIndex).Name & " index:" & TempCharIndex & FONTTYPENAMES.FONTTYPE_INFO)
End If
                
            End If
'-----------fin pluto:2.4-----------------
        

        
        
        'Pluto:2.18 añade mapas nuevos, vida restante npc castillos
        If MapInfo(Npclist(TempCharIndex).Pos.Map).Zona = "CASTILLO" And UserList(UserIndex).GuildInfo.GuildName <> "" Then
            Dim castiact As String
            If Npclist(TempCharIndex).Pos.Map = mapa_castillo1 Or Npclist(TempCharIndex).Pos.Map = 268 Then castiact = castillo1
            If Npclist(TempCharIndex).Pos.Map = mapa_castillo2 Or Npclist(TempCharIndex).Pos.Map = 269 Then castiact = castillo2
            If Npclist(TempCharIndex).Pos.Map = mapa_castillo3 Or Npclist(TempCharIndex).Pos.Map = 270 Then castiact = castillo3
            If Npclist(TempCharIndex).Pos.Map = mapa_castillo4 Or Npclist(TempCharIndex).Pos.Map = 271 Then castiact = castillo4
            If Npclist(TempCharIndex).Pos.Map = 185 Then castiact = fortaleza

            If UserList(UserIndex).GuildInfo.GuildName = castiact Then
                Call SendData(ToIndex, UserIndex, 0, "||3°" & Npclist(TempCharIndex).Stats.MinHP & "°" & Npclist(TempCharIndex).Char.CharIndex)
            End If
        End If
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNpc = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
        
    End If
AI:
    If foundchar = 0 Then
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
        'Call SendData(ToIndex, UserIndex, 0, "M9")
    End If

Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
        'Call SendData(ToIndex, UserIndex, 0, "M9")
    End If
End If

Exit Sub
fallo:
'Call LogError("LOOKATTILE" & Err.Number & " D: " & Err.Description)
Call LogError("LOOKATTILE " & Err.number & " D: " & Err.Description & " name: " & UserList(UserIndex).Name & " mapa: " & UserList(UserIndex).Pos.Map & " X: " & UserList(UserIndex).Pos.X & " Y: " & UserList(UserIndex).Pos.Y & " RatX: " & X & " RatY: " & Y)

End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As Byte
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
On Error GoTo fallo
Dim X As Integer
Dim Y As Integer

X = Pos.X - Target.X
Y = Pos.Y - Target.Y

'NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = WEST
    Exit Function
End If

'SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = WEST
    Exit Function
End If

'SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'Sur
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'norte
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'oeste
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = WEST
    Exit Function
End If

'este
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = EAST
    Exit Function
End If

'misma
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

Exit Function
fallo:
Call LogError("FINDDIRECTION" & Err.number & " D: " & Err.Description)

End Function


Public Function EsObjetoFijo(ByVal OBJType As Integer) As Boolean

EsObjetoFijo = OBJType = OBJTYPE_FOROS Or _
               OBJType = OBJTYPE_CARTELES Or _
               OBJType = OBJTYPE_ARBOLES Or _
               OBJType = OBJTYPE_YACIMIENTO

End Function
