Attribute VB_Name = "General"
Global LeerNPCs As New clsLeerInis
Global LeerNPCsHostiles As New clsLeerInis
Option Explicit

Sub VigilarEventosInvocacion(ByVal UserIndex As Integer)
If MapData(mapi, mapix1, mapiy1).UserIndex > 0 And MapData(mapi, mapix2, mapiy2).UserIndex > 0 And MapData(mapi, mapix3, mapiy3).UserIndex > 0 And MapData(mapi, mapix4, mapiy4).UserIndex > 0 And MapInfo(mapi).invocado = 0 Then
 Call SendData(ToAll, 0, 0, "K5")
MapInfo(mapi).invocado = 1
Dim bichosala
bichosala = 585
Call SpawnNpc(bichosala, UserList(MapData(mapi, mapix1, mapiy1).UserIndex).Pos, True, False)
End If
End Sub
Sub VigilarEventosCasas(ByVal UserIndex As Integer)
'pluto:2.17
Dim X As Byte
Dim Y As Byte
Dim Map As Integer
Map = UserList(UserIndex).Pos.Map
X = UserList(UserIndex).Pos.X
Y = UserList(UserIndex).Pos.Y

'pluto:rayos puerta
    If (X = 49 Or X = 50) And Y = 56 And UserList(UserIndex).flags.Muerto = 0 Then
    Call SendData(ToMap, UserIndex, Map, "TW" & 108)
    End If
'pluto:sala sangre casa
If UserList(UserIndex).flags.Privilegios > 0 Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
If X < 80 Or Y < 48 Or Y > 51 Then Exit Sub
If ((X = 85 And Y = 50) Or (X = 81 And Y = 49) Or (X = 80 And Y = 51) Or (X = 80 And Y = 48) Or (X = 88 And Y = 48) Or (X = 89 And Y = 51) Or (X = 92 And Y = 51)) Then
'If (MapData(Mapcasa, 85, 50).Userindex > 0 Or MapData(Mapcasa2, 81, 49).Userindex > 0 Or MapData(Mapcasa2, 80, 51).Userindex > 0 Or MapData(Mapcasa2, 80, 48).Userindex > 0 Or MapData(Mapcasa2, 88, 48).Userindex > 0 Or MapData(Mapcasa2, 89, 51).Userindex > 0 Or MapData(Mapcasa2, 92, 51).Userindex > 0) And (UserList(i).flags.Muerto = 0) And (UserList(i).flags.Privilegios = 0) And UserList(i).Pos.Y > 46 And UserList(i).Pos.Y < 52 And UserList(i).Pos.X > 78 And UserList(i).Pos.X < 93 Then
Call SendData(ToIndex, UserIndex, 0, "|| La Habitación de Sangre te ha matado." & "´" & FontTypeNames.FONTTYPE_TALK)
Call UserDie(UserIndex)
Call SendData(ToMap, 0, Map, "TW" & 115)
End If


'------------------------------------


End Sub
Sub VigilarEventosTrampas(ByVal UserIndex As Integer)
'pluto:6.0A --> trampas
Dim X As Byte
Dim Y As Byte
Dim Map As Integer
Map = UserList(UserIndex).Pos.Map
X = UserList(UserIndex).Pos.X
Y = UserList(UserIndex).Pos.Y
If (X = 71 And Y = 41) Or (X = 67 And Y = 66) Or (X = 55 And Y = 62) Or (X = 47 And Y = 54) Or (X = 40 And Y = 25) Or (X = 12 And Y = 11) Or (X = 92 And Y = 43) Then
Call Trampa(UserIndex, 34)
End If
If Map = 178 And ((X = 52 And Y = 54) Or (X = 15 And Y = 40) Or (X = 29 And Y = 61) Or (X = 35 And Y = 22) Or (X = 56 And Y = 22) Or (X = 70 And Y = 82)) Then
Call Trampa(UserIndex, 37)
End If

'Gusano

If RandomNumber(1, 100) > 98 Then Call Gusano(UserIndex)

End Sub
'pluto:hoy
Sub ContinuarQuest(ByVal UserIndex As Integer)
'on error GoTo fallo



Dim npcfile As String
Dim misi As String
'pluto:7.0
Dim NpcQuest As Integer
Dim TotalQuest As Integer
Dim n As Integer
NpcQuest = UserList(UserIndex).Mision.NpcQuest
npcfile = DatPath & "QUEST" & NpcQuest & ".TXT"
TotalQuest = GetVar(npcfile, "INIT", "TotalQuest")

If UserList(UserIndex).Mision.Cargada = False Then
Call CargarQuest(UserIndex)
UserList(UserIndex).Mision.Cargada = True
End If

If UserList(UserIndex).Mision.NEnemigos > 0 Then
Dim Objeton As String
'Dim Consegui As String
Dim NomNpc As String
Objeton = ""

Objeton = UserList(UserIndex).Mision.NEnemigos
For n = 1 To UserList(UserIndex).Mision.NEnemigos
NomNpc = GetVar(DatPath & "NPCs-HOSTILES.dat", "NPC" & ReadField(1, UserList(UserIndex).Mision.Enemigo(n), 45), "Name")
Objeton = Objeton & "," & NomNpc & "," & UserList(UserIndex).Mision.NEnemigosConseguidos(n)
Next
End If
'Debug.Print Objeton
'Call SendData2(ToIndex, UserIndex, 0, 109, NpcQuest & "," & UserList(UserIndex).Mision.Numero & "," & UserList(UserIndex).Mision.Actual & "," & UserList(UserIndex).Mision.NEnemigos & "," & Objeton & "," & UserList(UserIndex).Mision.TimeMision)

'tiempo invertido
Dim t1 As String
Dim t2 As String
Dim T3 As Integer
Dim T4 As Integer
t1 = UserList(UserIndex).Mision.TimeComienzo
t2 = Now
T3 = DateDiff("n", t1, t2)
T4 = UserList(UserIndex).Mision.TimeMision - T3
If T4 < 0 Then T4 = 0
Call SendData2(ToIndex, UserIndex, 0, 109, NpcQuest & "," & UserList(UserIndex).Mision.numero & "," & UserList(UserIndex).Mision.Actual & "," & UserList(UserIndex).Mision.NEnemigos & "," & UserList(UserIndex).Mision.PjConseguidos & "," & T4 & "," & Objeton)


Exit Sub
fallo:
Call LogError("CONTINUARQUEST" & Err.number & " D: " & Err.Description)

End Sub
'pluto:7.0
Function ComprobarObjetivos(ByVal UserIndex As Integer) As Boolean
Dim npcfile As String
Dim misi As String
Dim NpcQuest As Integer
Dim TotalQuest As Integer
Dim n As Integer
NpcQuest = UserList(UserIndex).Mision.NpcQuest
npcfile = DatPath & "QUEST" & NpcQuest & ".TXT"
TotalQuest = GetVar(npcfile, "INIT", "TotalQuest")
ComprobarObjetivos = True

'tiempo invertido
Dim t1 As String
Dim t2 As String
Dim T3 As Integer
t1 = UserList(UserIndex).Mision.TimeComienzo
t2 = Now
T3 = DateDiff("n", t1, t2)

If T3 > UserList(UserIndex).Mision.TimeMision Then
ComprobarObjetivos = False
Call SendData(ToIndex, UserIndex, 0, "|| Tiempo acabado, debes abortar la misión." & "´" & FontTypeNames.FONTTYPE_info)
Exit Function
End If

'enemigos matados
If UserList(UserIndex).Mision.NEnemigos > 0 Then
    For n = 1 To UserList(UserIndex).Mision.NEnemigos
        If UserList(UserIndex).Mision.NEnemigosConseguidos(n) < UserList(UserIndex).Mision.EnemigoCantidad(n) Then
        ComprobarObjetivos = False
        Exit Function
        End If
    Next
End If
'pjs matados
'GoTo a:
If UserList(UserIndex).Mision.Level > 0 Then
    If UserList(UserIndex).Mision.PjConseguidos < UserList(UserIndex).Mision.Cantidad Then
    ComprobarObjetivos = False
    Exit Function
    End If
End If
'a:
'comprobar inventario
If UserList(UserIndex).Mision.NObjetos > 0 Then
    For n = 1 To UserList(UserIndex).Mision.NObjetos
    If Not TieneObjetos(val(ReadField(1, UserList(UserIndex).Mision.Objeto(n), 45)), val(ReadField(2, UserList(UserIndex).Mision.Objeto(n), 45)), UserIndex) Then
    ComprobarObjetivos = False
    Exit Function
    End If
    Next
End If

End Function
'pluto:hoy
Sub iniciarquest(ByVal UserIndex As Integer)
'on error GoTo fallo
Dim npcfile As String
Dim misi As String
'pluto:7.0
Dim NpcQuest As Integer
Dim TotalQuest As Integer
Dim n As Integer
NpcQuest = Npclist(UserList(UserIndex).flags.TargetNpc).numero
UserList(UserIndex).Mision.NpcQuest = NpcQuest
npcfile = DatPath & "QUEST" & NpcQuest & ".TXT"
TotalQuest = GetVar(npcfile, "INIT", "TotalQuest")
'quitar esto
UserList(UserIndex).Mision.estado = 0
UserList(UserIndex).Mision.Actual = 0
'UserList(UserIndex).Mision.Actual1 = 0


If UserList(UserIndex).Mision.estado = 0 Then 'And UserList(UserIndex).Mision.Actual < TotalQuest Then

Dim Aleo As Integer
Select Case NpcQuest
Case 269
UserList(UserIndex).Mision.Actual1 = UserList(UserIndex).Mision.Actual1 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual1
Case 270
UserList(UserIndex).Mision.Actual2 = UserList(UserIndex).Mision.Actual2 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual2
Case 271
UserList(UserIndex).Mision.Actual3 = UserList(UserIndex).Mision.Actual3 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual3
Case 272
UserList(UserIndex).Mision.Actual4 = UserList(UserIndex).Mision.Actual4 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual4
'Case 273
'UserList(UserIndex).Mision.Actual5 = UserList(UserIndex).Mision.Actual5 + 1
'UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual5
Case 273

Case 302
UserList(UserIndex).Mision.Actual6 = UserList(UserIndex).Mision.Actual6 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual6
Case 303
UserList(UserIndex).Mision.Actual7 = UserList(UserIndex).Mision.Actual7 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual7
Case 304
UserList(UserIndex).Mision.Actual8 = UserList(UserIndex).Mision.Actual8 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual8
Case 305
UserList(UserIndex).Mision.Actual9 = UserList(UserIndex).Mision.Actual9 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual9
Case 306
UserList(UserIndex).Mision.Actual10 = UserList(UserIndex).Mision.Actual10 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual10
Case 307
UserList(UserIndex).Mision.Actual11 = UserList(UserIndex).Mision.Actual11 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual11
Case 308
UserList(UserIndex).Mision.Actual12 = UserList(UserIndex).Mision.Actual12 + 1
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual12

OtraQuest:
Aleo = RandomNumber(1, TotalQuest)
UserList(UserIndex).Mision.Actual = Aleo
UserList(UserIndex).Mision.SoloClase = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "SoloClase")
UserList(UserIndex).Mision.NivelMaximo = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "NivelMaximo"))
UserList(UserIndex).Mision.NivelMinimo = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "NivelMinimo"))

     'restricciones por clase
    If UserList(UserIndex).Mision.SoloClase <> "" Then
        If UCase$(UserList(UserIndex).Mision.SoloClase) <> UCase$(UserList(UserIndex).clase) Then GoTo OtraQuest
    End If
    'restricciones por niveles
    If UserList(UserIndex).Mision.NivelMaximo < UserList(UserIndex).Stats.ELV Then GoTo OtraQuest
    If UserList(UserIndex).Mision.NivelMinimo > UserList(UserIndex).Stats.ELV Then GoTo OtraQuest

UserList(UserIndex).Mision.Actual5 = Aleo

End Select



'UserList(UserIndex).Mision.Numero = UserList(UserIndex).Mision.Numero + 1
'UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual + 1
UserList(UserIndex).Mision.estado = 1
UserList(UserIndex).Mision.Titulo = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Titulo")
UserList(UserIndex).Mision.tX = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Tx")
UserList(UserIndex).Mision.TimeMision = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Time"))
'guardamos hora comienzo
UserList(UserIndex).Mision.TimeComienzo = Now


'----Enemigos a matar------------
UserList(UserIndex).Mision.NEnemigos = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "NEnemigos"))
If UserList(UserIndex).Mision.NEnemigos > 0 Then

For n = 1 To UserList(UserIndex).Mision.NEnemigos
'Debug.Print GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Enemigo" & n)
UserList(UserIndex).Mision.NEnemigosConseguidos(n) = 0
UserList(UserIndex).Mision.Enemigo(n) = val(ReadField(1, GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Enemigo" & n), 45))
UserList(UserIndex).Mision.EnemigoCantidad(n) = val(ReadField(2, GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Enemigo" & n), 45))
Next
End If
'--------------------------------
'----Objetos a conseguir---------
UserList(UserIndex).Mision.NObjetos = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "NObjetos"))
If UserList(UserIndex).Mision.NObjetos > 0 Then
'ReDim UserList(UserIndex).Mision.Objeto(1 To UserList(UserIndex).Mision.NObjetos)
For n = 1 To UserList(UserIndex).Mision.NObjetos
'Debug.Print GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Objeto" & n)

UserList(UserIndex).Mision.Objeto(n) = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Objeto" & n)
Next
End If
'-------------------------------
UserList(UserIndex).Mision.exp = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Exp")
UserList(UserIndex).Mision.oro = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Oro")
'----Objetos recompensa---------
UserList(UserIndex).Mision.NObjetosR = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "NObjetosR"))
If UserList(UserIndex).Mision.NObjetosR > 0 Then

For n = 1 To UserList(UserIndex).Mision.NObjetosR
UserList(UserIndex).Mision.ObjetoR(n) = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "ObjetoR" & n)
Next
End If
'-------------------------------

UserList(UserIndex).Mision.Level = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Level"))
UserList(UserIndex).Mision.Entrega = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Entrega"))
UserList(UserIndex).Mision.Cantidad = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Cantidad"))
UserList(UserIndex).Mision.clase = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Clase")

End If

If UserList(UserIndex).Mision.Actual = TotalQuest Then
UserList(UserIndex).Stats.Puntos = UserList(UserIndex).Stats.Puntos + 2000
'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + 5000000
Call AddtoVar(UserList(UserIndex).Stats.GLD, 5000000, MAXORO)
End If
'misi = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "tx")
'Call SendData(ToIndex, UserIndex, 0, "!!Quest Número " & UserList(UserIndex).Mision.Actual & " : " & misi)

Dim Objeton As String
'Dim Consegui As String
Dim NomNpc As String
Objeton = ""

Objeton = UserList(UserIndex).Mision.NEnemigos
For n = 1 To UserList(UserIndex).Mision.NEnemigos
NomNpc = GetVar(DatPath & "NPCs-HOSTILES.dat", "NPC" & ReadField(1, UserList(UserIndex).Mision.Enemigo(n), 45), "Name")
Objeton = Objeton & "," & NomNpc & "," & UserList(UserIndex).Mision.NEnemigosConseguidos(n)
Next


UserList(UserIndex).Mision.SoloClase = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "SoloClase")
UserList(UserIndex).Mision.NivelMaximo = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "NivelMaximo"))
UserList(UserIndex).Mision.NivelMinimo = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "NivelMinimo"))




Call SendData2(ToIndex, UserIndex, 0, 109, NpcQuest & "," & UserList(UserIndex).Mision.numero & "," & UserList(UserIndex).Mision.Actual & "," & UserList(UserIndex).Mision.NEnemigos & "," & UserList(UserIndex).Mision.PjConseguidos & "," & UserList(UserIndex).Mision.TimeMision & "," & Objeton)

If UserList(UserIndex).Stats.ELV < UserList(UserIndex).Mision.NivelMinimo Then
ResetUserMision (UserIndex)
End If

Exit Sub
fallo:
Call LogError("INICIARQUEST" & Err.number & " D: " & Err.Description)

End Sub
Sub CargarQuest(ByVal UserIndex As Integer)
'on error GoTo fallo
Dim npcfile As String
Dim misi As String
'pluto:7.0
Dim NpcQuest As Integer
Dim TotalQuest As Integer
Dim n As Integer
NpcQuest = UserList(UserIndex).Mision.NpcQuest 'Npclist(UserList(UserIndex).flags.TargetNpc).Numero
'UserList(UserIndex).Mision.NpcQuest = NpcQuest
npcfile = DatPath & "QUEST" & NpcQuest & ".TXT"
TotalQuest = GetVar(npcfile, "INIT", "TotalQuest")

Select Case NpcQuest
Case 269 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual1
Case 270 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual2
Case 271 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual3
Case 272 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual4
Case 273 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual5
Case 302 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual6
Case 303 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual7
Case 304 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual8
Case 305 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual9
Case 306 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual10
Case 307 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual11
Case 308 'npcquest actual
UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual12
End Select



'UserList(UserIndex).Mision.Numero = UserList(UserIndex).Mision.Numero + 1
'UserList(UserIndex).Mision.Actual = UserList(UserIndex).Mision.Actual + 1
'UserList(UserIndex).Mision.estado = 1
UserList(UserIndex).Mision.Titulo = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Titulo")
UserList(UserIndex).Mision.tX = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Tx")
UserList(UserIndex).Mision.TimeMision = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Time"))

'----Enemigos a matar------------
UserList(UserIndex).Mision.NEnemigos = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "NEnemigos"))
If UserList(UserIndex).Mision.NEnemigos > 0 Then
'ReDim UserList(UserIndex).Mision.Enemigo(1 To UserList(UserIndex).Mision.NEnemigos)
'ReDim UserList(UserIndex).Mision.NEnemigosConseguidos(1 To UserList(UserIndex).Mision.NEnemigos)
'ReDim UserList(UserIndex).Mision.EnemigoCantidad(1 To UserList(UserIndex).Mision.NEnemigos)

For n = 1 To UserList(UserIndex).Mision.NEnemigos
'Debug.Print GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Enemigo" & n)
'UserList(UserIndex).Mision.NEnemigosConseguidos(n) = 0
UserList(UserIndex).Mision.Enemigo(n) = val(ReadField(1, GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Enemigo" & n), 45))
UserList(UserIndex).Mision.EnemigoCantidad(n) = val(ReadField(2, GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Enemigo" & n), 45))
Next
End If
'--------------------------------
'----Objetos a conseguir---------
UserList(UserIndex).Mision.NObjetos = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "NObjetos"))
If UserList(UserIndex).Mision.NObjetos > 0 Then
'ReDim UserList(UserIndex).Mision.Objeto(1 To UserList(UserIndex).Mision.NObjetos)
For n = 1 To UserList(UserIndex).Mision.NObjetos
'Debug.Print GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Objeto" & n)

UserList(UserIndex).Mision.Objeto(n) = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Objeto" & n)
Next
End If
'-------------------------------
UserList(UserIndex).Mision.exp = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Exp")
UserList(UserIndex).Mision.oro = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Oro")
'----Objetos recompensa---------
UserList(UserIndex).Mision.NObjetosR = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "NObjetosR"))
If UserList(UserIndex).Mision.NObjetosR > 0 Then
'ReDim UserList(UserIndex).Mision.ObjetoR(1 To UserList(UserIndex).Mision.NObjetosR)
For n = 1 To UserList(UserIndex).Mision.NObjetosR
'Debug.Print GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "ObjetoR" & n)

UserList(UserIndex).Mision.ObjetoR(n) = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "ObjetoR" & n)
Next
End If
'-------------------------------

UserList(UserIndex).Mision.Level = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Level"))
UserList(UserIndex).Mision.Entrega = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Entrega"))
UserList(UserIndex).Mision.Cantidad = val(GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Cantidad"))
'UserList(UserIndex).Mision.Objeto = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Objeto")
'UserList(UserIndex).Mision.Enemigo = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Enemigo")
UserList(UserIndex).Mision.clase = GetVar(npcfile, "QUEST" & UserList(UserIndex).Mision.Actual, "Clase")

'End If


'Dim Objeton As String
'Dim Consegui As String
'Dim NomNpc As String
'Objeton = ""

'Objeton = UserList(UserIndex).Mision.NEnemigos
'For n = 1 To UserList(UserIndex).Mision.NEnemigos
'NomNpc = GetVar(DatPath & "NPCs-HOSTILES.dat", "NPC" & ReadField(1, UserList(UserIndex).Mision.Enemigo(n), 45), "Name")
'Objeton = Objeton & "," & NomNpc & "," & UserList(UserIndex).Mision.NEnemigosConseguidos(n)
'Next
'Debug.Print Objeton
'Call SendData2(ToIndex, UserIndex, 0, 109, NpcQuest & "," & UserList(UserIndex).Mision.Numero & "," & UserList(UserIndex).Mision.Actual & "," & UserList(UserIndex).Mision.NEnemigos & "," & Objeton & "," & UserList(UserIndex).Mision.TimeMision)

Exit Sub
fallo:
Call LogError("CARGARQUEST" & Err.number & " D: " & Err.Description)

End Sub
Sub DarCuerpoDesnudo(ByVal UserIndex As Integer)
On Error GoTo fallo

'PLUTO:2.15
If UserList(UserIndex).Bebe > 0 Then

If UserList(UserIndex).raza = "Vampiro" Or UserList(UserIndex).raza = "Elfo Oscuro" Then
UserList(UserIndex).Char.Body = 342
ElseIf UserList(UserIndex).raza = "Orco" Then
UserList(UserIndex).Char.Body = 341
Else
UserList(UserIndex).Char.Body = 340
End If


UserList(UserIndex).Char.Head = 0
UserList(UserIndex).flags.Desnudo = 1
Exit Sub
End If
'------------

If UserList(UserIndex).Remort = 0 Then
Select Case UCase$(UserList(UserIndex).raza)
    Case "HUMANO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 21
                     'pluto:6.5
                     If UserList(UserIndex).flags.DragCredito3 = 1 Then UserList(UserIndex).Char.Body = 425
                        If UserList(UserIndex).flags.DragCredito3 = 2 Then UserList(UserIndex).Char.Body = 424

                Case "MUJER"
                     UserList(UserIndex).Char.Body = 39
      End Select
      
        Case "CICLOPE"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 351
                     'pluto:6.5
                     If UserList(UserIndex).flags.DragCredito3 = 1 Then UserList(UserIndex).Char.Body = 425
                        If UserList(UserIndex).flags.DragCredito3 = 2 Then UserList(UserIndex).Char.Body = 424

                Case "MUJER"
                     UserList(UserIndex).Char.Body = 351
      End Select
      
      
    Case "ELFO OSCURO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 32
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 40
      End Select
    
     Case "VAMPIRO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 32
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 40
      End Select
     Case "ORCO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 215
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 217
      End Select
    
    
    Case "ENANO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 53
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 60
      End Select
    Case "GNOMO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 53
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 60
      End Select
       'pluto:7.0
       Case "GOBLIN"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 178
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 212
      End Select
      
      
    Case Else
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 21
                     'pluto:6.5
                     If UserList(UserIndex).flags.DragCredito3 = 1 Then UserList(UserIndex).Char.Body = 425
                        If UserList(UserIndex).flags.DragCredito3 = 2 Then UserList(UserIndex).Char.Body = 424

                Case "MUJER"
                     UserList(UserIndex).Char.Body = 39
      End Select
    
End Select
End If
If UserList(UserIndex).Remort = 1 Then
'pluto:2-3-04
Select Case UCase$(UserList(UserIndex).raza)
    Case "HUMANO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 264
                      'pluto:6.5
                    If UserList(UserIndex).flags.DragCredito3 = 1 Then UserList(UserIndex).Char.Body = 425
                    If UserList(UserIndex).flags.DragCredito3 = 2 Then UserList(UserIndex).Char.Body = 424

                Case "MUJER"
                     UserList(UserIndex).Char.Body = 266
      End Select
      
        Case "CICLOPE"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 351
                      'pluto:6.5
                    If UserList(UserIndex).flags.DragCredito3 = 1 Then UserList(UserIndex).Char.Body = 425
                    If UserList(UserIndex).flags.DragCredito3 = 2 Then UserList(UserIndex).Char.Body = 424

                Case "MUJER"
                     UserList(UserIndex).Char.Body = 351
      End Select
      
    Case "ELFO OSCURO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 265
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 267
      End Select
    
     Case "VAMPIRO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 265
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 267
      End Select
     Case "ORCO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 268
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 269
      End Select
    
    
    Case "ENANO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 270
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 271
      End Select
      'PLUTO:7.0
      Case "GOBLIN"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 178
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 212
      End Select
      
    Case "GNOMO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                     UserList(UserIndex).Char.Body = 270
                Case "MUJER"
                     UserList(UserIndex).Char.Body = 271
      End Select
    Case Else
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                      UserList(UserIndex).Char.Body = 264
                      'pluto:6.5
                    If UserList(UserIndex).flags.DragCredito3 = 1 Then UserList(UserIndex).Char.Body = 425
                    If UserList(UserIndex).flags.DragCredito3 = 2 Then UserList(UserIndex).Char.Body = 424

                Case "MUJER"
                     UserList(UserIndex).Char.Body = 266
      End Select
    
End Select
End If
UserList(UserIndex).flags.Desnudo = 1

Exit Sub
fallo:
Call LogError("DARCUERPODESNUDO" & Err.number & " D: " & Err.Description)

End Sub


Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, ByVal X As Integer, ByVal Y As Integer, b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado
On Error GoTo fallo
Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)
Exit Sub
fallo:
Call LogError("BLOQUEAR" & Err.number & " D: " & Err.Description)

End Sub


Function HayAgua(Map As Integer, X As Integer, Y As Integer) As Boolean
On Error GoTo fallo
If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(Map, X, Y).Graphic(1) >= 1505 And _
       MapData(Map, X, Y).Graphic(1) <= 1520 And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

Exit Function
fallo:
Call LogError("HAYAGUA" & Err.number & " D: " & Err.Description)

End Function




Sub LimpiarMundo()

On Error GoTo fallo

Dim i As Integer

For i = 1 To TrashCollector.Count
    Dim d As cGarbage
    Set d = TrashCollector(1)
    Call EraseObj(ToMap, 0, d.Map, 1, d.Map, d.X, d.Y)
    Call TrashCollector.Remove(1)
    Set d = Nothing
Next i
'pluto:2.23------------
'Call securityip.IpSecurityMantenimientoLista
'----------------------
Exit Sub
fallo:
Call LogError("LIMPIARMUNDO" & Err.number & " D: " & Err.Description)


End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim k As Integer, SD As String
SD = UBound(SpawnList) & ","
For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k
Call SendData2(ToIndex, UserIndex, 0, 42, SD)
Exit Sub
fallo:
Call LogError("ENVIARSPAWNLIST " & Err.number & " D: " & Err.Description)

End Sub

Sub ConfigListeningSocket(ByRef obj As Object, ByVal Port As Integer)
#If UsarQueSocket = 0 Then
On Error GoTo fallo
obj.AddressFamily = AF_INET
obj.Protocol = IPPROTO_IP
obj.SocketType = SOCK_STREAM
obj.Binary = False
obj.Blocking = False
obj.BufferSize = 1024
obj.LocalPort = Port
obj.backlog = 5
obj.Listen
Exit Sub
fallo:
Call LogError("CONFIGLISTENINGSOCKET " & Err.number & " D: " & Err.Description)
#End If
End Sub




Sub Main()
On Error GoTo fallo

Call SetKey("CLIENTE AODRAG v5.0")
Call LoadMotd

Prision.Map = 66
Libertad.Map = 66

Prision.X = 75
Prision.Y = 47
Libertad.X = 75
Libertad.Y = 65


LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")



ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer


IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"

'pluto:6.8-------------------------
Randomize Timer
HOYESDIA = Date$
'pluto:6.8
EventoDia = val(GetVar(IniPath & "eventodia.txt", "INIT", "Evento"))

'pluto:6.9
If EventoDia = 1 Or EventoDia = 4 Then
Call CargarDiaEspecial
End If
'----------------------

'nati: ahora sube de 5 en vez de a 3.
LevelSkill(1).LevelValue = 5
LevelSkill(2).LevelValue = 10
LevelSkill(3).LevelValue = 15
LevelSkill(4).LevelValue = 20
LevelSkill(5).LevelValue = 25
LevelSkill(6).LevelValue = 30
LevelSkill(7).LevelValue = 35
LevelSkill(8).LevelValue = 40
LevelSkill(9).LevelValue = 45
LevelSkill(10).LevelValue = 50
LevelSkill(11).LevelValue = 55
LevelSkill(12).LevelValue = 60
LevelSkill(13).LevelValue = 65
LevelSkill(14).LevelValue = 70
LevelSkill(15).LevelValue = 75
LevelSkill(16).LevelValue = 80
LevelSkill(17).LevelValue = 85
LevelSkill(18).LevelValue = 90
LevelSkill(19).LevelValue = 100
LevelSkill(20).LevelValue = 105
LevelSkill(21).LevelValue = 110
LevelSkill(22).LevelValue = 115
LevelSkill(23).LevelValue = 120
LevelSkill(24).LevelValue = 125
LevelSkill(25).LevelValue = 130
LevelSkill(26).LevelValue = 135
LevelSkill(27).LevelValue = 140
LevelSkill(28).LevelValue = 145
LevelSkill(29).LevelValue = 150
LevelSkill(30).LevelValue = 155
LevelSkill(31).LevelValue = 160
LevelSkill(32).LevelValue = 165
LevelSkill(33).LevelValue = 170
LevelSkill(34).LevelValue = 175
LevelSkill(35).LevelValue = 180
LevelSkill(36).LevelValue = 185
LevelSkill(37).LevelValue = 190
LevelSkill(38).LevelValue = 195
LevelSkill(39).LevelValue = 200
LevelSkill(40).LevelValue = 200
LevelSkill(41).LevelValue = 200
LevelSkill(42).LevelValue = 200
LevelSkill(43).LevelValue = 200
LevelSkill(44).LevelValue = 200
LevelSkill(45).LevelValue = 200
LevelSkill(46).LevelValue = 200
LevelSkill(47).LevelValue = 200
LevelSkill(48).LevelValue = 200
LevelSkill(49).LevelValue = 200
LevelSkill(50).LevelValue = 200
LevelSkill(51).LevelValue = 200
LevelSkill(52).LevelValue = 200
LevelSkill(53).LevelValue = 200
LevelSkill(54).LevelValue = 200
LevelSkill(55).LevelValue = 200
LevelSkill(56).LevelValue = 200
LevelSkill(57).LevelValue = 200
LevelSkill(58).LevelValue = 200
LevelSkill(59).LevelValue = 200
LevelSkill(60).LevelValue = 200
LevelSkill(61).LevelValue = 200
LevelSkill(62).LevelValue = 200
LevelSkill(63).LevelValue = 200
LevelSkill(64).LevelValue = 200
LevelSkill(65).LevelValue = 200
LevelSkill(66).LevelValue = 200
LevelSkill(67).LevelValue = 200
LevelSkill(68).LevelValue = 200
LevelSkill(69).LevelValue = 200
LevelSkill(70).LevelValue = 200
'pluto:7.0
NOmbrelogro(1) = "Animales"
NOmbrelogro(2) = "Arañas"
NOmbrelogro(3) = "Goblins"
NOmbrelogro(4) = "Orcos"
NOmbrelogro(5) = "Lagartos"
NOmbrelogro(6) = "Genios"
NOmbrelogro(7) = "Hobbits"
NOmbrelogro(8) = "Ogros"
NOmbrelogro(9) = "Npc-Magias"
NOmbrelogro(10) = "No-Muertos"
NOmbrelogro(11) = "Darks"
NOmbrelogro(12) = "Trolls"
NOmbrelogro(13) = "Beholders"
NOmbrelogro(14) = "Golems"
NOmbrelogro(15) = "Npc-Marinos"
NOmbrelogro(16) = "Ents"
NOmbrelogro(17) = "Licantropos"
NOmbrelogro(18) = "Medusas"
NOmbrelogro(19) = "Ciclopes"
NOmbrelogro(20) = "Npc-Polares"
NOmbrelogro(21) = "Devastadores"
NOmbrelogro(22) = "Gigantes"
NOmbrelogro(23) = "Piratas"
NOmbrelogro(24) = "Uruks"
NOmbrelogro(25) = "Demonios"
NOmbrelogro(26) = "Devirs"
NOmbrelogro(27) = "Gollums"
NOmbrelogro(28) = "Dragones"
NOmbrelogro(29) = "Ettins"
NOmbrelogro(30) = "Puertas"
NOmbrelogro(31) = "Reyes"
NOmbrelogro(32) = "Defensores"
NOmbrelogro(33) = "Raids"
NOmbrelogro(34) = "Npc-Navidad"

ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"
ListaRazas(6) = "Orco"
ListaRazas(7) = "Vampiro"
ListaRazas(8) = "Ciclope"
ListaRazas(9) = "Goblin"
ReDim ListaClases(1 To NUMCLASES) As String

ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Bandido"
ListaClases(9) = "Paladin"
ListaClases(10) = "Cazador"
ListaClases(11) = "Pescador"
ListaClases(12) = "Herrero"
ListaClases(13) = "Leñador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Pirata"
ListaClases(17) = "Ermitaño"
ListaClases(18) = "Arquero"
'pluto:2.3
ListaClases(19) = "Domador"

ReDim SkillsNames(1 To NUMSKILLS) As String

SkillsNames(1) = "Suerte"
SkillsNames(2) = "Aprendizaje de Magias"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Esquivar Cuerpo/Cuerpo"
SkillsNames(5) = "Golpear Cuerpo/Cuerpo"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar Criaturas"
SkillsNames(19) = "Golpeo con Proyectiles"
SkillsNames(20) = "Golpeo con Armas Dobles"
SkillsNames(21) = "Navegacion"
SkillsNames(22) = "Daños en Magia"
SkillsNames(23) = "Defensa en Magias"
SkillsNames(24) = "Esquivar Magias"
SkillsNames(25) = "Daño en Armas"
SkillsNames(26) = "Defensa en Armas"
SkillsNames(27) = "Aprendizaje de Armas"
SkillsNames(28) = "Daño de Proyectiles"
SkillsNames(29) = "Defensa de Proyectiles"
SkillsNames(30) = "Aprendizaje de Proyectiles"
SkillsNames(31) = "Esquivar Proyectiles"


ReDim UserSkills(1 To NUMSKILLS) As Integer
'pluto:2.3
ReDim UserMONTURA(1 To MAXMONTURA) As Integer

ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"


frmCargando.Show



Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

frmMain.Caption = frmMain.Caption & " V." & pluto1 & "." & pluto2 & "." & pluto3
ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).Caption = "Iniciando Arrays..."

Call LoadGuildsDB

Call CargarSpawnList
Call CargarForbidenWords
'¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
frmCargando.Label1(2).Caption = "Cargando Server.ini"

MaxUsers = 0


Call LoadSini
'pluto fusion
Call CargaNpcsDat

frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
Call LoadOBJData
    
frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
Call CargarHechizos

    
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero
'pluto:hoy
Call Loadtrivial
Call LoadEgipto
'pluto:2.17
Call LoadPorcentajesMascotas
'pluto:2.14
BodyTorneo = 325
'pluto:2.4
Call Loadrecord
'[MerLiNz:6]
Call LoadObjMagicosermitano
'[\END]


If BootDelBackUp Then
    
    frmCargando.Label1(2).Caption = "Cargando BackUp"
    Call CargarBackUp
Else
    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapData
End If

'pluto:2.15
Caballero = True
'[Tite] Inicializo variables de party
numPartys = 0

'pluto:6.2------------------------
frmMain.ws_server.Close
'asignamos un puerto pluto:6.3
If ServerPrimario = 2 Then
'frmMain.ws_server.LocalPort = "7665" ' para local
frmMain.ws_server.LocalPort = "10290"
Else
frmMain.ws_server.LocalPort = "7664"
End If
'ponemos a la escucha el puerto asignado
frmMain.ws_server.Listen
'----------------------------------

'pluto:2.18 'carga salidas de telport castillo delante del ettin
'Dim ns As Byte
'For ns = 166 To 169
'MapData(ns, 58, 50).TileExit.Map = ns + 102
'MapData(ns, 58, 50).TileExit.X = 40
'MapData(ns, 58, 50).TileExit.Y = 53
'Next
'----------------------
'pluto:2.4 mete criaturas cabalgar
'GoTo AUI
Dim CabalgaPos As WorldPos
Dim nz As Integer
Dim ini As Integer
Dim mapita As Integer
    CabalgaPos.X = 20 + RandomNumber(1, 50)
    CabalgaPos.Y = 20 + RandomNumber(1, 50)

'unicornios

a:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a:
ini = SpawnNpc(616, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a:
Call WriteVar(IniPath & "cabalgar.txt", "Unicornio", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)

'caballos negros

a2:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a2:
ini = SpawnNpc(617, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a2
Call WriteVar(IniPath & "cabalgar.txt", "Caballo Negro", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)

'tigres

a3:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a3:
ini = SpawnNpc(618, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a3
Call WriteVar(IniPath & "cabalgar.txt", "Tigre Blanco", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)

'dumbos

a4:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a4:
ini = SpawnNpc(619, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a4
Call WriteVar(IniPath & "cabalgar.txt", "Elefante", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)

'dragon
a5:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a5:
ini = SpawnNpc(620, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a5
Call WriteVar(IniPath & "cabalgar.txt", "Dragón Dorado", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)

'jabalí
a6:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a6:
ini = SpawnNpc(670, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a6
Call WriteVar(IniPath & "cabalgar.txt", "Jabalí Gigante", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)

'escarabajo
a7:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a7:
ini = SpawnNpc(671, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a7
Call WriteVar(IniPath & "cabalgar.txt", "Escarabajo", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)
'hipopotamo
a8:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a8:
ini = SpawnNpc(672, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a8
Call WriteVar(IniPath & "cabalgar.txt", "Rinosaurio", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)

'pantera
a9:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a9:
ini = SpawnNpc(673, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a9
Call WriteVar(IniPath & "cabalgar.txt", "Cerbero", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)

'ciervo
a10:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a10:
ini = SpawnNpc(674, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a10
Call WriteVar(IniPath & "cabalgar.txt", "Wyvern", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)

'avestruz
a11:
mapita = RandomNumber(1, 277)
CabalgaPos.Map = mapita
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a11:
ini = SpawnNpc(675, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a11
Call WriteVar(IniPath & "cabalgar.txt", "Avestruz", "Mapa", val(mapita) & " ->" & CabalgaPos.Map)

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'AUI:
Dim loopc As Integer

'Resetea las conexiones de los usuarios
For loopc = 1 To MaxUsers
UserList(loopc).ConnID = -1
UserList(loopc).ConnIDValida = False
Next loopc
'Pluto:6.2
'frmMain.Macrear.Enabled = True
'frmMain.AutoSave.Enabled = True

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

With frmMain
    .AutoSave.Enabled = True
    .tLluvia.Enabled = True
    .tPiqueteC.Enabled = True
    .Timer1.Enabled = True
    If ClientsCommandsQueue <> 0 Then
        .CmdExec.Enabled = True
    Else
        .CmdExec.Enabled = False
    End If
    .GameTimer.Enabled = True
    .tLluviaEvent.Enabled = True
    .FX.Enabled = True
    .Auditoria.Enabled = True
    .KillLog.Enabled = True
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
End With

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Configuracion de los sockets
'pluto:2.23------------------
'Call securityip.InitIpTables(1000)
'---------------------------
#If UsarQueSocket = 1 Then

Call IniciaWsApi(frmMain.hwnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then

frmCargando.Label1(2).Caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

Call ConfigListeningSocket(frmMain.Socket1, Puerto)

#ElseIf UsarQueSocket = 2 Then

frmMain.Serv.Iniciar Puerto

#ElseIf UsarQueSocket = 3 Then

frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿




Unload frmCargando


'Log
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " server iniciado " & pluto1 & "."; pluto2 & "." & pluto3
Close #n

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

'tInicioServer = GetTickCount() And &H7FFFFFFF
'Call InicializaEstadisticas

Randomize Timer
'ResetThread.CreateNewThread AddressOf ThreadResetActions, tpNormal

'Call MainThread

Exit Sub
fallo:
Call LogError("MAIN" & Err.number & " D: " & Err.Description)

End Sub

Function FileExist(file As String, FileType As VbFileAttribute) As Boolean

On Error GoTo fallo
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************

If Dir(file, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If
Exit Function
fallo:
Call LogError("FICHEROEXISTE " & Err.number & " D: " & Err.Description & " File: " & file & " Filetype: " & FileType)

End Function
Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'All these functions are much faster using the "$" sign
'after the function. This happens for a simple reason:
'The functions return a variant without the $ sign. And
'variants are very slow, you should never use them.
On Error GoTo fallo
'*****************************************************************
'Devuelve el string del campo
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = Mid$(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i

FieldNum = FieldNum + 1
If FieldNum = Pos Then
    ReadField = Mid$(Text, LastPos + 1)
End If
Exit Function
fallo:
Call LogError("READFIELD" & Err.number & " D: " & Err.Description)

End Function
Public Function Tilde(Data As String) As String
 
Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(Data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
 
End Function
Function MapaValido(ByVal Map As Integer) As Boolean
On Error GoTo fallo
MapaValido = Map >= 1 And Map <= NumMaps
Exit Function
fallo:
Call LogError("MAPAVALIDO" & Err.number & " D: " & Err.Description)

End Function

Sub MostrarNumUsers()
On Error GoTo fallo
frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers
Exit Sub
fallo:
Call LogError("MOSTRARNUMUSERS" & Err.number & " D: " & Err.Description)

End Sub


Public Sub LogCriticEvent(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub


fallo:
Call LogError("LOGCRITICEVENT" & Err.number & " D: " & Err.Description)


End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub
fallo:
Call LogError("LOGEJERCITOREAL" & Err.number & " D: " & Err.Description)


End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub
fallo:
Call LogError("LOGEJERCITOCAOS" & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogInitModificados(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\InitModificados.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

'Dim tindex As Integer
'tindex = NameIndex("AoDraGBoT")
'If tindex <= 0 Then Exit Sub
'Call SendData(ToIndex, tindex, 0, "|| Error: " & Desc & FONTTYPENAMES.FONTTYPE_TALK)


Exit Sub
fallo:
Call LogError("LOGINITMODIFICADOS" & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogMapa191(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Bloqueo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub
fallo:
Call LogError("LOGBLOQUEO " & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogDonaciones(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\donaciones.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub
fallo:
Call LogError("LOGDONACIONES " & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogCasino(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\casino.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub
fallo:
Call LogError("LOGCASINO " & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogTeclado(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\teclado.log" For Append Shared As #nfile
Print #nfile, Date & " " & Format(Time, "hh/mm/ss") & " " & Desc
Close #nfile

Exit Sub
fallo:
Call LogError("LOGTECLADO " & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogParty(Desc As String)
On Error GoTo fallo

'pluto:6.5
'Exit Sub

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\party.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

'pluto:6.6
Exit Sub
Dim Tindex As Integer
Tindex = NameIndex("AoDraGBoT")
If Tindex <= 0 Then Exit Sub
Call SendData(ToIndex, Tindex, 0, "|| Error: " & Desc & "´" & FontTypeNames.FONTTYPE_TALK)

Exit Sub
fallo:
Call LogError("LOGParty " & Err.number & " D: " & Err.Description)


End Sub
Public Sub Logrenumusers(Desc As String, Desc2 As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\RecordConectados.log" For Append Shared As #nfile
Print #nfile, Date & "*" & Time & " --> " & Desc & " Máximo Usuarios."
Print #nfile, Date & "*" & Time & " --> " & Desc2 & " Usuarios de Media."

Close #nfile

Exit Sub
fallo:
Call LogError("LOGRENUMUSERS " & Err.number & " D: " & Err.Description)


End Sub

Public Sub LogCambiarPJ(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CambiarPj.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile
'Dim tindex As Integer
'tindex = NameIndex("AoDraGBoT")
'If tindex <= 0 Then Exit Sub
'Call SendData(ToIndex, tindex, 0, "|| CambiarPJ: " & Desc & FONTTYPENAMES.FONTTYPE_TALK)

Exit Sub
fallo:
Call LogError("LOGCAMBIARPJ" & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogClanMov(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\ClanMov.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile
Exit Sub
fallo:
Call LogError("LOGCLANMOV" & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogRecuperarClaves(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\RecuperaClaves.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Dim Tindex As Integer
Tindex = NameIndex("AoDraGBoT")
If Tindex <= 0 Then Exit Sub
Call SendData(ToIndex, Tindex, 0, "|| Claves : " & Desc & "´" & FontTypeNames.FONTTYPE_TALK)


Exit Sub
fallo:
Call LogError("LOGRECUPERARCLAVES " & Err.number & " D: " & Err.Description)

End Sub
Public Sub Logpass(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Pass.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub
fallo:
Call LogError("LOGPASS" & Err.number & " D: " & Err.Description)


End Sub

Public Sub LogError(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
'Print #nfile, Desc
Close #nfile
'quitar esto
'Call SendData(ToAll, 0, 0, "|| Error: " & Desc & "´" & FontTypeNames.FONTTYPE_talk)

Dim Tindex As Integer
Tindex = NameIndex("AoDraGBoT")
If Tindex <= 0 Then Exit Sub
Call SendData(ToIndex, Tindex, 0, "|| Error: " & Desc & "´" & FontTypeNames.FONTTYPE_TALK)

Exit Sub
fallo:
Call LogError("LOGERROR" & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogMascotas(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\mascotas.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Dim Tindex As Integer
Tindex = NameIndex("AoDraGBoT")
If Tindex <= 0 Then Exit Sub
Call SendData(ToIndex, Tindex, 0, "|| LogMascota: " & Desc & "´" & FontTypeNames.FONTTYPE_TALK)

Exit Sub
fallo:
Call LogError("LOGMascota" & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogNpcFundidor(Desc As String)
On Error GoTo fallo

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\NpcFundidor.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile
'pluto:6.6
'Dim Tindex As Integer
'Tindex = NameIndex("AoDraGBoT")
'If Tindex <= 0 Then Exit Sub
'Call SendData(ToIndex, Tindex, 0, "|| LogNpcfundidor: " & Desc & "´" & FontTypeNames.FONTTYPE_talk)

Exit Sub
fallo:
Call LogError("LOGNPCFUNDIDOR" & Err.number & " D: " & Err.Description)


End Sub
Public Sub LogStatic(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:


End Sub

Public Sub LogGM(Nombre As String, texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub SaveDayStats()
'on error GoTo errhandler

'Dim nfile As Integer
'nfile = FreeFile ' obtenemos un canal
'Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile

'Print #nfile, "<stats>"
'Print #nfile, "<ao>"
'Print #nfile, "<dia>" & Date & "</dia>"
'Print #nfile, "<hora>" & Time & "</hora>"
'Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
'Print #nfile, "<max_user>" & DayStats.Maxusuarios & "</max_user>"
'Print #nfile, "</ao>"
'Print #nfile, "</stats>"


'Close #nfile

Exit Sub

'errhandler:

End Sub


Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub
Function ValidInputNP(ByVal cad As String) As Boolean
On Error GoTo fallo
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If Arg = "" Then Exit Function

Next i

ValidInputNP = True
Exit Function
fallo:
Call LogError("VALIDINPUTNP" & Err.number & " D: " & Err.Description)

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error GoTo fallo

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

Dim loopc As Integer
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

ReDim UserList(1 To MaxUsers)

For loopc = 1 To MaxUsers
    UserList(loopc).ConnID = -1
    UserList(loopc).ConnIDValida = False
Next loopc

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos


#If UsarQueSocket = 0 Then

'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = val(Puerto)
frmMain.Socket1.Listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

'Log it
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " servidor reiniciado."
Close #n

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  Exit Sub
fallo:
Call LogError("RESTART" & Err.number & " D: " & Err.Description)

End Sub


Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
 On Error GoTo fallo
    If MapInfo(UserList(UserIndex).Pos.Map).Lluvia = 0 Then
        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 1 And _
           MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 2 And _
           MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
Exit Function
fallo:
Call LogError("INTEMPERIE" & Err.number & " D: " & Err.Description)
    
End Function

Public Sub EfectoLluvia(ByVal UserIndex As Integer)
On Error GoTo errhandler


If UserList(UserIndex).flags.UserLogged Then
    If Intemperie(UserIndex) Then
                Dim modifi As Long
                'pluto:2.17
                Dim ff As Byte

                ff = 4 - CInt(UserList(UserIndex).Stats.UserSkills(Supervivencia) / 50)
                modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, ff) 'ff era un 3
                If modifi > 1 Then
                Call QuitarSta(UserIndex, modifi)
                Call SendData(ToIndex, UserIndex, 0, "L1")
                Call SendUserStatsEnergia(UserIndex)
                End If
    End If
End If

Exit Sub
errhandler:
 LogError ("Error en EfectoLluvia")
End Sub


Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
    'pluto:6.9
If Npclist(UserList(UserIndex).MascotasIndex(i)).NPCtype = 60 Then GoTo nop
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
        Else
        Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0
        Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    End If
nop:

Next i

Exit Sub
fallo:
Call LogError("TIEMPOINVOCACION Pj: " & UserList(UserIndex).Name & " Mindex: " & UserList(UserIndex).MascotasIndex(i) & " E: " & Err.number & " D: " & Err.Description)

End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim modifi As Integer
Dim ff As Byte
'pluto:2.15
If UserList(UserIndex).Bebe > 0 Then Exit Sub

If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
  UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
Else
  If MapInfo(UserList(UserIndex).Pos.Map).Terreno = Nieve Then
  
    ff = 6 - CInt(UserList(UserIndex).Stats.UserSkills(Supervivencia) / 50)
 
    modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, ff)
    Call SendData(ToIndex, UserIndex, 0, "M3")
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
    If UserList(UserIndex).Stats.MinHP < 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Has muerto de frio!!." & "´" & FontTypeNames.FONTTYPE_info)
            UserList(UserIndex).Stats.MinHP = 0
            Call UserDie(UserIndex)
    End If
  Else

    ff = 6 - CInt(UserList(UserIndex).Stats.UserSkills(Supervivencia) / 50)
   
    modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, ff)
    Call QuitarSta(UserIndex, modifi)
    Call SendData(ToIndex, UserIndex, 0, "M2")
  End If

  UserList(UserIndex).Counters.Frio = 0
Call SendUserStatsVida(UserIndex)
Call SendUserStatsEnergia(UserIndex)
End If
Exit Sub
fallo:
Call LogError("EFECTOFRIO " & Err.number & " D: " & Err.Description)

End Sub
Public Sub EfectoIncor(ByVal UserIndex As Integer)
On Error GoTo fallo
'pluto:6.0A
'If MapInfo(UserList(UserIndex).Pos.Map).Invisible = 1 Then
'UserList(UserIndex).Counters.Invisibilidad = IntervaloInvisible
'End If


If UserList(UserIndex).Counters.Incor < 60 Then 'cambio a 20 antes: 70
  UserList(UserIndex).Counters.Incor = UserList(UserIndex).Counters.Incor + 1
Else
  'Call SendData(ToIndex, UserIndex, 0, "E3")
  UserList(UserIndex).flags.Incor = False
  UserList(UserIndex).Counters.Incor = 0
  'Call SendData2(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
'UserList(UserIndex).Char.FX = 0

  'Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 16, UserList(UserIndex).Char.CharIndex & ",0")
End If
   Exit Sub
fallo:
Call LogError("EFECTOINCOR " & Err.number & " D: " & Err.Description)
         
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
On Error GoTo fallo
'pluto:7.0 añadimos bonus elfo oscuro
If MapInfo(UserList(UserIndex).Pos.Map).Invisible = 1 Then
UserList(UserIndex).Counters.Invisibilidad = IntervaloInvisible + UserList(UserIndex).BonusElfoOscuro
End If


If UserList(UserIndex).Counters.Invisibilidad < IntervaloInvisible + UserList(UserIndex).BonusElfoOscuro And MapInfo(UserList(UserIndex).Pos.Map).Pk = True Then
  UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad + 1
Else
  Call SendData(ToIndex, UserIndex, 0, "E3")
  UserList(UserIndex).flags.Oculto = 0
  UserList(UserIndex).Counters.Invisibilidad = 0
  UserList(UserIndex).flags.Invisible = 0
  Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 17, UserList(UserIndex).Char.CharIndex & ",0")
End If
   Exit Sub
fallo:
Call LogError("EFECTOINVISIBILIDAD " & Err.number & " D: " & Err.Description)
         
End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
On Error GoTo fallo
'pluto:2.14
If Npclist(NpcIndex).flags.PoderEspecial4 > 0 Then
Dim aa As Integer
aa = RandomNumber(1, 1000)
If aa > 998 Then Npclist(NpcIndex).Contadores.Paralisis = 0
End If

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
End If
Exit Sub
fallo:
Call LogError("EFECTOPARALISISNPC" & Err.number & " D: " & Err.Description)

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).Counters.Ceguera > 0 Then
    UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
Else
    Call SendData2(ToIndex, UserIndex, 0, 55)
     UserList(UserIndex).flags.Ceguera = 0
End If

'pluto:2.4.5
If UserList(UserIndex).Counters.Estupidez > 0 Then
    UserList(UserIndex).Counters.Estupidez = UserList(UserIndex).Counters.Estupidez - 1
Else
'[Tite] Añado condicion de que no este montado para que le quite la estupidez
    If UserList(UserIndex).flags.Montura = 0 Then
        Call SendData2(ToIndex, UserIndex, 0, 56)
        UserList(UserIndex).flags.Estupidez = 0
    End If
End If
    

Exit Sub
fallo:
Call LogError("EFECTOCEGUEESTU" & Err.number & " D: " & Err.Description)

End Sub
Public Sub EfectoProtec(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).Counters.Protec > 0 Then
    UserList(UserIndex).Counters.Protec = UserList(UserIndex).Counters.Protec - 1
Else
    'Call SendData2(ToIndex, UserIndex, 0, 55)
     UserList(UserIndex).flags.Protec = 0
     Call SendData(ToIndex, UserIndex, 0, "S2")
End If



Exit Sub
fallo:
Call LogError("EFECTOprotec" & Err.number & " D: " & Err.Description)

End Sub

Public Sub EfectoRon(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).Counters.Ron > 0 Then
    UserList(UserIndex).Counters.Ron = UserList(UserIndex).Counters.Ron - 1
Else
    'Call SendData2(ToIndex, UserIndex, 0, 55)
     UserList(UserIndex).flags.Ron = 0
     Call SendData(ToIndex, UserIndex, 0, "S2")
End If



Exit Sub
fallo:
Call LogError("EFECTO Ron" & Err.number & " D: " & Err.Description)

End Sub

Public Sub EfectoMorphUser(ByVal UserIndex As Integer)
On Error GoTo fallo


If UserList(UserIndex).Counters.Morph > 0 Then
    UserList(UserIndex).Counters.Morph = UserList(UserIndex).Counters.Morph - 1

Else
    '[gau]
    If UserList(UserIndex).flags.Morph > 0 Then Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).flags.Morph, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
    
    UserList(UserIndex).flags.Morph = 0
    'UserList(UserIndex).Flags.Angel = 0
    
End If
Exit Sub
fallo:
Call LogError("EFECTOMORPHUSER " & Err.number & " D: " & Err.Description)

End Sub
Public Sub EfectoMacrear(ByVal UserIndex As Integer, ByVal Macreanda As Byte)
On Error GoTo fallo
    
If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub


Select Case Macreanda
Case 1
Call SendData(ToPUserAreaCercana, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
Call DoTalar(UserIndex)
Exit Sub
Case 2
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_MINERO)
Call DoMineria(UserIndex)
Exit Sub
Case 3
Call DoDomar(UserIndex, UserList(UserIndex).flags.TargetNpc)
Exit Sub
Case 4
Call FundirMineral(UserIndex)
Exit Sub
Case 5
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_PESCAR)
Call DoPescar(UserIndex)
Exit Sub

End Select


Exit Sub
fallo:
End Sub
Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
On Error GoTo fallo
'nati: agrego el "Not UserList(UserIndex).flags.Morph = 214" para que el berserker sea imparalizable en su estado.
If UserList(UserIndex).Counters.Paralisis > 0 And UserList(UserIndex).flags.Angel = 0 And UserList(UserIndex).flags.Demonio = 0 And Not UserList(UserIndex).flags.Morph = 214 Then
UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
Else
    UserList(UserIndex).flags.Paralizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call SendData2(ToIndex, UserIndex, 0, 68)
End If
Exit Sub
fallo:
Call LogError("EFECTOPARALISISUSER" & Err.number & " D: " & Err.Description)

End Sub
Public Sub RecStamina(UserIndex As Integer, EnviarStats As Boolean, Intervalo As Integer)
On Error GoTo fallo
'pluto:2.18
If UserList(UserIndex).Pos.Map < 1 Or _
    UserList(UserIndex).Pos.Map > NumMaps Or _
    UserList(UserIndex).Pos.X < 2 Or _
    UserList(UserIndex).Pos.Y < 2 Or _
    UserList(UserIndex).Pos.X > 99 Or _
    UserList(UserIndex).Pos.Y > 99 Then Exit Sub
'--------------------------------------------------

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub
      
Dim massta As Integer
If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
   If UserList(UserIndex).Counters.STACounter < Intervalo Then
       UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
   Else
       UserList(UserIndex).Counters.STACounter = 0
       massta = CInt(RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)))
       UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + massta
       If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
        Call SendData(ToIndex, UserIndex, 0, "M1") 'descansas
        EnviarStats = True
       End If
End If
Exit Sub
fallo:
Call LogError("RECSTAMINA: UI:" & UserIndex & " Mapa: " & UserList(UserIndex).Pos.Map & " D: " & Err.Description)

End Sub

Public Sub EfectoVeneno(UserIndex As Integer, EnviarStats As Boolean)
On Error GoTo fallo
Dim n As Integer
If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub
If UserList(UserIndex).Counters.veneno < IntervaloVeneno Then
  UserList(UserIndex).Counters.veneno = UserList(UserIndex).Counters.veneno + 1
Else
  Call SendData(ToIndex, UserIndex, 0, "||Estas envenenado, si no te curas moriras." & "´" & FontTypeNames.FONTTYPE_VENENO)
  Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 30 & "," & 2)

  UserList(UserIndex).Counters.veneno = 0
  n = RandomNumber(1, 5) * UserList(UserIndex).flags.Envenenado
  UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - n
  If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
  EnviarStats = True
End If
Exit Sub
fallo:
Call LogError("EFECTOVENENO " & Err.number & " D: " & Err.Description)

End Sub

Public Sub DuracionPociones(UserIndex As Integer)
On Error GoTo fallo
'Controla la duracion de las pociones
'UserList(UserIndex).flags.DuracionEfecto = 100
If UserList(UserIndex).flags.DuracionEfecto > 0 Then
   UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1
   If UserList(UserIndex).flags.DuracionEfecto = 0 Then
        UserList(UserIndex).flags.TomoPocion = False
        UserList(UserIndex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
        UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
        Next
 Call SendData(ToIndex, UserIndex, 0, "S2")
   End If
End If
Exit Sub
fallo:
Call LogError("DURACION POCIONES " & Err.number & " D: " & Err.Description)

End Sub

Public Sub HambreYSed(UserIndex As Integer, fenviarAyS As Boolean)
On Error GoTo fallo
'pluto:6.8
If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub

Dim ff As Byte
ff = 10 - CInt(UserList(UserIndex).Stats.UserSkills(Supervivencia) / 20)

'Sed
If UserList(UserIndex).Stats.MinAGU > 0 Then
    If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
          UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
    Else
      
          UserList(UserIndex).Counters.AGUACounter = 0
          UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - ff
                            
          If UserList(UserIndex).Stats.MinAGU <= 0 Then
               UserList(UserIndex).Stats.MinAGU = 0
               UserList(UserIndex).flags.Sed = 1
          End If
                            
          fenviarAyS = True
                            
    End If
End If

'hambre
If UserList(UserIndex).Stats.MinHam > 0 Then
   If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
        UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
   Else
        UserList(UserIndex).Counters.COMCounter = 0
        
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - ff
        
        If UserList(UserIndex).Stats.MinHam <= 0 Then
               UserList(UserIndex).Stats.MinHam = 0
               UserList(UserIndex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If
Exit Sub
fallo:
Call LogError("HAMBREYSED" & Err.number & " D: " & Err.Description)

End Sub

Public Sub Sanar(UserIndex As Integer, EnviarStats As Boolean, Intervalo As Integer)
On Error GoTo fallo
If UserList(UserIndex).Pos.Map < 1 Or _
    UserList(UserIndex).Pos.Map > NumMaps Or _
    UserList(UserIndex).Pos.X < 2 Or _
    UserList(UserIndex).Pos.Y < 2 Or _
    UserList(UserIndex).Pos.X > 99 Or _
    UserList(UserIndex).Pos.Y > 99 Then Exit Sub

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub
       

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
   If UserList(UserIndex).Counters.HPCounter < Intervalo Then
      UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
   Else
      mashit = CInt(RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)))
                           
      UserList(UserIndex).Counters.HPCounter = 0
      UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + mashit
      If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
         Call SendData(ToIndex, UserIndex, 0, "M4")
         EnviarStats = True
      End If
End If
Exit Sub
fallo:
Call LogError("SANAR " & Err.number & " D: " & Err.Description)

End Sub

Public Sub CargaNpcsDat()
'Dim NpcFile As String
'
'NpcFile = DatPath & "NPCs.dat"
'ANpc = INICarga(NpcFile)
'Call INIConf(ANpc, 0, "", 0)
'
'NpcFile = DatPath & "NPCs-HOSTILES.dat"
'Anpc_host = INICarga(NpcFile)
'Call INIConf(Anpc_host, 0, "", 0)

Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
LeerNPCs.Abrir npcfile

npcfile = DatPath & "NPCs-HOSTILES.dat"
LeerNPCsHostiles.Abrir npcfile

End Sub

Public Sub DescargaNpcsDat()
'If ANpc <> 0 Then Call INIDescarga(ANpc)
'If Anpc_host <> 0 Then Call INIDescarga(Anpc_host)

End Sub
Sub PasarSegundo()
    Exit Sub
    Dim i As Integer
    For i = 1 To LastUser
        'Cerrar usuario
        If UserList(i).Counters.Saliendo Then
            UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
            If UserList(i).Counters.Salir <= 0 Then
                'If NumUsers <> 0 Then NumUsers = NumUsers - 1
                'Call aDos.RestarConexion(frmMain.Socket2(i).PeerAddress)
                'Call SendData(ToIndex, i, 0, "||Gracias por jugar AoDraG Online" & FONTTYPENAMES.FONTTYPE_INFO)
                Call SendData2(ToIndex, i, 0, 7)
                
                Call CloseUser(i)
                Exit Sub
'                Call CloseUser(i)
'                UserList(i).ConnID = -1: UserList(i).NumeroPaquetesPorMiliSec = 0
'                frmMain.Socket2(i).Disconnect
'                frmMain.Socket2(i).Cleanup
'                'Unload frmMain.Socket2(i)
'                Call ResetUserSlot(i)
'            Else
'                Call SendData(ToIndex, i, 0, "||En " & UserList(i).Counters.Salir & " segundos se cerrará el juego..." & FONTTYPENAMES.FONTTYPE_INFO)
            End If
        End If
 Next
End Sub
