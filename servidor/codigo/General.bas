Attribute VB_Name = "General"
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

Option Explicit

Global LeerNPCs As New clsIniReader

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'***************************************************
Dim CuerpoDesnudo As Integer
Select Case UserList(UserIndex).genero
    Case eGenero.Hombre
        Select Case UserList(UserIndex).raza
            Case eRaza.Humano
                CuerpoDesnudo = 21
            Case eRaza.Drow
                CuerpoDesnudo = 32
            Case eRaza.Elfo
                CuerpoDesnudo = 210
            Case eRaza.Gnomo
                CuerpoDesnudo = 222
            Case eRaza.Enano
                CuerpoDesnudo = 53
            Case eRaza.Orco
                CuerpoDesnudo = 24
            Case eRaza.NoMuerto
                CuerpoDesnudo = 32
        End Select
    Case eGenero.Mujer
        Select Case UserList(UserIndex).raza
            Case eRaza.Humano
                CuerpoDesnudo = 39
            Case eRaza.Drow
                CuerpoDesnudo = 40
            Case eRaza.Elfo
                CuerpoDesnudo = 259
            Case eRaza.Gnomo
                CuerpoDesnudo = 260
            Case eRaza.Enano
                CuerpoDesnudo = 60
            Case eRaza.Orco
                CuerpoDesnudo = 25
            Case eRaza.NoMuerto
                CuerpoDesnudo = 40
        End Select
End Select

If Mimetizado Then
    UserList(UserIndex).CharMimetizado.body = CuerpoDesnudo
Else
    UserList(UserIndex).Char.body = CuerpoDesnudo
End If

UserList(UserIndex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Boolean)
'b ahora es boolean,
'b=true bloquea el tile en (x,y)
'b=false desbloquea el tile en (x,y)
'toMap = true -> Envia los datos a todo el mapa
'toMap = false -> Envia los datos al user
'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s

If toMap Then
    Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
Else
    Call WriteBlockPosition(sndIndex, X, Y, b)
End If

End Sub


Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If Map > 0 And Map < NumMaps + 1 And X > 0 And X < XMaxMapSize + 1 And Y > 0 And Y < YMaxMapSize + 1 Then
    If ((MapData(Map, X, Y).Graphic(1) >= 1505 And MapData(Map, X, Y).Graphic(1) <= 1520) Or _
    (MapData(Map, X, Y).Graphic(1) >= 5665 And MapData(Map, X, Y).Graphic(1) <= 5680) Or _
    (MapData(Map, X, Y).Graphic(1) >= 13547 And MapData(Map, X, Y).Graphic(1) <= 13562)) And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function

Private Function HayLava(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'***************************************************
If Map > 0 And Map < NumMaps + 1 And X > 0 And X < XMaxMapSize + 1 And Y > 0 And Y < YMaxMapSize + 1 Then
    If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
        HayLava = True
    Else
        HayLava = False
    End If
Else
  HayLava = False
End If

End Function

Sub LimpiarMundo()
'***************************************************
'Author: Unknown
'Last Modification: 05/09/2012 - ^[GS]^
'***************************************************
On Error GoTo errHandler
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpiando Mundo...", FontTypeNames.FONTTYPE_INFOBOLD))
    
    If aLimpiarMundo.CantItems > 0 Then
        Call aLimpiarMundo.EraseAllItems
    End If
    
    Call SecurityIp.IpSecurityMantenimientoLista
    
    Exit Sub

errHandler:
    Call LogError("Error producido en el sub LimpiarMundo: " & Err.Description)
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
Dim k As Long
Dim npcNames() As String

ReDim npcNames(1 To UBound(SpawnList)) As String

For k = 1 To UBound(SpawnList)
    npcNames(k) = SpawnList(k).NpcName
Next k

Call WriteSpawnList(UserIndex, npcNames())

End Sub

Sub Main()
On Error Resume Next
    Dim f As Date
    Dim i As Byte
    
    ChDir App.Path
    ChDrive App.Path
    
    Call LoadMotd
    Call BanIpCargar
    Call BanHD_load
    
    Prision.Map = 17
    Libertad.Map = 17
    
    Prision.X = 43
    Prision.Y = 52
    Libertad.X = 43
    Libertad.Y = 65
    
    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")
    
    IniPath = App.Path & "\"
    DatPath = App.Path & "\Dat\"
       
    LevelSkill(1).LevelValue = 2
    LevelSkill(2).LevelValue = 4
    LevelSkill(3).LevelValue = 6
    LevelSkill(4).LevelValue = 8
    LevelSkill(5).LevelValue = 10
    LevelSkill(6).LevelValue = 12
    LevelSkill(7).LevelValue = 14
    LevelSkill(8).LevelValue = 16
    LevelSkill(9).LevelValue = 18
    LevelSkill(10).LevelValue = 20
    LevelSkill(11).LevelValue = 22
    LevelSkill(12).LevelValue = 24
    LevelSkill(13).LevelValue = 26
    LevelSkill(14).LevelValue = 28
    LevelSkill(15).LevelValue = 30
    LevelSkill(16).LevelValue = 32
    LevelSkill(17).LevelValue = 34
    LevelSkill(18).LevelValue = 36
    LevelSkill(19).LevelValue = 38
    LevelSkill(20).LevelValue = 40
    LevelSkill(21).LevelValue = 42
    LevelSkill(22).LevelValue = 44
    LevelSkill(23).LevelValue = 46
    LevelSkill(24).LevelValue = 48
    LevelSkill(25).LevelValue = 50
    LevelSkill(26).LevelValue = 52
    LevelSkill(27).LevelValue = 54
    LevelSkill(28).LevelValue = 56
    LevelSkill(29).LevelValue = 58
    LevelSkill(30).LevelValue = 60
    LevelSkill(31).LevelValue = 62
    LevelSkill(32).LevelValue = 64
    LevelSkill(33).LevelValue = 66
    LevelSkill(34).LevelValue = 68
    LevelSkill(35).LevelValue = 70
    LevelSkill(36).LevelValue = 72
    LevelSkill(37).LevelValue = 74
    LevelSkill(38).LevelValue = 76
    LevelSkill(39).LevelValue = 78
    LevelSkill(40).LevelValue = 80
    LevelSkill(41).LevelValue = 82
    LevelSkill(42).LevelValue = 84
    LevelSkill(43).LevelValue = 86
    LevelSkill(44).LevelValue = 88
    LevelSkill(45).LevelValue = 90
    LevelSkill(46).LevelValue = 92
    LevelSkill(47).LevelValue = 94
    LevelSkill(48).LevelValue = 96
    LevelSkill(49).LevelValue = 98
    LevelSkill(50).LevelValue = 100
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    ListaRazas(eRaza.Orco) = "Orco"
    ListaRazas(eRaza.NoMuerto) = "No-Muerto"
    
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.talar) = "Talar arboles"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
    SkillsNames(eSkill.Wrestling) = "Wrestling"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Energia) = "Energia"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    
    
    frmCargando.Show
    
    '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    
    MaxUsers = 0
    
    If Not LoadSini Then
        MsgBox "Error al cargar el archivo Server.ini"
        Exit Sub
    End If
    
    
    frmCargando.Label1(2).Caption = "Cargando Base de Datos"
    'Irongete: Conecto a SQL
    Call MySQL_Connect
    
    '19/11/2015 Irongete: Todos los personajes empiezan deslogeados al arrancar el servidor
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("UPDATE personaje SET logged = '0'")
    
    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    IniPath = App.Path & "\"
    CharPath = App.Path & "\Charfile\"
    
    Set aLimpiarMundo = New clsLimpiarMundo
    
    ' Initialize classes
    #If SocketType = 1 Then
        Set WSAPISock2Usr = New Collection
    #End If
    
    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    DoEvents
    
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    
    '05/11/2015 Cargar los clanes
    Call CargarClanes
    
    
    Call CargarSpawnList
    Call CargarForbidenWords
    
    centinelaActivado = True
    
    'frmCargando.Label1(2).Caption = "Conectando a base de datos"
    'Call CargarDB
    
    '*************************************************
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat
    '*************************************************
    
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    'Call LoadOBJData
    Call LoadOBJData
        
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos
        
        
    frmCargando.Label1(2).Caption = "Cargando Objetos de Herrería"
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    
    frmCargando.Label1(2).Caption = "Cargando Objetos de Carpintería"
    Call LoadObjCarpintero
    
    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance    '4/01/08 Pablo ToxicWaste
    
    frmCargando.Label1(2).Caption = "Cargando Ranking"
    Call CargarRank
    
    frmCargando.Label1(2).Caption = "Cargando Castillos"
    Call CargarCastillos
    
    If BootDelBackUp Then
        
        frmCargando.Label1(2).Caption = "Cargando BackUp"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData
    End If
    
    
    'Cargar las zonas
    Call cargar_zonas_sql
    
    'Cargar los efectos
    Call cargar_efectos_sql
    

    ' Internet IP
    'frmCargando.Label1(2).Caption = "Buscando IP en Internet..." ' GSZ
    'frmMain.txtIP.Caption = frmMain.Inet1.OpenURL("http://ip1.dynupdate.no-ip.com:8245/")  'ReadField(1, ReadField(i, frmMain.Inet1.OpenURL("http://aodrag.es/client/sinfo.txt"), Asc("|")), Asc(":")) '
    'frmMain.txtPort.Caption = Puerto
    DoEvents
    If frmMain.txtIP.Caption = vbNullString Then frmMain.txtIP.Caption = "N/A"

    frmMain.Inet1.OpenURL ("http://www.aodrag.es/heartbeat.php")
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Dim LoopC As Integer
    
    'Resetea las conexiones de los usuarios
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    With frmMain
        .AutoSave.Enabled = True
        .tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .TimerMinuto.Enabled = True
        .Auditoria.Enabled = True
        .KillLog.Enabled = True
        .NPC_AI.Enabled = True
        .npcataca.Enabled = True
    End With
    
    Call SocketConfig
    
    If frmMain.Visible Then frmMain.txStatus.Text = "Escuchando conexiones entrantes ..."
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    frmCargando.Label1(2).Caption = "Cargando Clima"
    Call SortearHorario 'Lorwik> Lo coloco aqui o no funciona
    
    Unload frmCargando
       
    'Log
    Dim n As Integer
    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #n
    
    'Ocultar
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If
    
    tInicioServer = GetTickCount() And &H7FFFFFFF
    Call InicializaEstadisticas
    
    
    

End Sub

Private Sub SocketConfig()
'*****************************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'Sets socket config.
'*****************************************************************
On Error Resume Next

    Call SecurityIp.InitIpTables(1000)
    
#If SocketType = 1 Then
    
    If LastSockListen >= 0 Then Call apiclosesocket(LastSockListen) 'Cierra el socket de escucha
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(iniPuerto, hWndMsg, "")
    If SockListen <> -1 Then
        Call WriteVar(IniPath & "Server.ini", "CONEXION", "LastSockListen", SockListen) ' Guarda el socket escuchando
    Else
        MsgBox "Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly
    End If
    
    DoEvents
#ElseIf SocketType = 2 Then 'WINSOCKETS
    frmMain.wskListen.Close
    frmMain.wskListen.LocalPort = Puerto
    frmMain.wskListen.listen
    Dim i As Integer
    For i = 1 To MaxUsers
        Load frmMain.wskClient(i)
    Next i
#End If
    
    If frmMain.Visible Then frmMain.txStatus.Text = "Escuchando conexiones entrantes ..."
    
End Sub


Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = LenB(dir$(File, FileType)) <> 0
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()

frmMain.CantUsuarios.Caption = NumUsers

End Sub


Public Sub LogCriticEvent(desc As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

errHandler:

End Sub

Public Sub LogEjercitoReal(desc As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, desc
Close #nfile

Exit Sub

errHandler:

End Sub

Public Sub LogEjercitoCaos(desc As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, desc
Close #nfile

Exit Sub

errHandler:

End Sub


Public Sub LogIndex(ByVal Index As Integer, ByVal desc As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

errHandler:

End Sub


Public Sub LogError(desc As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

errHandler:

End Sub

Public Sub LogStatic(desc As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

errHandler:

End Sub

Public Sub LogTarea(desc As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

errHandler:


End Sub


Public Sub LogClanes(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & str
Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\IP.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & str
Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & str
Close #nfile

End Sub



Public Sub LogGM(nombre As String, texto As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
Open App.Path & "\logs\" & nombre & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Close #nfile

Exit Sub

errHandler:

End Sub

Public Sub LogAsesinato(texto As String)
On Error GoTo errHandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Close #nfile

Exit Sub

errHandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errHandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errHandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CH.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Close #nfile

Exit Sub

errHandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errHandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo errHandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Print #nfile, ""
Close #nfile

Exit Sub

errHandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If LenB(Arg) = 0 Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.Text = "Reiniciando."
    
    Dim LoopC As Long
      
    #If SocketType = 1 Then
    
        'Cierra el socket de escucha
        If SockListen >= 0 Then Call apiclosesocket(SockListen)
        
        'Inicia el socket de escucha
        SockListen = ListenForConnect(Puerto, hWndMsg, "")
    #ElseIf SocketType = 2 Then
        frmMain.wskListen.Close
        frmMain.wskListen.LocalPort = Puerto
        frmMain.wskListen.listen
    #End If

    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next

    'Initialize statistics!!
    Call Statistics.Initialize
    
    For LoopC = 1 To UBound(UserList())
        Set UserList(LoopC).incomingData = Nothing
        Set UserList(LoopC).outgoingData = Nothing
    Next LoopC
    
    ReDim UserList(1 To MaxUsers) As User
    
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    
    LastUser = 0
    NumUsers = 0
    
    Call FreeNPCs
    Call FreeCharIndexes
    
    If Not LoadSini Then
        MsgBox "Error al cargar el archivo Server.ini"
        Exit Sub
    End If
    
    Call LoadOBJData
    
    Call LoadMapData
    
    Call CargarHechizos
    
    If frmMain.Visible Then frmMain.txStatus.Text = "Escuchando conexiones entrantes ..."
    
    'Log it
    Dim n As Integer
    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & time & " servidor reiniciado."
    Close #n
    
    'Ocultar
    
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If

  
End Sub

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    
    If MapInfo(UserList(UserIndex).Pos.Map).zona <> "DUNGEON" Then
        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 1 And _
           MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 2 And _
           MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If NPCList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           NPCList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           NPCList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If NPCList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
    
    Dim modifi As Integer
    
    If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
        UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
    Else
        If MapInfo(UserList(UserIndex).Pos.Map).Terreno = Nieve Then
            Call WriteMultiMessage(UserIndex, eMessages.Frio)
            modifi = Porcentaje(VidaMaxima(UserIndex), 5)
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
            
            If UserList(UserIndex).Stats.MinHP < 1 Then
                Call WriteMultiMessage(UserIndex, eMessages.Murio)
                UserList(UserIndex).Stats.MinHP = 0
                Call UserDie(UserIndex)
            End If
            
            Call WriteUpdateHP(UserIndex)
        Else
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)
            Call QuitarSta(UserIndex, modifi)
            Call WriteUpdateSta(UserIndex)
        End If
        
        UserList(UserIndex).Counters.Frio = 0
    End If
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'If user is standing on lava, take health points from him
'***************************************************
    If UserList(UserIndex).Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
        UserList(UserIndex).Counters.Lava = UserList(UserIndex).Counters.Lava + 1
    Else
        If HayLava(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
            Call WriteMultiMessage(UserIndex, eMessages.Quema)
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Porcentaje(VidaMaxima(UserIndex), 5)
            
            If UserList(UserIndex).Stats.MinHP < 1 Then
                 Call WriteMultiMessage(UserIndex, eMessages.Murio)
                UserList(UserIndex).Stats.MinHP = 0
                Call UserDie(UserIndex)
            End If
            
            Call WriteUpdateHP(UserIndex)
        End If
        
        UserList(UserIndex).Counters.Lava = 0
    End If
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
'******************************************************
'Author: Unknown
'Last Update: 04/11/2008 (NicoNZ)
'
'******************************************************
    Dim Barco As ObjData
    
    With UserList(UserIndex)
        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            Call WriteMultiMessage(UserIndex, eMessages.Mimetico)
            
            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    If .Faccion.ArmadaReal = 1 Or .Faccion.Legion = 1 Then
                        .Char.body = iFragataReal
                    ElseIf .Faccion.FuerzasCaos = 1 Then
                        .Char.body = iFragataCaos
                    Else
                        Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
                        If criminal(UserIndex) Then
                            If Barco.Ropaje = iBarca Then .Char.body = iBarcaPk
                            If Barco.Ropaje = iGalera Then .Char.body = iGaleraPk
                            If Barco.Ropaje = iGaleon Then .Char.body = iGaleonPk
                        Else
                            If Barco.Ropaje = iBarca Then .Char.body = iBarcaCiuda
                            If Barco.Ropaje = iGalera Then .Char.body = iGaleraCiuda
                            If Barco.Ropaje = iGaleon Then .Char.body = iGaleonCiuda
                        End If
                    End If
                Else
                    .Char.body = iFragataFantasmal
                End If
                
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            Else
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            End If
            
            With .Char
                Call ChangeUserChar(UserIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
            End With
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
        End If
    End With
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Invisibilidad < IntervaloInvisible Then
    UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad + 1
Else
    UserList(UserIndex).Counters.Invisibilidad = 0
    UserList(UserIndex).flags.invisible = 0
    If UserList(UserIndex).flags.Oculto = 0 Then
        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    End If
End If

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If NPCList(NpcIndex).Contadores.Paralisis > 0 Then
    NPCList(NpcIndex).Contadores.Paralisis = NPCList(NpcIndex).Contadores.Paralisis - 1
Else
    NPCList(NpcIndex).flags.Paralizado = 0
    NPCList(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Ceguera > 0 Then
    UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
Else
    If UserList(UserIndex).flags.Ceguera = 1 Then
        UserList(UserIndex).flags.Ceguera = 0
        Call WriteBlindNoMore(UserIndex)
    End If
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call WriteDumbNoMore(UserIndex)
    End If

End If


End Sub

Public Sub EfectoMorphUser(ByVal UserIndex As Integer)
On Error GoTo fallo

    If UserList(UserIndex).Counters.Morph > 0 Then
        UserList(UserIndex).Counters.Morph = UserList(UserIndex).Counters.Morph - 1
    Else
        '[gau]
        If UserList(UserIndex).flags.Morph > 0 Then Call ChangeUserChar(UserIndex, UserList(UserIndex).flags.Morph, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        UserList(UserIndex).flags.Morph = 0
    End If
    Exit Sub
fallo:
Call LogError("EFECTOMORPHUSER " & Err.Number & " D: " & Err.Description)

End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Paralisis > 0 Then
    UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
Else
    UserList(UserIndex).flags.Paralizado = 0
    UserList(UserIndex).flags.Inmovilizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call WriteParalizeOK(UserIndex)
End If

End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal intervalo As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub


Dim massta As Integer
If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
    If UserList(UserIndex).Counters.STACounter < intervalo Then
        UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
    Else
        EnviarStats = True
        UserList(UserIndex).Counters.STACounter = 0
        If UserList(UserIndex).flags.Desnudo Or _
        UserList(UserIndex).flags.Makro <> 0 Then Exit Sub 'Desnudo y trabajando no sube energia. (ToxicWaste)
       
        massta = RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + massta
        If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
        End If
    End If
End If

End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)
Dim n As Integer

If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
  UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
Else
  Call WriteConsoleMsg(UserIndex, "Estás envenenado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
  UserList(UserIndex).Counters.Veneno = 0
  n = RandomNumber(1, 5)
  UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - n
  If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
  Call WriteUpdateHP(UserIndex)
End If

End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)
'***************************************************
'Author: ??????
'Last Modification: 11/27/09 (Budi)
'Cuando se pierde el efecto de la poción updatea fz y agi (No me gusta que ambos atributos aunque se haya modificado solo uno, pero bueno :p)
'***************************************************
    With UserList(UserIndex)
        'Controla la duracion de las pociones
        If .flags.DuracionEfecto > 0 Then
           .flags.DuracionEfecto = .flags.DuracionEfecto - 1
           If .flags.DuracionEfecto = 0 Then
                .flags.TomoPocion = False
                .flags.TipoPocion = 0
                'volvemos los atributos al estado normal
                Dim LoopX As Integer
                
                For LoopX = 1 To NUMATRIBUTOS
                    .Stats.UserAtributos(LoopX) = .Stats.UserAtributosBackUP(LoopX)
                Next LoopX
                
                Call WriteUpdateStrenghtAndDexterity(UserIndex)
           End If
        End If
    End With

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)

If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Exit Sub

'Sed
If UserList(UserIndex).Stats.MinAGU > 0 Then
    If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
        UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
    Else
        UserList(UserIndex).Counters.AGUACounter = 0
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
        
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
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10
        If UserList(UserIndex).Stats.MinHam <= 0 Then
               UserList(UserIndex).Stats.MinHam = 0
               UserList(UserIndex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal intervalo As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(UserIndex).Stats.MinHP < VidaMaxima(UserIndex) Then
    If UserList(UserIndex).Counters.HPCounter < intervalo Then
        UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
    Else
        mashit = Porcentaje(UserList(UserIndex).Stats.MaxSta, 4)
        
        UserList(UserIndex).Counters.HPCounter = 0
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + mashit
        If UserList(UserIndex).Stats.MinHP > VidaMaxima(UserIndex) Then UserList(UserIndex).Stats.MinHP = VidaMaxima(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFOBOLD)
        EnviarStats = True
    End If
End If

End Sub

Public Sub CargaNpcsDat()
    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    TotalNPCDat = GetVar(npcfile, "INIT", "NumNPCs")
    Call LeerNPCs.Initialize(npcfile)
End Sub

Sub PasarSegundo()
On Error GoTo errHandler
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
                Call WriteConsoleMsg(i, "Cerrando en " & UserList(i).Counters.Salir - 1, FontTypeNames.FONTTYPE_INFO)
                UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
                If UserList(i).Counters.Salir <= 0 Then
                    Call WriteConsoleMsg(i, "Gracias por jugar AODrag", FontTypeNames.FONTTYPE_INFO)
                    Call WriteDisconnect(i) '
                    Call FlushBuffer(i)
                    
                    Call CloseSocket(i)
                End If
            End If
        End If
        If UserList(i).flags.GanoDueloSet = True Then
            If UserList(i).flags.TimeDueloSet > 0 Then
                UserList(i).flags.TimeDueloSet = UserList(i).flags.TimeDueloSet - 1
        Else
                Call WarpUserChar(i, 1, 41, 88, True)
                UserList(i).flags.GanoDueloSet = False
                Call SaveUser(i)
            End If
        End If
    Next i
Exit Sub

errHandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.Description & " - " & Err.Number & " - UserIndex: " & i)
    Resume Next
End Sub
 
Public Function ReiniciarAutoUpdate() As Double

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp

    'commit experiencias
    'Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
    HaciendoBackup = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i)
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    HaciendoBackup = False
End Sub


Sub InicializaEstadisticas()
Dim Ta As Long
Ta = GetTickCount() And &H7FFFFFFF

Call EstadisticasWeb.Inicializa(frmMain.hWnd)
Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub

Public Sub FreeNPCs()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all NPC Indexes
'***************************************************
    Dim LoopC As Long
    
    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        NPCList(LoopC).flags.Active = False
    Next LoopC
End Sub

Public Sub FreeCharIndexes()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all char indexes
'***************************************************
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
End Sub

Public Sub DropToUser(ByVal UserIndex As Integer, ByVal tIndex As Integer, ByVal userSlot As Byte, ByVal Amount As Integer)
'***************************************************
'Autor: maTih.-
'Last Modification: -
'
'***************************************************

Dim targetObj   As Obj
Dim targetUser  As Integer
Dim targetMSG   As String

       With UserList(UserIndex)
            'Save the userIndex.
            targetUser = tIndex
       
            'It is a valid user?.
                If UserList(targetUser).ConnID <> -1 Then
                    targetObj.Amount = Amount
                    targetObj.ObjIndex = .Invent.Object(userSlot).ObjIndex
           
                    'I give the object to another user.
                    MeterItemEnInventario targetUser, targetObj
          
                    'Remove the object to my userIndex.
                    QuitarUserInvItem UserIndex, userSlot, Amount
                    
                    'Update my inventory.
                    UpdateUserInv False, UserIndex, userSlot
                    
                    'Avistage a users.
                    If Amount <> 1 Then
                       targetMSG = "Le has arrojado " & Amount & " - " & ObjData(targetObj.ObjIndex).Name
                    Else
                       targetMSG = "Le has arrojado tu " & ObjData(targetObj.ObjIndex).Name
                    End If
                    
                    WriteConsoleMsg UserIndex, targetMSG & " ah " & UserList(targetUser).Name & "!", FontTypeNames.FONTTYPE_CITIZEN
                    
                    targetMSG = UserList(tIndex).Name
                    
                    'Prepare message to other user.
                    If Amount <> 1 Then
                       targetMSG = targetMSG & " Te ha arrojado " & Amount & " - " & ObjData(targetObj.ObjIndex).Name
                    Else
                       targetMSG = targetMSG & " Te ha arrojado su " & ObjData(targetObj.ObjIndex).Name
                    End If
                    
                    WriteConsoleMsg tIndex, targetMSG, FontTypeNames.FONTTYPE_CITIZEN
                    
                    Exit Sub
                End If
        End With
End Sub

Sub DropToNPC(ByVal UserIndex As Integer, ByVal tNPC As Integer, ByVal userSlot As Byte, ByVal Amount As Integer)
'***************************************************
'Autor: maTih.-
'Last Modification: -
'
'***************************************************

On Error Resume Next

With UserList(UserIndex)

Dim sellOk     As Boolean

     'ITs banquero?
     If NPCList(tNPC).NPCType = eNPCType.Banquero Then
        'Deposit the obj.
        UserDejaObj UserIndex, userSlot, Amount
        'Avistage user.
        WriteConsoleMsg UserIndex, "Has depositado " & Amount & " - " & ObjData(.Invent.Object(userSlot).ObjIndex).Name, FontTypeNames.FONTTYPE_CITIZEN
        'Update the inventory
        UpdateUserInv False, UserIndex, userSlot
        Exit Sub
     End If

     'NPC is a merchant?
     If NPCList(tNPC).Comercia <> 0 Then
        Comercio eModoComercio.Venta, UserIndex, tNPC, userSlot, Amount
     End If
        
End With

End Sub

'Anti-Cheats Lac(Loopzer Anti-Cheats)
Public Sub ResetearLac(UserIndex As Integer)
With UserList(UserIndex).Lac
    .LCaminar.init Lac_Camina
    .LPociones.init Lac_Pociones
    .LUsar.init Lac_Usar
    .LPegar.init Lac_Pegar
    .LLanzar.init Lac_Lanzar
    .LTirar.init Lac_Tirar
End With
 
End Sub

Public Sub CargaLac(UserIndex As Integer)
With UserList(UserIndex).Lac
    Set .LCaminar = New Cls_InterGTC
    Set .LLanzar = New Cls_InterGTC
    Set .LPegar = New Cls_InterGTC
    Set .LPociones = New Cls_InterGTC
    Set .LTirar = New Cls_InterGTC
    Set .LUsar = New Cls_InterGTC
 
    .LCaminar.init Lac_Camina
    .LPociones.init Lac_Pociones
    .LUsar.init Lac_Usar
    .LPegar.init Lac_Pegar
    .LLanzar.init Lac_Lanzar
    .LTirar.init Lac_Tirar
End With
 
End Sub

Public Sub DescargaLac(UserIndex As Integer)
Exit Sub
With UserList(UserIndex).Lac
    Set .LCaminar = Nothing
    Set .LLanzar = Nothing
    Set .LPegar = Nothing
    Set .LPociones = Nothing
    Set .LTirar = Nothing
    Set .LUsar = Nothing
End With
End Sub

'/Anti-Cheats Lac(Loopzer Anti-Cheats)

Sub CambiarCabeza(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 14/03/2007
'Elije una cabeza para el usuario y le da un body
'*************************************************
Dim NewHead As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte
UserGenero = UserList(UserIndex).genero
UserRaza = UserList(UserIndex).raza
Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(1, 40)
            Case eRaza.Elfo
                NewHead = RandomNumber(101, 112)
            Case eRaza.Drow
                NewHead = RandomNumber(200, 210)
            Case eRaza.Enano
                NewHead = RandomNumber(300, 306)
            Case eRaza.Gnomo
                NewHead = RandomNumber(401, 406)
            Case eRaza.Orco
                NewHead = RandomNumber(130, 130)
            Case eRaza.NoMuerto
                NewHead = RandomNumber(507, 507)
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(70, 79)
            Case eRaza.Elfo
                NewHead = RandomNumber(170, 178)
            Case eRaza.Drow
                NewHead = RandomNumber(270, 278)
            Case eRaza.Gnomo
                NewHead = RandomNumber(370, 372)
            Case eRaza.Enano
                NewHead = RandomNumber(470, 476)
            Case eRaza.Orco
                NewHead = RandomNumber(131, 131)
            Case eRaza.NoMuerto
                NewHead = RandomNumber(506, 506)
        End Select
End Select
UserList(UserIndex).OrigChar.Head = NewHead
Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
End Sub

Public Sub AdvertirUsuario(ByVal UserIndex As Integer, ByVal UserAdvertido As Integer)

'Con esto evitamos lo siguiente : _
*Que el comando no lo puedan utilizar jerarkias menores a semi dioses _
*Que no se puedan advertir a personajes con rango mayor a semi dioses _
*Que no se pueda advertir a personajes offlines. <span style="color: #e1e1e1;">(Evitar</span> Bug de chars nulos) ;-)
    If UserList(UserIndex).flags.Privilegios < PlayerType.SemiDios Then Exit Sub
    If UserList(UserAdvertido).flags.Privilegios > PlayerType.SemiDios Then Exit Sub
    If UserAdvertido <= 0 Then Exit Sub
     
    If UserList(UserAdvertido).flags.Ban = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes advertir a " & UserList(UserAdvertido).Name & ", ya que se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call OtorgarAdvertencia(UserAdvertido, UserIndex)
        Call Encarcelar(UserAdvertido, 5, UserList(UserIndex).Name)
        Call WriteConsoleMsg(UserIndex, "Has aplicado la advertencia sobre " & UserList(UserAdvertido).Name & ".", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub
 
Public Sub OtorgarAdvertencia(ByVal UserIndex As Integer, ByVal Advertidor As Integer)
    Dim MaxAdvertencias As Byte
    UserList(UserIndex).flags.Advertencias = UserList(UserIndex).flags.Advertencias + 1
     
    MaxAdvertencias = 5 'reemplazar el número 5 por la cantidad máxima de advertencias.
    
    If UserList(UserIndex).flags.Advertencias >= MaxAdvertencias Then
        Call WriteShowMessageBox(UserIndex, "Has llegado al tope de advertencias y has sido expulsado del servidor de forma permanente. La pena fué aplicada por el administrador " & UserList(Advertidor).Name & "")
        'Call BanCharacter(Advertidor, UserIndex, "Acumulación de Advertencias")
        UserList(UserIndex).flags.Ban = 1
        Call CloseSocket(UserIndex)
    Else
        Call WriteShowMessageBox(UserIndex, "El administrador " & UserList(Advertidor).Name & " te ha advertido. Tienes " & UserList(UserIndex).flags.Advertencias & " advertencias, recuerda que a las " & MaxAdvertencias & " serás expulsado del servidor de forma permanente.")
    End If
End Sub

Public Function Tilde(data As String) As String
    Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
    With RichTextBox
        If Len(.Text) > 10000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf)
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
            .Text = ""
        End If
       
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
       
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
       
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
       
    End With
End Sub
