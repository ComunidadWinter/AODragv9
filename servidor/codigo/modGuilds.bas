Attribute VB_Name = "modGuilds"
'**************************************************************
' modGuilds.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Implemented by Mariano Barrou (El Oso)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

'guilds nueva version. Hecho por el oso, eliminando los problemas
'de sincronizacion con los datos en el HD... entre varios otros
'��

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private GUILDINFOFILE   As String
'archivo .\guilds\guildinfo.ini o similar

Private Const MAX_GUILDS As Integer = 1000
'cantidad maxima de guilds en el servidor

Public CANTIDADDECLANES As Integer
'cantidad actual de clanes en el servidor

Private guilds(1 To MAX_GUILDS) As clsClan
'array global de guilds, se indexa por userlist().guildindex

Private Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES As Byte = 10
'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Private Const MAXANTIFACCION As Byte = 5
'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

'13/02/2016 Lorwik: Limite de miembros del clan
Private Const MIEMBROSMAXIMOS As Byte = 30
'cantidad maxima de miembros en un clan.


Public Enum ALINEACION_GUILD
    ALINEACION_LEGION = 1
    ALINEACION_CRIMINAL = 2
    ALINEACION_NEUTRO = 3
    ALINEACION_CIUDA = 4
    ALINEACION_ARMADA = 5
    ALINEACION_MASTER = 6
End Enum
'alineaciones permitidas

Public Enum SONIDOS_GUILD
    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45
End Enum
'numero de .wav del cliente

Public Enum RELACIONES_GUILD
    GUERRA = -1
    PAZ = 0
    ALIADOS = 1
End Enum
'estado entre clanes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'05/11/2015 Irongete: Rework de esta funci�n para cargar los clanes desde SQL
' Antigua LoadGuildsDB()

Public Sub CargarClanes()
    On Error GoTo errHandler

    Dim i           As Integer
    Dim TempStr     As String
    Dim Alin        As ALINEACION_GUILD
        
    
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT nombre FROM clan")
    
     CANTIDADDECLANES = CInt(RS.RecordCount)
    
    i = 1
    If RS.RecordCount > 0 Then
        While Not RS.EOF
             Set guilds(i) = New clsClan
             Call guilds(i).Inicializar(CStr(RS!nombre), i, 0)
             i = i + 1
            RS.MoveNext
        Wend
    End If
  
    Exit Sub
    
errHandler:
    Debug.Print "Error en CargarClanes(): " & Err.Description
    
End Sub

Public Function m_ConectarMiembroAClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
Dim NuevoL  As Boolean
Dim NuevaA  As Boolean
Dim News    As String

    If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
    If m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
        Call guilds(GuildIndex).ConectarMiembro(UserIndex)
        UserList(UserIndex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
    Else
        m_ConectarMiembroAClan = m_ValidarPermanencia(UserIndex, True, NuevaA, NuevoL)
        If NuevoL Then News = "El clan tiene nuevo l�der."
        If NuevaA Then News = News & "El clan tiene nueva alineaci�n."
        If NuevoL Or NuevaA Then Call guilds(GuildIndex).SetGuildNews(News)
    End If

End Function


Public Function m_ValidarPermanencia(ByVal UserIndex As Integer, ByVal SumaAntifaccion As Boolean, ByRef CambioAlineacion As Boolean, ByRef CambioLider As Boolean) As Boolean
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 25/03/2009
'25/03/2009: ZaMa - Desequipo los items faccionarios que tenga el funda al abandonar la faccion
'***************************************************

Dim GuildIndex  As Integer
Dim ML()        As String
Dim M           As String
Dim UI          As Integer
Dim Sale        As Boolean
Dim i           As Long
Dim totalMembers As Integer

    m_ValidarPermanencia = True
    GuildIndex = UserList(UserIndex).GuildIndex
    If GuildIndex > CANTIDADDECLANES And GuildIndex <= 0 Then Exit Function
    
    If Not m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
    
        Call LogClanes(UserList(UserIndex).Name & " de " & guilds(GuildIndex).GuildName & " es expulsado en validar permanencia")
    
        m_ValidarPermanencia = False
        If SumaAntifaccion Then guilds(GuildIndex).PuntosAntifaccion = guilds(GuildIndex).PuntosAntifaccion + 1
        
        CambioAlineacion = (m_EsGuildFounder(UserList(UserIndex).Name, GuildIndex) Or guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION)
        
        Call LogClanes(UserList(UserIndex).Name & " de " & guilds(GuildIndex).GuildName & IIf(CambioAlineacion, " SI ", " NO ") & "provoca cambio de alinaecion. MAXANT:" & (guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION) & ", GUILDFOU:" & m_EsGuildFounder(UserList(UserIndex).Name, GuildIndex))
        
        If CambioAlineacion Then
            'aca tenemos un problema, el fundador acaba de cambiar el rumbo del clan o nos zarpamos de antifacciones
            'Tenemos que resetear el lider, revisar si el lider permanece y si no asignarle liderazgo al fundador

            Call guilds(GuildIndex).CambiarAlineacion(ALINEACION_NEUTRO)
            guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION
            'para la nueva alineacion, hay que revisar a todos los Pjs!

            'uso GetMemberList y no los iteradores pq voy a rajar gente y puedo alterar
            'internamente al iterador en el proceso
            CambioLider = False
            ML = guilds(GuildIndex).GetMemberList()
            totalMembers = UBound(ML)
            
            'El user en ML(0) es el funda, no tiene setnido verificar su permanencia!
            For i = 1 To totalMembers
                M = ML(i)
                
                'vamos a violar un poco de capas..
                UI = NameIndex(M)
                If UI > 0 Then
                    Sale = Not m_EstadoPermiteEntrar(UI, GuildIndex)
                Else
                    Sale = Not m_EstadoPermiteEntrarChar(M, GuildIndex)
                End If

                If Sale Then
                    If m_EsGuildFounder(M, GuildIndex) Then 'hay que sacarlo de las armadas
                     
                        If UI > 0 Then
                            If UserList(UI).Faccion.ArmadaReal <> 0 Then
                                Call ExpulsarFaccionReal(UI)
                            ElseIf UserList(UI).Faccion.FuerzasCaos <> 0 Then
                                Call ExpulsarFaccionCaos(UI)
                            End If
                           UserList(UI).Faccion.Reenlistadas = 200
                        Else
                            If FileExist(CharPath & M & ".chr") Then
                                Call WriteVar(CharPath & M & ".chr", "FACCIONES", "EjercitoCaos", 0)
                                Call WriteVar(CharPath & M & ".chr", "FACCIONES", "EjercitoReal", 0)
                                Call WriteVar(CharPath & M & ".chr", "FACCIONES", "Reenlistadas", 200)
                            End If
                        End If
                        m_ValidarPermanencia = True
                    Else    'sale si no es guildfounder
                        If m_EsGuildLeader(M, GuildIndex) Then
                            'pierde el liderazgo
                            CambioLider = True
                            Call guilds(GuildIndex).SetLeader(guilds(GuildIndex).Fundador)
                        End If

                        Call m_EcharMiembroDeClan(-1, M)
                    End If
                End If
            Next i
        Else
            'no se va el fundador, el peor caso es que se vaya el lider
            
            'If m_EsGuildLeader(UserList(UserIndex).Name, GuildIndex) Then
            '    Call LogClanes("Se transfiere el liderazgo de: " & Guilds(GuildIndex).GuildName & " a " & Guilds(GuildIndex).Fundador)
            '    Call Guilds(GuildIndex).SetLeader(Guilds(GuildIndex).Fundador)  'transferimos el lideraztgo
            'End If
            Call m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)   'y lo echamos
        End If
    End If
End Function

Public Sub m_DesconectarMiembroDelClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
    If UserList(UserIndex).GuildIndex > CANTIDADDECLANES Then Exit Sub
    Call guilds(GuildIndex).DesConectarMiembro(UserIndex)
End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean

    Dim UserId As Integer
    
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id FROM personaje WHERE nombre = '" & PJ & "'")
    UserId = RS!id
    
    Set RS = SQL.Execute("SELECT id FROM clan WHERE lider = '" & UserId & "'")
    
    If RS.RecordCount > 0 Then
        m_EsGuildLeader = True
    Else
        m_EsGuildLeader = False
        
    End If
    
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).Fundador)))
End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
'UI echa a Expulsado del clan de Expulsado
Dim UserIndex   As Integer
Dim GI          As Integer
    
    m_EcharMiembroDeClan = 0

    UserIndex = NameIndex(Expulsado)
    If UserIndex > 0 Then
        'pj online
        GI = UserList(UserIndex).GuildIndex
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
                Call guilds(GI).DesConectarMiembro(UserIndex)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                UserList(UserIndex).GuildIndex = 0
                Call RefreshCharStatus(UserIndex)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    Else
        'pj offline
        GI = GetGuildIndexFromChar(Expulsado)
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    End If

End Function

Public Sub ActualizarWebSite(ByVal UserIndex As Integer, ByRef Web As String)
Dim GI As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then Exit Sub
    
    Call guilds(GI).SetURL(Web)
    
End Sub


Public Sub ChangeCodexAndDesc(ByRef desc As String, ByRef codex() As String, ByVal GuildIndex As Integer)
    Dim i As Long
    
    If GuildIndex < 1 Or GuildIndex > CANTIDADDECLANES Then Exit Sub
    
    With guilds(GuildIndex)
        Call .SetDesc(desc)
        
        For i = 0 To UBound(codex())
            Call .SetCodex(i, codex(i))
        Next i
        
        For i = i To CANTIDADMAXIMACODEX
            Call .SetCodex(i, vbNullString)
        Next i
    End With
End Sub

Public Sub ActualizarNoticias(ByVal UserIndex As Integer, ByRef Datos As String)
Dim GI              As Integer

    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then Exit Sub
    
    Call guilds(GI).SetGuildNews(Datos)
        
End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, ByRef desc As String, ByRef GuildName As String, ByRef URL As String, ByRef codex() As String, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
Dim CantCodex       As Integer
Dim i               As Integer
Dim DummyString     As String

    CrearNuevoClan = False
    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function
    End If

    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de clan inv�lido."
        Exit Function
    End If
    
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT nombre FROM clan WHERE nombre = '" & GuildName & "'")
    
    If RS.RecordCount > 0 Then
        refError = "Nombre de clan inv�lido."
        Exit Function
    End If
    RS.Close
    
    CantCodex = UBound(codex()) + 1
    
    CANTIDADDECLANES = CANTIDADDECLANES + 1

    '01/10/2015 Irongete: A�ado el clan a la base de datos
    Dim Ahora As String
    Ahora = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & " " & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now)
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("INSERT INTO clan (fundador, nombre, fecha_fundacion, lider, `desc`) VALUES ('" & UserList(FundadorIndex).id & "', '" & GuildName & "', '" & Ahora & "', '" & UserList(FundadorIndex).id & "', '" & desc & "')")
    
    '12/11/2015 Irongete: Id del clan recien creado
    Dim clanid As Integer
    Set RS = SQL.Execute("SELECT id FROM clan WHERE nombre = '" & GuildName & "'")
    clanid = RS!id
    
    '13/11/2015 Irongete: Inicializar los puntos de castillos para este clan
    Set RS = SQL.Execute("INSERT INTO rel_clan_puntos (id_clan, puntoscastillos, minutoscastillo1, minutoscastillo2, minutoscastillo3, minutoscastillo4, minutoscastillo5) VALUES ('" & clanid & "', '0', '0', '0', '0', '0', '0')")
    
    UserList(FundadorIndex).GuildIndex = clanid
    
    '17/11/2015 Irongete: Actualizo el campo del clan del personaje
    Set RS = SQL.Execute("UPDATE personaje SET id_clan = '" & clanid & "' WHERE id = '" & UserList(FundadorIndex).id & "'")
    
  
    Call RefreshCharStatus(FundadorIndex)
    
    'For i = 1 To CANTIDADDECLANES - 1
    '    Call guilds(i).ProcesarFundacionDeOtroClan
    'Next i
    
    CrearNuevoClan = True
End Function

Public Sub SendGuildNews(ByVal UserIndex As Integer)
Dim GuildIndex  As Integer
Dim i               As Integer
Dim go As Integer

    GuildIndex = UserList(UserIndex).GuildIndex
    If GuildIndex = 0 Then Exit Sub

    Dim enemies() As String
    
    If guilds(GuildIndex).CantidadEnemys Then
        ReDim enemies(0 To guilds(GuildIndex).CantidadEnemys - 1) As String
    Else
        ReDim enemies(0)
    End If
    
    Dim allies() As String
    
    If guilds(GuildIndex).CantidadAllies Then
        ReDim allies(0 To guilds(GuildIndex).CantidadAllies - 1) As String
    Else
        ReDim allies(0)
    End If
    
    i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
    go = 0
    
    While i > 0
        enemies(go) = guilds(i).GuildName
        i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
        go = go + 1
    Wend
    
    i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
    go = 0
    
    While i > 0
        allies(go) = guilds(i).GuildName
        i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
    Wend

    Call WriteGuildNews(UserIndex, guilds(GuildIndex).GetGuildNews, enemies, allies)

    If guilds(GuildIndex).EleccionesAbiertas Then
        Call WriteConsoleMsg(UserIndex, "Hoy es la votacion para elegir un nuevo l�der para el clan!!.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(UserIndex, "La eleccion durara 24 horas, se puede votar a cualquier miembro del clan.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(UserIndex, "Para votar escribe /VOTO NICKNAME.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(UserIndex, "Solo se computara un voto por miembro. Tu voto no puede ser cambiado.", FontTypeNames.FONTTYPE_GUILD)
    End If

End Sub

Public Function m_PuedeSalirDeClan(ByRef nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
'sale solo si no es fundador del clan.

    m_PuedeSalirDeClan = False
    If GuildIndex = 0 Then Exit Function
    
    'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
    If QuienLoEchaUI = -1 Then
        m_PuedeSalirDeClan = True
        Exit Function
    End If

    'cuando UI no puede echar a nombre?
    'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
    If UserList(QuienLoEchaUI).flags.Privilegios And PlayerType.User Then
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).Name), GuildIndex) Then
            If UCase$(UserList(QuienLoEchaUI).Name) <> UCase$(nombre) Then      'si no sale voluntariamente...
                Exit Function
            End If
        End If
    End If

    m_PuedeSalirDeClan = UCase$(guilds(GuildIndex).Fundador) <> UCase$(nombre)

End Function

Public Function PuedeFundarUnClan(ByVal UserIndex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean

    PuedeFundarUnClan = False
    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, no puedes fundar otro"
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.ELV < 40 Then
        If Not UserList(UserIndex).Stats.DragCredits >= 50 Then
            refError = "Para fundar un clan necesitas ser nivel 40 o disponer de 50 DragCreditos."
            Exit Function
        End If
    End If
    
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            If UserList(UserIndex).Faccion.ArmadaReal <> 1 Then
                refError = "Para fundar un clan real debes ser miembro de la armada."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            If criminal(UserIndex) Then
                refError = "Para fundar un clan de ciudadanos no debes ser criminal."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            If Not criminal(UserIndex) Then
                refError = "Para fundar un clan de criminales no debes ser ciudadano."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_LEGION
            If UserList(UserIndex).Faccion.FuerzasCaos <> 1 Then
                refError = "Para fundar un clan del mal debes pertenecer a la legi�n oscura"
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_MASTER
            If UserList(UserIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                refError = "Para fundar un clan sin alineaci�n debes ser un dios."
                Exit Function
            End If
    End Select
    
    PuedeFundarUnClan = True
    
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, ByVal GuildIndex As Integer) As Boolean
Dim Promedio    As Long
Dim ELV         As Integer
Dim f           As Byte

    m_EstadoPermiteEntrarChar = False
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace(Personaje, ".", vbNullString)
    End If
    
    If FileExist(CharPath & Personaje & ".chr") Then
        Promedio = CLng(GetVar(CharPath & Personaje & ".chr", "REP", "Promedio"))
        Select Case guilds(GuildIndex).Alineacion
            Case ALINEACION_GUILD.ALINEACION_ARMADA
                If Promedio >= 0 Then
                    ELV = CInt(GetVar(CharPath & Personaje & ".chr", "Stats", "ELV"))
                    If ELV >= 25 Then
                        f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoReal"))
                    End If
                    m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f <> 0, True)
                End If
            Case ALINEACION_GUILD.ALINEACION_CIUDA
                m_EstadoPermiteEntrarChar = Promedio >= 0
            Case ALINEACION_GUILD.ALINEACION_CRIMINAL
                m_EstadoPermiteEntrarChar = Promedio < 0
            Case ALINEACION_GUILD.ALINEACION_NEUTRO
                m_EstadoPermiteEntrarChar = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoReal")) = 0
                m_EstadoPermiteEntrarChar = m_EstadoPermiteEntrarChar And (CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoCaos")) = 0)
            Case ALINEACION_GUILD.ALINEACION_LEGION
                If Promedio < 0 Then
                    ELV = CInt(GetVar(CharPath & Personaje & ".chr", "Stats", "ELV"))
                    If ELV >= 25 Then
                        f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoCaos"))
                    End If
                    m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f <> 0, True)
                End If
            Case Else
                m_EstadoPermiteEntrarChar = True
        End Select
    End If
End Function

Private Function m_EstadoPermiteEntrar(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
    Select Case guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            m_EstadoPermiteEntrar = Not criminal(UserIndex) And _
                    IIf(UserList(UserIndex).Stats.ELV >= 25, UserList(UserIndex).Faccion.ArmadaReal <> 0, True)
        Case ALINEACION_GUILD.ALINEACION_LEGION
            m_EstadoPermiteEntrar = criminal(UserIndex) And _
                    IIf(UserList(UserIndex).Stats.ELV >= 25, UserList(UserIndex).Faccion.FuerzasCaos <> 0, True)
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            m_EstadoPermiteEntrar = UserList(UserIndex).Faccion.ArmadaReal = 0 And UserList(UserIndex).Faccion.FuerzasCaos = 0
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            m_EstadoPermiteEntrar = Not criminal(UserIndex)
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            m_EstadoPermiteEntrar = criminal(UserIndex)
        Case Else   'game masters
            m_EstadoPermiteEntrar = True
    End Select
End Function


Public Function String2Alineacion(ByRef S As String) As ALINEACION_GUILD
    Select Case S
        Case "Neutro"
            String2Alineacion = ALINEACION_NEUTRO
        Case "Legi�n oscura"
            String2Alineacion = ALINEACION_LEGION
        Case "Armada Real"
            String2Alineacion = ALINEACION_ARMADA
        Case "Game Masters"
            String2Alineacion = ALINEACION_MASTER
        Case "Legal"
            String2Alineacion = ALINEACION_CIUDA
        Case "Criminal"
            String2Alineacion = ALINEACION_CRIMINAL
    End Select
End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            Alineacion2String = "Neutro"
        Case ALINEACION_GUILD.ALINEACION_LEGION
            Alineacion2String = "Legi�n oscura"
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            Alineacion2String = "Armada Real"
        Case ALINEACION_GUILD.ALINEACION_MASTER
            Alineacion2String = "Game Masters"
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            Alineacion2String = "Legal"
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            Alineacion2String = "Criminal"
    End Select
End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
    Select Case Relacion
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "A"
        Case RELACIONES_GUILD.GUERRA
            Relacion2String = "G"
        Case RELACIONES_GUILD.PAZ
            Relacion2String = "P"
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "?"
    End Select
End Function

Public Function String2Relacion(ByVal S As String) As RELACIONES_GUILD
    Select Case UCase$(Trim$(S))
        Case vbNullString, "P"
            String2Relacion = RELACIONES_GUILD.PAZ
        Case "G"
            String2Relacion = RELACIONES_GUILD.GUERRA
        Case "A"
            String2Relacion = RELACIONES_GUILD.ALIADOS
        Case Else
            String2Relacion = RELACIONES_GUILD.PAZ
    End Select
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
Dim car     As Byte
Dim i       As Integer

'old function by morgo

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))

    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        GuildNameValido = False
        Exit Function
    End If
    
Next i

GuildNameValido = True

End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
Dim i   As Integer

YaExiste = False
GuildName = UCase$(GuildName)

For i = 1 To CANTIDADDECLANES
    YaExiste = (UCase$(guilds(i).GuildName) = GuildName)
    If YaExiste Then Exit Function
Next i



End Function

Public Function v_AbrirElecciones(ByVal UserIndex As Integer, ByRef refError As String) As Boolean
Dim GuildIndex      As Integer

    v_AbrirElecciones = False
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ning�n clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GuildIndex) Then
        refError = "No eres el l�der de tu clan"
        Exit Function
    End If
    
    If guilds(GuildIndex).EleccionesAbiertas Then
        refError = "Las elecciones ya est�n abiertas"
        Exit Function
    End If
    
    v_AbrirElecciones = True
    Call guilds(GuildIndex).AbrirElecciones
    
End Function

Public Function v_UsuarioVota(ByVal UserIndex As Integer, ByRef Votado As String, ByRef refError As String) As Boolean
Dim GuildIndex      As Integer
Dim list()          As String
Dim i As Long

    v_UsuarioVota = False
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ning�n clan"
        Exit Function
    End If

    If Not guilds(GuildIndex).EleccionesAbiertas Then
        refError = "No hay elecciones abiertas en tu clan."
        Exit Function
    End If
    
    
    list = guilds(GuildIndex).GetMemberList()
    For i = 0 To UBound(list())
        If UCase$(Votado) = list(i) Then Exit For
    Next i
    
    If i > UBound(list()) Then
        refError = Votado & " no pertenece al clan"
        Exit Function
    End If
    
    
    If guilds(GuildIndex).YaVoto(UserList(UserIndex).Name) Then
        refError = "Ya has votado, no puedes cambiar tu voto"
        Exit Function
    End If
    
    Call guilds(GuildIndex).ContabilizarVoto(UserList(UserIndex).Name, Votado)
    v_UsuarioVota = True

End Function

Public Sub v_RutinaElecciones()
Dim i       As Integer

On Error GoTo errh
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Revisando elecciones", FontTypeNames.FONTTYPE_SERVER))
    For i = 1 To CANTIDADDECLANES
        If Not guilds(i) Is Nothing Then
            If guilds(i).RevisarElecciones Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & guilds(i).GetLeader & " es el nuevo lider de " & guilds(i).GuildName & "!", FontTypeNames.FONTTYPE_SERVER))
            End If
        End If
proximo:
    Next i
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Elecciones revisadas", FontTypeNames.FONTTYPE_SERVER))
Exit Sub
errh:
    Call LogError("modGuilds.v_RutinaElecciones():" & Err.Description)
    Resume proximo
End Sub

Private Function GetGuildIndexFromChar(ByRef PlayerName As String) As Integer
'aca si que vamos a violar las capas deliveradamente ya que
'visual basic no permite declarar metodos de clase
Dim Temps   As String
    If InStrB(PlayerName, "\") <> 0 Then
        PlayerName = Replace(PlayerName, "\", vbNullString)
    End If
    If InStrB(PlayerName, "/") <> 0 Then
        PlayerName = Replace(PlayerName, "/", vbNullString)
    End If
    If InStrB(PlayerName, ".") <> 0 Then
        PlayerName = Replace(PlayerName, ".", vbNullString)
    End If
    Temps = GetVar(CharPath & PlayerName & ".chr", "GUILD", "GUILDINDEX")
    If IsNumeric(Temps) Then
        GetGuildIndexFromChar = CInt(Temps)
    Else
        GetGuildIndexFromChar = 0
    End If
End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
'me da el indice del guildname
Dim i As Integer

    GuildIndex = 0
    GuildName = UCase$(GuildName)
    For i = 1 To CANTIDADDECLANES
        If UCase$(guilds(i).GuildName) = GuildName Then
            GuildIndex = i
            Exit Function
        End If
    Next i
End Function

Public Function MiembrosGuild(ByVal UserIndex As Integer) As Byte
'13/02/2016 - Lorwik.
'Devuelve la cantidad de miembros de la guild.
    MiembrosGuild = guilds(UserList(UserIndex).GuildIndex).CantidadDeMiembros
End Function

Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As String
Dim i As Integer
    
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        While i > 0
            'No mostramos dioses y admins
            If i <> UserIndex And ((UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or (UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0)) Then _
                m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).Name & ","
            i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        Wend
    End If
    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)
    End If
End Function

Public Function PrepareGuildsList() As String()
    Dim tStr() As String
    Dim i As Long
    
    If CANTIDADDECLANES = 0 Then
        ReDim tStr(0) As String
    Else
        ReDim tStr(CANTIDADDECLANES - 1) As String
        
        For i = 1 To CANTIDADDECLANES
            tStr(i - 1) = guilds(i).GuildName
        Next i
    End If
    
    PrepareGuildsList = tStr
End Function

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByRef GuildName As String)
    Dim codex(CANTIDADMAXIMACODEX - 1)  As String
    Dim GI      As Integer
    Dim i       As Long
    
    
    codex(1) = ""

    'Call Protocol.WriteGuildDetails(UserIndex, GuildName, "fundador", "fecha fundacion", "lider", _
                                    "url", "cantidad miembros", "elecciones", "alineacion", _
                                    "cantidad enemigos", "cantidad aliados", "puntos antifaccion", _
                                    "codex", "descripcion")

                            
    Call Protocol.WriteGuildDetails(UserIndex, GuildName, "fundador", "fecha fundacion", "lider", _
                                    "url", 0, False, "alineacion", _
                                    0, 0, "0", _
                                    codex, "descripcion")

End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)
'***************************************************
'Autor: Mariano Barrou (El Oso)
'Last Modification: 12/10/06
'Las Modified By: Juan Mart�n Sotuyo Dodero (Maraxus)
'***************************************************
    Dim GI      As Integer
    Dim guildList() As String
    Dim MemberList() As String
    Dim aspirantsList() As String
    Dim Tmp As String
    '13/11/2015 Irongete: Modificado para que funcione con SQL
    
    '14/11/2015 Irongete: Lista clanes
        Dim clanes(200) As String
        Dim i As Integer
        i = 0
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT nombre FROM clan ORDER BY id ASC")
        
        If RS.RecordCount > 0 Then
            While Not RS.EOF
                 clanes(i) = CStr(RS!nombre)
                 i = i + 1
                RS.MoveNext
            Wend
        End If
        guildList = clanes

    With UserList(UserIndex)
        GI = .GuildIndex
       
        If GI <= 0 Or GI > CANTIDADDECLANES Then
            'Send the guild list instead
            Call Protocol.WriteGuildList(UserIndex, guildList)
            Exit Sub
        End If
        
        '14/11/2015 Irongete: Comprobar si es el lider del clan
        
        If Not m_EsGuildLeader(.Name, GI) Then
            'Send the guild list instead
            Call Protocol.WriteGuildList(UserIndex, guildList)
            Exit Sub
        End If
        
        
        '14/11/2015 Irongete: Lista de miembros del clan
        Dim miembros(200) As String
        i = 0
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT nombre FROM personaje WHERE id_clan = '" & .GuildIndex & "'")
        
        If RS.RecordCount > 0 Then
            While Not RS.EOF
                 miembros(i) = CStr(RS!nombre)
                 i = i + 1
                RS.MoveNext
            Wend
        End If
        MemberList = miembros
        
        '14/11/2015 Irongete: Lista de solicitudes al clan
        Dim aspirantes(200) As String
        i = 0
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT j1.nombre AS nombre FROM rel_clan_solicitud JOIN personaje j1 ON j1.id = rel_clan_solicitud.id_personaje WHERE rel_clan_solicitud.id_clan = '" & .GuildIndex & "'")

        If RS.RecordCount > 0 Then
            While Not RS.EOF
                 aspirantes(i) = CStr(RS!nombre)
                 i = i + 1
                RS.MoveNext
            Wend
        End If
        aspirantsList = aspirantes
        
        Call WriteGuildLeaderInfo(UserIndex, guildList, MemberList, "", aspirantsList)
   
    End With
End Sub


Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
    'itera sobre los onlinemembers
    m_Iterador_ProximoUserIndex = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        m_Iterador_ProximoUserIndex = guilds(GuildIndex).m_Iterador_ProximoUserIndex()
    End If
End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
    'itera sobre los gms escuchando este clan
    Iterador_ProximoGM = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        Iterador_ProximoGM = guilds(GuildIndex).Iterador_ProximoGM()
    End If
End Function

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, ByVal tipo As RELACIONES_GUILD) As Integer
    'itera sobre las propuestas
    r_Iterador_ProximaPropuesta = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        r_Iterador_ProximaPropuesta = guilds(GuildIndex).Iterador_ProximaPropuesta(tipo)
    End If
End Function

Public Function GMEscuchaClan(ByVal UserIndex As Integer, ByVal GuildName As String) As Integer
Dim GI As Integer

    'listen to no guild at all
    If LenB(GuildName) = 0 And UserList(UserIndex).EscucheClan <> 0 Then
        'Quit listening to previous guild!!
        Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
        guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
        Exit Function
    End If
    
'devuelve el guildindex
    GI = GuildIndex(GuildName)
    If GI > 0 Then
        If UserList(UserIndex).EscucheClan <> 0 Then
            If UserList(UserIndex).EscucheClan = GI Then
                'Already listening to them...
                Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
                GMEscuchaClan = GI
                Exit Function
            Else
                'Quit listening to previous guild!!
                Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
                guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
            End If
        End If
        
        Call guilds(GI).ConectarGM(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = GI
        UserList(UserIndex).EscucheClan = GI
    Else
        Call WriteConsoleMsg(UserIndex, "Error, el clan no existe", FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = 0
    End If
    
End Function

Public Sub GMDejaDeEscucharClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
'el index lo tengo que tener de cuando me puse a escuchar
    UserList(UserIndex).EscucheClan = 0
    'Call guilds(GuildIndex).DesconectarGM(UserIndex)
End Sub
Public Function r_DeclararGuerra(ByVal UserIndex As Integer, ByRef GuildGuerra As String, ByRef refError As String) As Integer
Dim GI  As Integer
Dim GIG As Integer

    r_DeclararGuerra = 0
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ning�n clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el l�der de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildGuerra) = vbNullString Then
        refError = "No has seleccionado ning�n clan"
        Exit Function
    End If
    
    GIG = GuildIndex(GuildGuerra)
    If guilds(GI).GetRelacion(GIG) = GUERRA Then
        refError = "Tu clan ya est� en guerra con " & GuildGuerra & "."
        Exit Function
    End If
        
    If GI = GIG Then
        refError = "No puedes declarar la guerra a tu mismo clan"
        Exit Function
    End If

    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_DeclararGuerra: " & GI & " declara a " & GuildGuerra)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.GUERRA)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.GUERRA)
    
    r_DeclararGuerra = GIG

End Function


Public Function r_AceptarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPaz As String, ByRef refError As String) As Integer
'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ning�n clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el l�der de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPaz) = vbNullString Then
        refError = "No has seleccionado ning�n clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPaz)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.GUERRA Then
        refError = "No est�s en guerra con ese clan"
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay ninguna propuesta de paz para aceptar"
        Exit Function
    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.PAZ)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.PAZ)
    
    r_AceptarPropuestaDePaz = GIG
End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDeAlianza = 0
    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ning�n clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el l�der de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ning�n clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, ALIADOS) Then
        refError = "No hay propuesta de alianza del clan " & GuildPro
        Exit Function
    End If
    
    Call guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de alianza. " & guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDeAlianza = GIG

End Function


Public Function r_RechazarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDePaz = 0
    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ning�n clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el l�der de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ning�n clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay propuesta de paz del clan " & GuildPro
        Exit Function
    End If
    
    Call guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de paz. " & guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDePaz = GIG

End Function


Public Function r_AceptarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildAllie As String, ByRef refError As String) As Integer
'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    r_AceptarPropuestaDeAlianza = 0
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ning�n clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el l�der de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildAllie) = vbNullString Then
        refError = "No has seleccionado ning�n clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildAllie)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.PAZ Then
        refError = "No est�s en paz con el clan, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.ALIADOS) Then
        refError = "No hay ninguna propuesta de alianza para aceptar."
        Exit Function
    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.ALIADOS)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.ALIADOS)
    
    r_AceptarPropuestaDeAlianza = GIG

End Function


Public Function r_ClanGeneraPropuesta(ByVal UserIndex As Integer, ByRef OtroClan As String, ByVal tipo As RELACIONES_GUILD, ByRef Detalle As String, ByRef refError As String) As Boolean
Dim OtroClanGI      As Integer
Dim GI              As Integer

    r_ClanGeneraPropuesta = False
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ning�n clan"
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroClan)
    
    If OtroClanGI = GI Then
        refError = "No puedes declarar relaciones con tu propio clan"
        Exit Function
    End If
    
    If OtroClanGI <= 0 Or OtroClanGI > CANTIDADDECLANES Then
        refError = "El sistema de clanes esta inconsistente, el otro clan no existe!"
        Exit Function
    End If
    
    If guilds(OtroClanGI).HayPropuesta(GI, tipo) Then
        refError = "Ya hay propuesta de " & Relacion2String(tipo) & " con " & OtroClan
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el l�der de tu clan"
        Exit Function
    End If
    
    'de acuerdo al tipo procedemos validando las transiciones
    If tipo = RELACIONES_GUILD.PAZ Then
        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.GUERRA Then
            refError = "No est�s en guerra con " & OtroClan
            Exit Function
        End If
    ElseIf tipo = RELACIONES_GUILD.GUERRA Then
        'por ahora no hay propuestas de guerra
    ElseIf tipo = RELACIONES_GUILD.ALIADOS Then
        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.PAZ Then
            refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtroClan
            Exit Function
        End If
    End If
    
    Call guilds(OtroClanGI).SetPropuesta(tipo, GI, Detalle)
    r_ClanGeneraPropuesta = True

End Function

Public Function r_VerPropuesta(ByVal UserIndex As Integer, ByRef OtroGuild As String, ByVal tipo As RELACIONES_GUILD, ByRef refError As String) As String
Dim OtroClanGI      As Integer
Dim GI              As Integer
    
    r_VerPropuesta = vbNullString
    refError = vbNullString
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ning�n clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el l�der de tu clan"
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroGuild)
    
    If Not guilds(GI).HayPropuesta(OtroClanGI, tipo) Then
        refError = "No existe la propuesta solicitada"
        Exit Function
    End If
    
    r_VerPropuesta = guilds(GI).GetPropuesta(OtroClanGI, tipo)
    
End Function

Public Function r_ListaDePropuestas(ByVal UserIndex As Integer, ByVal tipo As RELACIONES_GUILD) As String()

    Dim GI  As Integer
    Dim i   As Integer
    Dim proposalCount As Integer
    Dim proposals() As String
    
    GI = UserList(UserIndex).GuildIndex
    
    If GI > 0 And GI <= CANTIDADDECLANES Then
        With guilds(GI)
            proposalCount = .CantidadPropuestas(tipo)
            
            'Resize array to contain all proposals
            If proposalCount > 0 Then
                ReDim proposals(proposalCount - 1) As String
            Else
                ReDim proposals(0) As String
            End If
            
            'Store each guild name
            For i = 0 To proposalCount - 1
                proposals(i) = guilds(.Iterador_ProximaPropuesta(tipo)).GuildName
            Next i
        End With
    End If
    
    r_ListaDePropuestas = proposals
End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByVal guild As Integer, ByRef Detalles As String)
    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")
    End If
    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")
    End If
    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")
    End If
    Call guilds(guild).InformarRechazoEnChar(Aspirante, Detalles)
End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")
    End If
    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")
    End If
    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")
    End If
    a_ObtenerRechazoDeChar = GetVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo")
    Call WriteVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo", vbNullString)
End Function

Public Function a_RechazarAspirante(ByVal UserIndex As Integer, ByRef nombre As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer

    a_RechazarAspirante = False
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ning�n clan"
        Exit Function
    End If

    NroAspirante = guilds(GI).NumeroDeAspirante(nombre)

    If NroAspirante = 0 Then
        refError = nombre & " no es aspirante a tu clan"
        Exit Function
    End If

    Call guilds(GI).RetirarAspirante(nombre, NroAspirante)
    refError = "Fue rechazada tu solicitud de ingreso a " & guilds(GI).GuildName
    a_RechazarAspirante = True

End Function

Public Function a_DetallesAspirante(ByVal UserIndex As Integer, ByRef nombre As String) As String
Dim GI              As Integer
Dim NroAspirante    As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        Exit Function
    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(nombre)
    If NroAspirante > 0 Then
        a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(NroAspirante)
    End If
    
End Function

Public Sub SendDetallesPersonaje(ByVal UserIndex As Integer, ByVal Personaje As String)
    Dim GI          As Integer
    Dim NroAsp      As Integer
    Dim GuildName   As String
    Dim UserFile    As clsIniReader
    Dim miembro     As String
    Dim GuildActual As Integer
    Dim list()      As String
    Dim i           As Long
    
    GI = UserList(UserIndex).GuildIndex
    
    Personaje = UCase$(Personaje)
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Call Protocol.WriteConsoleMsg(UserIndex, "No perteneces a ning�n clan", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        Call Protocol.WriteConsoleMsg(UserIndex, "No eres el l�der de tu clan", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace$(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace$(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace$(Personaje, ".", vbNullString)
    End If
    
    NroAsp = guilds(GI).NumeroDeAspirante(Personaje)
    
    If NroAsp = 0 Then
        list = guilds(GI).GetMemberList()
        
        For i = 0 To UBound(list())
            If Personaje = list(i) Then Exit For
        Next i
        
        If i > UBound(list()) Then
            Call Protocol.WriteConsoleMsg(UserIndex, "El personaje no es ni aspirante ni miembro del clan", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    'ahora traemos la info
    
    Set UserFile = New clsIniReader
    
    With UserFile
        .Initialize (CharPath & Personaje & ".chr")
        
        ' Get the character's current guild
        GuildActual = val(.GetValue("GUILD", "GuildIndex"))
        If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
            GuildName = "<" & guilds(GuildActual).GuildName & ">"
        Else
            GuildName = "Ninguno"
        End If
        
        'Get previous guilds
        miembro = .GetValue("GUILD", "Miembro")
        If Len(miembro) > 400 Then
            miembro = ".." & Right$(miembro, 400)
        End If
        
        Call Protocol.WriteCharacterInfo(UserIndex, Personaje, .GetValue("INIT", "Raza"), .GetValue("INIT", "Clase"), _
                                .GetValue("INIT", "Genero"), .GetValue("STATS", "ELV"), .GetValue("STATS", "GLD"), _
                                .GetValue("STATS", "Banco"), .GetValue("REP", "Promedio"), .GetValue("GUILD", "Pedidos"), _
                                GuildName, miembro, .GetValue("FACCIONES", "EjercitoReal"), .GetValue("FACCIONES", "EjercitoCaos"), _
                                .GetValue("FACCIONES", "CiudMatados"), .GetValue("FACCIONES", "CrimMatados"))
    End With
    
    Set UserFile = Nothing
End Sub

Public Function a_NuevoAspirante(ByVal UserIndex As Integer, ByRef clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
Dim ViejoSolicitado     As String
Dim ViejoGuildINdex     As Integer
Dim ViejoNroAspirante   As Integer
Dim NuevoGuildIndex     As Integer

    a_NuevoAspirante = False

    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro"
        Exit Function
    End If

    NuevoGuildIndex = GuildIndex(clan)
    If NuevoGuildIndex = 0 Then
        refError = "Ese clan no existe! Avise a un administrador."
        Exit Function
    End If
    
    If Not m_EstadoPermiteEntrar(UserIndex, NuevoGuildIndex) Then
        refError = "Tu no puedes entrar a un clan de alineaci�n " & Alineacion2String(guilds(NuevoGuildIndex).Alineacion)
        Exit Function
    End If
    
    '13/02/2016 Lorwik: Limite de miembros del clan
    If guilds(NuevoGuildIndex).CantidadDeMiembros = MIEMBROSMAXIMOS Then
        refError = "El clan ya cuenta con la cantidad maxima de miembros."
        Exit Function
    End If

    If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = "El clan tiene demasiados aspirantes. Cont�ctate con un miembro para que procese las solicitudes."
        Exit Function
    End If

    ViejoSolicitado = GetVar(CharPath & UserList(UserIndex).Name & ".chr", "GUILD", "ASPIRANTEA")

    If LenB(ViejoSolicitado) <> 0 Then
        'borramos la vieja solicitud
        ViejoGuildINdex = CInt(ViejoSolicitado)
        If ViejoGuildINdex <> 0 Then
            ViejoNroAspirante = guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(UserIndex).Name)
            If ViejoNroAspirante > 0 Then
                Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(UserIndex).Name, ViejoNroAspirante)
            End If
        Else
            'RefError = "Inconsistencia en los clanes, avise a un administrador"
            'Exit Function
        End If
    End If
    
    Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(UserIndex).Name, Solicitud)
    a_NuevoAspirante = True
End Function

Public Function a_AceptarAspirante(ByVal UserIndex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer
Dim AspiranteUI     As Integer

    'un pj ingresa al clan :D

    a_AceptarAspirante = False
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ning�n clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el l�der de tu clan"
        Exit Function
    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(Aspirante)
    
    If NroAspirante = 0 Then
        refError = "El Pj no es aspirante al clan"
        Exit Function
    End If
    
    '13/02/2016 Lorwik: Limite de miembros del clan
    If guilds(UserIndex).CantidadDeMiembros = MIEMBROSMAXIMOS Then
        refError = "No puedes aceptar mas miembros. El clan ya cuenta con la cantidad maxima de miembros."
        Exit Function
    End If
    
    AspiranteUI = NameIndex(Aspirante)
    If AspiranteUI > 0 Then
        'pj Online
        If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
            refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf Not UserList(AspiranteUI).GuildIndex = 0 Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    Else
        If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
            refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf GetGuildIndexFromChar(Aspirante) Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    End If
    'el pj es aspirante al clan y puede entrar
    
    Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
    Call guilds(GI).AceptarNuevoMiembro(Aspirante)
    
    ' If player is online, update tag
    If AspiranteUI > 0 Then
        Call RefreshCharStatus(AspiranteUI)
    End If
    
    a_AceptarAspirante = True
End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT nombre FROM clan WHERE id = '" & GuildIndex & "'")
    GuildName = RS!nombre
End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildLeader = guilds(GuildIndex).GetLeader
End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)
End Function

Public Function GuildFounder(ByVal GuildIndex As Integer) As String
'***************************************************
'Autor: ZaMa
'Returns the guild founder's name
'Last Modification: 25/03/2009
'***************************************************
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildFounder = guilds(GuildIndex).Fundador
End Function
