Attribute VB_Name = "Drag_Party"
Option Explicit

Public Const PARTY_MAX_DISTANCIA_PARA_EXP As Byte = 20
Public Const PARTY_PORCENTAJE_EXP_POR_CADA_MIEMBRO As Integer = 3


Public Sub AñadirAParty(ByVal PartyId As Integer, ByVal InvitadoIndex As Integer)
    
    Dim Index As Integer
    
    '20/11/2015 Irongete: Añado el personaje a la party
    UserList(InvitadoIndex).PartyId = PartyId
    Set RS = SQL.Execute("INSERT INTO rel_party_personaje (id_party, id_personaje) VALUES ('" & PartyId & "', '" & UserList(InvitadoIndex).ID & "')")
        
    '18/11/2015 Irongete: Elimino la invitacion
    Set RS = SQL.Execute("DELETE FROM rel_party_invitacion WHERE id_invitado = '" & UserList(InvitadoIndex).ID & "'")
    
    '18/11/2015 Irongete: Les envio los mensajes
    Call BroadcastParty(PartyId, UserList(InvitadoIndex).Name & " ha entrado en la party.")
    
    '20/11/2015 Irongete: Envio al jugador cual es su PartyId
    Call WriteSetPartyId(InvitadoIndex, 0, PartyId)
    
    '17/02/2016 Irongete: Le envio al nuevo miembro el CharIndex de los miembros que ya estaban dentro de la party
    'y a los que ya estaban dentro les envio el del nuevo miembro
    Set RS = SQL.Execute("SELECT j1.userindex FROM rel_party_personaje JOIN personaje j1 ON j1.id = rel_party_personaje.id_personaje WHERE id_party = '" & PartyId & "' AND rel_party_personaje.id_personaje <> '" & UserList(InvitadoIndex).ID & "'")
    If RS.RecordCount > 0 Then
        While Not RS.EOF
            Index = RS!UserIndex
            Call WriteSetPartyId(InvitadoIndex, Index, PartyId)
            Call WriteSetPartyId(Index, InvitadoIndex, PartyId)
            RS.MoveNext
        Wend
    End If
    
    
End Sub


'17/11/2015 Irongete: Crea una party
' @param LiderIndex userindex del lider de la party
' @param InvitadoIndex userindex de la persona que ha sido invitada
Public Sub CrearParty(ByVal LiderIndex As Integer, ByVal InvitadoIndex As Integer)
    On Error GoTo Errhandler

    Dim PartyId As Integer
    Dim Fecha As String
    Dim LiderId As Integer
    
    Fecha = Ahora()
    LiderId = UserList(LiderIndex).ID
    
    Set RS = New ADODB.Recordset
        
    '17/11/2015 Irongete: Inserto los datos en la base de datos
    Set RS = SQL.Execute("INSERT INTO party (fecha_creacion, lider) VALUES ('" & Fecha & "', '" & LiderId & "')")
    
    '17/11/2015 Irongete: Obtengo la Id de la party recién creada
    Set RS = SQL.Execute("SELECT id FROM party WHERE fecha_creacion = '" & Fecha & "' AND lider = '" & LiderId & "'")
    PartyId = RS!ID
    
    '17/11/2015 Irongete: Añado los miembros que forman la party
    UserList(LiderIndex).PartyId = PartyId
    UserList(InvitadoIndex).PartyId = PartyId
    Set RS = SQL.Execute("INSERT INTO rel_party_personaje (id_party, id_personaje) VALUES ('" & PartyId & "', '" & UserList(LiderIndex).ID & "')")
    Set RS = SQL.Execute("INSERT INTO rel_party_personaje (id_party, id_personaje) VALUES ('" & PartyId & "', '" & UserList(InvitadoIndex).ID & "')")
        
    '18/11/2015 Irongete: Elimino la invitacion
    Set RS = SQL.Execute("DELETE FROM rel_party_invitacion WHERE  id_invita = '" & UserList(LiderIndex).ID & "' AND id_invitado = '" & UserList(InvitadoIndex).ID & "'")
    
    '18/11/2015 Irongete: Les envio los mensajes
    Call WriteConsoleMsg(LiderIndex, "Has creado una party.", FontTypeNames.FONTTYPE_PARTY)
    Call BroadcastParty(PartyId, UserList(LiderIndex).Name & " ha entrado en la party.")
    Call BroadcastParty(PartyId, UserList(InvitadoIndex).Name & " ha entrado en la party.")
    
    '20/11/2015 Irongete: Envio a cada jugador cual es su PartyId
    Call WriteSetPartyId(LiderIndex, 0, PartyId)
    Call WriteSetPartyId(InvitadoIndex, 0, PartyId)
    
    '17/02/2016 Irongete: Envio a cada jugador el CharIndex de la otra persona para que le ponga el PartyId
    Call WriteSetPartyId(LiderIndex, InvitadoIndex, PartyId)
    Call WriteSetPartyId(InvitadoIndex, LiderIndex, PartyId)
    
    
Errhandler:
    Debug.Print "Error en  CrearParty(): " & Err.Description
End Sub


'18/11/2015 Irongete: Mensaje a todos los miembros de la party
Public Sub BroadcastParty(ByVal PartyId As Integer, ByVal Mensaje As String)
    Set RS = New ADODB.Recordset
    If PartyId > 0 Then
        '18/11/2015 Irongete: Envío el mensaje a todos los miembros de la party que estén logeados
        Set RS = SQL.Execute("SELECT j1.userindex AS UserIndex FROM rel_party_personaje JOIN personaje j1 ON j1.id = rel_party_personaje.id_personaje WHERE rel_party_personaje.id_party = '" & PartyId & "' AND j1.logged = '1'")
        While Not RS.EOF
            Call WriteConsoleMsg(RS!UserIndex, Mensaje, FontTypeNames.FONTTYPE_PARTY)
            RS.MoveNext
        Wend
    End If
End Sub

'17/11/2015 Irongete: Devuelve si un personaje está en party o no
Function EstaEnParty(ByRef UserIndex As Integer) As Boolean
    Set RS = SQL.Execute("SELECT id_party FROM rel_party_personaje WHERE id_personaje = '" & UserList(UserIndex).ID & "'")
    If RS.RecordCount = 1 Then
        EstaEnParty = True
    Else
        EstaEnParty = False
    End If
End Function

'17/11/2015 Irongete: Devuelve si un personaje tiene una invitación pendiente a party o no
Function TieneInvitacionAPartyPendiente(ByRef UserIndex As Integer) As Boolean
    Set RS = SQL.Execute("SELECT id_invita FROM rel_party_invitacion WHERE id_invitado = '" & UserList(UserIndex).ID & "'")
    
    '18/11/2015 Irongete: Si tiene mas de 1 invitacion pendiente algo no va bien, borro todo y acepto esta nueva
    If RS.RecordCount = 0 Then
        TieneInvitacionAPartyPendiente = False
    ElseIf RS.RecordCount > 0 Then
        Set RS = SQL.Execute("DELETE FROM rel_party_invitacion WHERE id_invitado = '" & UserList(UserIndex).ID & "'")
    Else
        TieneInvitacionAPartyPendiente = True
    End If
End Function

'17/11/2015 Irongete: Devuelve si un personaje es lider o no de una party
Function EsLiderDeParty(ByRef UserIndex As Integer) As Boolean
    Set RS = SQL.Execute("SELECT id FROM party WHERE lider = '" & UserList(UserIndex).ID & "'")
    If RS.RecordCount = 1 Then
        EsLiderDeParty = True
    Else
        EsLiderDeParty = False
    End If
End Function

'17/11/2015 Irongete: Devuelve la id de la party en la que esta el personaje o 0 si no está en ninguna
Function GetPartyId(ByRef UserIndex As Integer) As Integer
    Set RS = SQL.Execute("SELECT id_party FROM rel_party_personaje WHERE id_personaje = '" & UserList(UserIndex).ID & "'")
    If RS.RecordCount = 1 Then
        GetPartyId = RS!id_party
    Else
        GetPartyId = 0
    End If
End Function


'17/11/2015 Irongete: Adaptación de la funcion SolParty de Lorwik
' @param LiderIndex index de la persona que envía la invitación
Sub HandleInvitacionAParty(ByVal LiderIndex As Integer)

    With UserList(LiderIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        Dim InvitadoIndex As Integer

        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then Exit Sub

        If Not InRangoVision(LiderIndex, X, Y) Then
            Call WritePosUpdate(LiderIndex)
            Exit Sub
        End If
        
        Call ClickIzquierdo(LiderIndex, .Pos.Map, X, Y)
        
        InvitadoIndex = .flags.targetUser
                
        'Validate target
        If Not InvitadoIndex > 0 Then Exit Sub
        
        
        '¿Esta cerca?
        If Abs(UserList(InvitadoIndex).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
            Call WriteMultiMessage(InvitadoIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        'Prevent from hitting self
        If LiderIndex = InvitadoIndex Then
            Call WriteMultiMessage(LiderIndex, eMessages.SolicitudPropia)
            Exit Sub
        End If
            
        '17/11/2015 Irongete: Comprobar que el que recibe la invitación no está ya en party
        If EstaEnParty(InvitadoIndex) Then
            Call WriteConsoleMsg(LiderIndex, UserList(InvitadoIndex).Name & " ya está en party.", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        End If
        
        '17/11/2015 Irongete: Comprobar que el que recibe la invitación no tiene ya una invitación a party pendiente
        If TieneInvitacionAPartyPendiente(InvitadoIndex) Then
            Call WriteConsoleMsg(LiderIndex, UserList(InvitadoIndex).Name & " tiene una invitación a party pendiente.", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        End If
        
        '17/11/2015 Irongete: Le mando la invitación
        Call WriteConsoleMsg(LiderIndex, "Has invitado a " & UserList(InvitadoIndex).Name & " a party", FontTypeNames.FONTTYPE_PARTY)
        Call WriteConsoleMsg(InvitadoIndex, UserList(LiderIndex).Name & " te ha invitado a party.", FontTypeNames.FONTTYPE_PARTY)
        
        '18/11/2015 Irongete: Guardo la invitación en la base de datos
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("INSERT INTO rel_party_invitacion (id_invita, id_invitado, fecha) VALUES ('" & UserList(LiderIndex).ID & "', '" & UserList(InvitadoIndex).ID & "', '" & Ahora() & "')")
            
        Call WritePeticionInvitacionAParty(LiderIndex, InvitadoIndex)
  

    End With
End Sub

'19/11/2015 Irongete: Funcion que reparte la experiencia entre todos los miembros
' @param UserIndex userindex de la persona que da el golpe final al npc
' @param ExpFinal exp calculada antes de dar el bonus de party
' @param PartyId id de la party
' @param X posicion x de donde ha muerto el npc
' @param Y posicion y de donde ha muerto el npc
' @param NivelNpc Nivel del NPC que han matado

Public Sub RepartirExpParty(ByVal UserIndex As Integer, ExpFinal As Long, ByVal PartyId As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal NivelNpc As Integer)
    On Error GoTo Errhandler

    Dim PartyUserIndex As Integer
    Dim TotalOnlineParty As Integer
    Dim ExpExtra As Integer
    Dim JugadoresCerca As Integer
    Dim PorcentajePorJugador As Long
    Dim DistanciaMaximaParaExp As Integer

    Set RS = New ADODB.Recordset

    '18/11/2015 Irongete: Saco los valores de configuracion de la base de datos
    Set RS = SQL.Execute("SELECT party_porcentaje_exp_por_jugador,party_distancia_maxima_para_exp FROM config")
    PorcentajePorJugador = RS!party_porcentaje_exp_por_jugador
    DistanciaMaximaParaExp = RS!party_distancia_maxima_para_exp
    
    
    '19/11/2015 Irongete: Comprobar si hay miembros cerca o está el jugador solo
    JugadoresCerca = PartyCuantosCerca(UserIndex, PartyId, X, Y, DistanciaMaximaParaExp)
    If JugadoresCerca = 0 Then
    
        '19/11/2015 Irongete: Sumar la exp
        ExpFinal = ExpFinal
        
        'Si el usuario esta por debajo del nivel del NPC le damos solo la mitad de la exp.
        If UserList(UserIndex).Stats.ELV > NivelNpc Then
            ExpFinal = ExpFinal / 10
        End If
        
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + Fix(ExpFinal)
        If UserList(UserIndex).Stats.Exp > MAXEXP Then _
            UserList(UserIndex).Stats.Exp = MAXEXP
        Call CheckUserLevel(UserIndex)
        Call WriteUpdateUserStats(UserIndex)
        
        '19/11/2015 Irongete: Mensaje a los miembros de la party
        Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpFinal & " puntos de experiencia.", FontTypeNames.FONTTYPE_exp)
    
    
    Else
    
        
        '19/11/2015 Irongete: Calcular la experiencia que le voy a dar a los jugadores
        ExpExtra = CInt((ExpFinal * ((JugadoresCerca + 1) * PorcentajePorJugador)) / 100)
        ExpFinal = CInt((ExpFinal + ExpExtra) / (JugadoresCerca + 1))
        
                    
        '18/11/2015 Irongete: Userindex de los miembros de la party que están conectados
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT personaje.userindex AS PartyUserIndex FROM personaje INNER JOIN rel_party_personaje ON personaje.id = rel_party_personaje.id_personaje WHERE rel_party_personaje.id_party = '" & PartyId & "' AND personaje.logged = '1'")
        
            
        While Not RS.EOF
        
            PartyUserIndex = CInt(RS!PartyUserIndex)
            
            '19/11/2015 Irongete: Por si acaso
            If PartyUserIndex > 0 Then
            
                 '19/11/2015 Irongete: Comprobar la distancia
                 If Distance(UserList(PartyUserIndex).Pos.X, UserList(RS!PartyUserIndex).Pos.Y, X, Y) <= DistanciaMaximaParaExp Then
                    
                 
                    'Si el usuario esta por debajo del nivel del NPC le damos solo la mitad de la exp.
                    If UserList(PartyUserIndex).Stats.ELV > NivelNpc Then
                        UserList(PartyUserIndex).Stats.Exp = UserList(PartyUserIndex).Stats.Exp + ExpFinal / 10
                        Call WriteConsoleMsg(PartyUserIndex, "Has ganado " & CInt(ExpFinal / 10) & " puntos de experiencia. (+" & ExpExtra & " bonus party)", FontTypeNames.FONTTYPE_exp)
                    Else
                        UserList(PartyUserIndex).Stats.Exp = UserList(PartyUserIndex).Stats.Exp + ExpFinal
                        Call WriteConsoleMsg(PartyUserIndex, "Has ganado " & ExpFinal & " puntos de experiencia. (+" & ExpExtra & " bonus party)", FontTypeNames.FONTTYPE_exp)
                    End If

                 End If
            End If
            
            RS.MoveNext
        Wend
    
    End If
    
    
    
Errhandler:
    Debug.Print "Error en RepartirExp" & Err.Number & ": " & Err.source
    
End Sub
Public Sub BorrarParty(ByVal PartyId As Integer)

    Set RS = New ADODB.Recordset

    '20/11/2015 Irongete: Borro la party
    Set RS = SQL.Execute("DELETE FROM party WHERE id = '" & PartyId & "'")
        
End Sub
Public Sub NuevoLiderParty(ByVal NuevoLiderIndex As Integer)

    Dim PartyId As Integer
    Set RS = New ADODB.Recordset

    '20/11/2015 Irongete: Compruebo que el nuevo lider esta en party
    PartyId = GetPersonajePartyId(NuevoLiderIndex)
    If PartyId > 0 Then
    
        '20/11/2015 Irongete: Lo hago lider de la party
        Set RS = SQL.Execute("UPDATE party SET lider = '" & UserList(NuevoLiderIndex).ID & "' WHERE id = '" & PartyId & "'")
        Call BroadcastParty(PartyId, UserList(NuevoLiderIndex).Name & " es el nuevo lider de la party")
    End If
End Sub
'19/11/2015 Irongete: Funcion que devuelve cuantos jugadores de la party hay cerca de X,Y sin contar UserIndex que es quien le da el golpe al npc
' @param UserIndex userindex del jugador que mata el npc
' @param PartyId id de la party
' @param Distancia distancia maxima a la que se contara que el jugador está cerca
Function PartyCuantosCerca(ByVal UserIndex As Integer, ByVal PartyId As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Distancia As Integer) As Integer

    Dim PartyUserIndex As Integer
    Dim Total As Integer
    Dim DistanciaTotal As Double
    
    Total = 0
    
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT personaje.userindex FROM personaje INNER JOIN rel_party_personaje ON personaje.id = rel_party_personaje.id_personaje WHERE personaje.userindex <> '" & UserIndex & "' AND rel_party_personaje.id_party = '" & PartyId & "' AND personaje.logged = '1'")

    If RS.RecordCount > 0 Then
        While Not RS.EOF
            PartyUserIndex = RS!UserIndex
            DistanciaTotal = CInt(Distance(UserList(PartyUserIndex).Pos.X, UserList(PartyUserIndex).Pos.Y, X, Y))
            If DistanciaTotal <= Distancia Then
                Total = Total + 1
            End If
            RS.MoveNext
        Wend
    End If
    
    PartyCuantosCerca = Total

End Function
'20/11/2015 Irongete: Devuelve la id de la party en la que esta un personaje
Function GetPersonajePartyId(ByRef PersonajeIndex As Integer) As Integer
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id_party FROM rel_party_personaje WHERE id_personaje = '" & UserList(PersonajeIndex).ID & "'")
    If RS.RecordCount = 1 Then
        GetPersonajePartyId = RS!id_party
    Else
        GetPersonajePartyId = 0
    End If
End Function
