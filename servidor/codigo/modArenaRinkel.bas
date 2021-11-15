Attribute VB_Name = "modArenaRinkel"
Option Explicit

Private Const MaxCupo As Byte = 5
Private Cupo As Byte
Private Encurso As Boolean
Private Const MinLevel As Byte = 40
Private Const MaxRondas As Byte = 11 '10 rondas + el BOSS
Private RondaActual As Byte
Public TimerEventoRinkel As Long
Public TimerRondaEventoRinkel As Long

Private Type tParticipantes
    UserIndex As Integer
    Map As Byte
    X As Byte
    Y As Byte
End Type

Public Const MapaEvento As Byte = 30

Private Participantes(5) As tParticipantes

Private Function PuedeParticipar(ByVal UserIndex As Integer) As Boolean
'****************************************************
'Autor: Lorwik
'Fecha: 04/05/2016
'Descripción: Comprobamos si puede participar.
'****************************************************
    With UserList(UserIndex)
    
        '¿El evento ya comenzo?
        If Encurso = True Then
            Call WriteConsoleMsg(UserIndex, "El evento esta en curso, debes esperar a que termine.", FontTypeNames.FONTTYPE_INFO)
            PuedeParticipar = False
            Exit Function
        End If
    
        '¿El usuario esta muerto?
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            PuedeParticipar = False
            Exit Function
        End If
        
        '¿Ya esta participando?
        If .flags.ArenaRinkel = True Then
            Call WriteConsoleMsg(UserIndex, "¡Ya estas participando!", FontTypeNames.FONTTYPE_INFO)
            PuedeParticipar = False
            Exit Function
        End If
        
        '¿Tiene el nivel minimo requerido?
        If .Stats.ELV < MinLevel Then
            Call WriteConsoleMsg(UserIndex, "Tu nivel es: " & .Stats.ELV & ".El requerido es: " & MinLevel, FontTypeNames.FONTTYPE_INFO)
            PuedeParticipar = False
            Exit Function
        End If
        
        '¿Se supero el cupo maximo?
        If Cupo = MaxCupo Then
            Call WriteConsoleMsg(UserIndex, "El cupo esta completo.", FontTypeNames.FONTTYPE_INFO)
            PuedeParticipar = False
            Exit Function
        End If
        
    End With
    
    PuedeParticipar = True
End Function

Public Sub EntrarArenaRinkel(ByVal UserIndex As Integer)
'****************************************************
'Autor: Lorwik
'Fecha: 04/05/2016
'Descripción: El usuario solicita entrar al evento.
'****************************************************
    With UserList(UserIndex)
    
        'Puede participar?
        If PuedeParticipar(UserIndex) = False Then Exit Sub
    
        'Ponemos la flags para identificar que el usuario esta participando
        .flags.ArenaRinkel = True
        'Aumentamos el contador de cupo
        Cupo = Cupo + 1
        
        'Guardo la información de los participantes
        Participantes(Cupo).UserIndex = UserIndex
        Participantes(Cupo).Map = .Pos.Map
        Participantes(Cupo).X = .Pos.X
        Participantes(Cupo).Y = .Pos.Y
            
        'Le hacemos TP al usuario
        Call WarpUserChar(UserIndex, MapaEvento, 58, 62, True)
        
    End With
End Sub

Public Sub SalirArenaRinkel(ByVal UserIndex As Integer)
'********************************************************************************************
'Autor: Lorwik
'Fecha: 04/05/2016
'Descripción: Si un usuario quiere salir del evento, le reseteamos los flags y toda su info.
'********************************************************************************************
Dim i As Byte
Dim Slot As Byte

    With UserList(UserIndex)
    
        If .flags.ArenaRinkel = False Then
            Call WriteConsoleMsg(UserIndex, "No estas participando en el evento.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        For i = 1 To 5
            If UserIndex = Participantes(i).UserIndex Then Slot = i
        Next i
        
        'Le hacemos TP al usuario a su antigua pos
        Call WarpUserChar(UserIndex, Participantes(Slot).Map, Participantes(Slot).X, Participantes(Slot).Y, True)
        
        'Reseteo la información de ese slot
        Participantes(Slot).UserIndex = 0
        Participantes(Slot).Map = 0
        Participantes(Slot).X = 0
        Participantes(Slot).Y = 0
        
        'Ponemos la flags para identificar que el usuario esta participando
        .flags.ArenaRinkel = False
        'Aumentamos el contador de cupo
        Cupo = Cupo - 1
        
        If Cupo = 0 Then _
            Call FinalizarEventoRinkel
    End With
End Sub

Public Sub Preparar(ByVal UserIndex As Integer)
'*****************************************************************************
'Autor: Lorwik
'Fecha: 04/05/2016
'Descripción: Los usuarios estan preparados, esperamos 1 minuto por si acaso.
'*****************************************************************************
    If UserList(UserIndex).flags.ArenaRinkel = False Then
        Call WriteConsoleMsg(UserIndex, "¡No estas participando en las Arenas de Rinkel!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not TimerEventoRinkel = 0 Then
        Call WriteConsoleMsg(UserIndex, "El evento ya esta comenzando.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
        
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Arenas de Rinkel: " & UserList(UserIndex).Name & " va a comenzar el reto de la Arena de Rinkel ¿Quieres acompañarle y luchar junto a el? ¡Tienes 1 minuto para entrar al evento!", FontTypeNames.FONTTYPE_TALK))
        
    TimerEventoRinkel = 1000

End Sub

Public Sub RestarTimerEvento()
'***********************************************************
'Autor: Lorwik
'Fecha: 04/05/2016
'Descripción: Cuenta regresiva antes de comenzar el evento.
'***********************************************************
    TimerEventoRinkel = TimerEventoRinkel - 1

    If TimerEventoRinkel = 0 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Arena de Rinkel: ¡El evento ha dado comienzo!", FontTypeNames.FONTTYPE_TALK))
        Call SiguienteRonda
    
        Encurso = True
    End If
End Sub

Public Sub RestarTimerRonda()
    If Encurso = False Then Exit Sub
    
    TimerRondaEventoRinkel = TimerRondaEventoRinkel - 1
    
    If TimerRondaEventoRinkel = 0 Then Call SiguienteRonda
    
End Sub

Private Sub SiguienteRonda()
'***********************************************************************
'Autor: Lorwik
'Fecha: 04/05/2016
'Descripción: Pasamos a la siguiente ronda y hacemos spawn a los bichos
'***********************************************************************
    Dim i As Byte
    Dim BichoPos As WorldPos
    Dim CantBichos As Byte
    Dim LoopC
    
    If RondaActual = MaxRondas Then _
        Call FinalizarEventoRinkel
    
    RondaActual = RondaActual + 1
    
    'Avisamos a los usuarios de la ronda en la que se encuentran.
    For i = 1 To 5
        If Participantes(i).UserIndex > 0 Then _
            Call WriteConsoleMsg(Participantes(i).UserIndex, "Arena de Rinkel: Ronda nº." & RondaActual, FontTypeNames.FONTTYPE_FIGHT)
    Next i
    
    'El mapa del evento siempre sera el mismo.
    BichoPos.Map = MapaEvento
    
    'Bichos por ronda
    Select Case RondaActual
        Case 1
            'Primera ronda, 5 bichos basicos por usuario.
            CantBichos = 5 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(634, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            '****************************************************************
            
            LoopC = 0
            
        Case 2
            'Segunda ronda, 5 bichos basicos+ por usuario.
            CantBichos = 5 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(635, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            '****************************************************************
            
        Case 3
            'Tercera ronda, 5 bichos basicos+ por usuario, mas 3 bichos medio por usuario.
            CantBichos = 5 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(635, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            
            CantBichos = 3 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(637, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            '****************************************************************
            
        Case 4
                        
            CantBichos = 5 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(635, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
                
            LoopC = 0
                
            CantBichos = 3 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(638, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            '****************************************************************
            
        Case 5
                        
            CantBichos = 5 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(635, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
               
            LoopC = 0
               
            CantBichos = 3 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(638, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            
            CantBichos = 2 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(639, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            '****************************************************************
            
        Case 6
                        
            CantBichos = 5 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(635, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
                
            LoopC = 0
                
            CantBichos = 3 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(638, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            
            CantBichos = 2 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(640, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            '****************************************************************
            
        Case 7
                
            CantBichos = 3 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(641, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            
            CantBichos = 2 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(640, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            '****************************************************************
            
        Case 8
                
            CantBichos = 3 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(641, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            
            CantBichos = 2 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(642, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            '****************************************************************
            
        Case 9
                
            CantBichos = 3 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(641, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            
            CantBichos = 2 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(643, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            '****************************************************************
            
        Case 10
            
            CantBichos = 3 * Cupo
            
            Do While Not LoopC = CantBichos
                BichoPos.X = RandomNumber(53, 65)
                BichoPos.Y = RandomNumber(59, 66)
                
                Call SpawnNpc(644, BichoPos, True, False, True, 0.2 * Cupo)
                LoopC = LoopC + 1
            Loop
            
            LoopC = 0
            
            '****************************************************************
            
        Case 11 'Ultima ronda : BOSS
            'El BOSS aparecera en el centro.
            BichoPos.X = 58
            BichoPos.Y = 62
                                            
            Call SpawnNpc(636, BichoPos, True, False, True, 0.2 * Cupo)
            '****************************************************************
    End Select
    
End Sub

Private Sub FinalizarEventoRinkel()
'*****************************************************
'Autor: Lorwik
'Fecha: 04/05/2016
'Descripción: Finalizamos el evento y reseteamos todo.
'*****************************************************
Dim i As Byte
    
    'Buscamos a los participantes del evento y los reseteamos.
    For i = 1 To 5
        If Not Participantes(i).UserIndex = 0 Then
            Call WarpUserChar(Participantes(i).UserIndex, Participantes(i).Map, Participantes(i).X, Participantes(i).Y, True)
            UserList(Participantes(i).UserIndex).flags.ArenaRinkel = False
        End If
        
        Participantes(i).UserIndex = 0
        Participantes(i).Map = 0
        Participantes(i).X = 0
        Participantes(i).Y = 0
    Next i
    
    'Ponemos el Cupo a 0
    Cupo = 0
    
    'Ponemos la Ronda actual a 0
    RondaActual = 0
    
    'Buscamos y matamos los bichos que puedan haber.
    Call BichosVivos(True)

    Encurso = False
End Sub

Public Function BichosVivos(ByVal Matar As Boolean)
'*******************************************************************************
'Autor: Lorwik
'Fecha: 04/05/2016
'Descripción: Busca bichos vivos del evento en el mapa y los cuenta o los mata.
'*******************************************************************************
    Dim i As Long
    Dim NPCCount As Byte
    
    'Comprobamos si esta el evento en curso por si acaso.
    If Encurso = False Then Exit Function
    
    'Buscamos los bichos
    For i = 1 To LastNPC
        If NPCList(i).Pos.Map = MapaEvento Then
            If NPCList(i).flags.ArenasRinkel = 1 Then
                'Si la función fue llamada para contar:
                If Matar = False Then
                    NPCCount = NPCCount + 1
                Else 'Pero si fue para matar...
                    Call QuitarNPC(i)
                End If
            End If
        End If
    Next i
    
    'Si el objetivo es matar, en este punto ya lo hizo.
    If Matar = True Then Exit Function
    
    'Si no encontro ningún bicho pasamos a la siguiente ronda
    If NPCCount = 0 Then
        TimerRondaEventoRinkel = 200
        For i = 1 To 5
            If Participantes(i).UserIndex > 0 Then _
                Call WriteConsoleMsg(Participantes(i).UserIndex, "Arena de Rinkel: Ronda superada. 10 segundos para la siguiente ronda.", FontTypeNames.FONTTYPE_INFO)
        Next i
    End If
End Function
