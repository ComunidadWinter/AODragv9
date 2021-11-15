Attribute VB_Name = "ModDuelosClasicos"
'En este modulo se encuentra el sistema de duelos clasicos, el cual consiste en permanecer en la arena el mayor tiempo posible derrotando a todos los oponentes

Private Const MapaDuelos As Byte = 20
Private Const MapaDuelosELO As Byte = 16

Public Sub DesconectarDuelos(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .Pos.Map = MapaDuelos Or .Pos.Map = MapaDuelosELO Then _
            Call WarpUserChar(UserIndex, 1, 41, 88, True)
    End With
End Sub

Public Sub EsperarOponenteDueloClasico(ByVal UserIndex As Integer, ByVal ConELO As Boolean)
'**************************************************************************************************************************
'Autor: Lorwik
'Ponemos cola para duelos.
'01/02/2016 - Lorwik> Agregado ConELO para el nuevo tipo de duelos clasicos con ELO (asi no hay que duplicar el sistema) _
Si esta en True sera un torneo con ELO.
'**************************************************************************************************************************

Dim DueloX As Byte
Dim DueloY As Byte
Dim MapaFinal As Byte

    With UserList(UserIndex)
    
        '¿En que mapa va a tener lugar el duelo? (¿es con ELO o no?)
        If ConELO = False Then
            MapaFinal = MapaDuelos
        Else
            MapaFinal = MapaDuelosELO
        End If
        
        '***************************************************
        'Lorwik> Como utilizamos las mismas variables hacemos esto para evitar que se cambien de uno a otro _
        y puedan bugear el sistema.
        
        If .Pos.Map = MapaDuelos Then
             Call WriteMultiMessage(UserIndex, eMessages.DuelosClasicos)
            Exit Sub
        End If
    
        If .Pos.Map = MapaDuelosELO Then
            Call WriteMultiMessage(UserIndex, eMessages.DuelosClasicosELO)
            Exit Sub
        End If
        
        '**************************************************
        
        '¿Se encuentra ya en el mapa del torneo?
        If .Pos.Map = MapaFinal Then
            Call WriteMultiMessage(UserIndex, eMessages.SalaDuelos)
            Exit Sub
        End If
    
        '¿Tiene el nivel minimo requerido?
        If .Stats.ELV < 15 Then
            Call WriteMultiMessage(UserIndex, eMessages.MinLVLDuelos)
            Exit Sub
        End If
    
        '¿Esta Muerto?
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
            
        '¿Está trabajando?
        If .flags.Makro <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿Esta en zona insegura?
        If MapInfo(.Pos.Map).Pk Then
             Call WriteMultiMessage(UserIndex, eMessages.EstarSegura)
            Exit Sub
        End If
    
        '¿Esta invisible
        If .flags.invisible > 0 Then
            Call WriteMultiMessage(UserIndex, eMessages.invisible)
            Exit Sub
        End If
    
        '¿Esta invisible?
        If .flags.Paralizado = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Paralizado)
            Exit Sub
        End If
        
        '¿Hay gente duelando?
        If MapInfo(MapaFinal).NumUsers > 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.DuelosOcupado)
            Exit Sub
        End If
        
        '¿Es el primero en entrar?
        If MapInfo(MapaFinal).NumUsers = 0 Then
            If ConELO = False Then
                Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Torneo: " & .Name & " espera rival en la sala de torneo.", FontTypeNames.FONTTYPE_TALK))
            Else
                Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Duelos 1vs1 ELO: " & .Name & " espera rival en la Sala de torneos Clasicos con ELO.", FontTypeNames.FONTTYPE_TALK))
            End If
        End If
        
        If MapInfo(MapaFinal).NumUsers > 0 Then
            If ConELO = False Then
                Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Torneo: " & .Name & " aceptó el desafío!!!", FontTypeNames.FONTTYPE_TALK))
                .flags.DuelosClasicos = 0
            Else
                Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Duelos 1vs1 ELO: " & .Name & " aceptó el desafío!!!", FontTypeNames.FONTTYPE_TALK))
            End If
        End If
        
        DueloX = RandomNumber(35, 58)
        DueloY = RandomNumber(40, 62)
            
        Call WarpUserChar(UserIndex, MapaFinal, DueloX, DueloY, True)
    End With
End Sub

Public Sub MuerteEnTorneo(ByVal UserIndex As Integer, ByVal MuertoIndex As Integer)
'Autor: Lorwik
'Para identificar de que sistema se trata, comparamos su posicion actual con la de los mapas de duelo.

    If UserList(UserIndex).Pos.Map = MapaDuelos Then
        Call ContarMuerteDuelo(UserIndex, MuertoIndex)
    ElseIf UserList(UserIndex).Pos.Map = MapaDuelosELO Then
        Call ContarMuerteDueloELO(UserIndex, MuertoIndex)
    Else
        Exit Sub
    End If
    
End Sub

Private Sub ContarMuerteDuelo(ByVal GanadorIndex As Integer, ByVal PerdedorIndex As Integer)

    'Le sumamos una victoria
    UserList(GanadorIndex).flags.DuelosClasicos = UserList(GanadorIndex).flags.DuelosClasicos + 1
    
    
    'Si no esta en el mapa de torneo pasamos.
    If Not UserList(GanadorIndex).Pos.Map = MapaDuelos Then Exit Sub
    
    'Notificamos del ganador
    Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Torneo: " & UserList(GanadorIndex).Name & " ha derrotado a " & UserList(PerdedorIndex).Name & ", lleva " & UserList(GanadorIndex).flags.DuelosClasicos & " victorias consecutivas!", FontTypeNames.FONTTYPE_TALK))
    
    'Por matarle, le damos un premio.
    UserList(GanadorIndex).Stats.GLD = UserList(GanadorIndex).Stats.GLD + (3 * UserList(GanadorIndex).Stats.ELV)
    Call WriteConsoleMsg(GanadorIndex, "¡Has ganado " & (3 * UserList(GanadorIndex).flags.DuelosClasicos) & " monedas de oro!", FontTypeNames.FONTTYPE_oro)
    Call WriteUpdateGold(GanadorIndex)

    'Le reseteamos al perdedor las victorias
    UserList(PerdedorIndex).flags.DuelosClasicos = 0
    Call TirarTodosLosItems(PerdedorIndex)
    
    Select Case UserList(GanadorIndex).flags.DuelosClasicos
        Case 5
            Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Torneo: " & UserList(GanadorIndex).Name & " ganador del torneo 5 veces consecutivas!! Obtiene un premio de 500 de oro.", FontTypeNames.FONTTYPE_TALK))
            Call SendData(SendTarget.Toall, 0, PrepareMessagePlayWave(69, NO_3D_SOUND, NO_3D_SOUND))
            UserList(GanadorIndex).Stats.GLD = UserList(GanadorIndex).Stats.GLD + 500
            Call WriteUpdateGold(GanadorIndex)
            
        Case 10
            Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Torneo: " & UserList(GanadorIndex).Name & " ganador del torneo 10 veces consecutivas!! Obtiene un premio de 1000 de oro.", FontTypeNames.FONTTYPE_TALK))
            UserList(GanadorIndex).Stats.GLD = UserList(GanadorIndex).Stats.GLD + 1000
            Call SendData(SendTarget.Toall, 0, PrepareMessagePlayWave(70, NO_3D_SOUND, NO_3D_SOUND))
            Call WriteUpdateGold(GanadorIndex)
            
        Case 15
            Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Torneo: " & UserList(GanadorIndex).Name & " ganador del torneo 15 veces consecutivas!! Obtiene un premio de 1500 de oro.", FontTypeNames.FONTTYPE_TALK))
            UserList(GanadorIndex).Stats.GLD = UserList(GanadorIndex).Stats.GLD + 1500
            Call SendData(SendTarget.Toall, 0, PrepareMessagePlayWave(71, NO_3D_SOUND, NO_3D_SOUND))
            Call WriteUpdateGold(GanadorIndex)
            
        Case 20
            Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Torneo: " & UserList(GanadorIndex).Name & " ganador del torneo 20 veces consecutivas!! Obtiene un premio de 2000 de oro.", FontTypeNames.FONTTYPE_TALK))
            UserList(GanadorIndex).Stats.GLD = UserList(GanadorIndex).Stats.GLD + 2000
            Call SendData(SendTarget.Toall, 0, PrepareMessagePlayWave(72, NO_3D_SOUND, NO_3D_SOUND))
            Call WriteUpdateGold(GanadorIndex)
    End Select
End Sub

Private Sub ContarMuerteDueloELO(ByVal UserIndex As Integer, ByVal MuertoIndex As Integer)
Dim ELGANADOR As Long
Dim ELPERDEDOR As Long

    With UserList(UserIndex)
        'Si no esta en el mapa de torneo pasamos.
        If Not .Pos.Map = MapaDuelosELO Then Exit Sub
        
        'Notificamos del ganador
            Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Duelos 1vs1 ELO: " & UserList(MuertoIndex).Name & " derrota a " & .Name, FontTypeNames.FONTTYPE_TALK))
        
            'Calcularmos el ELO
            ELOGANADOR = CalcularELO(UserIndex, MuertoIndex, True)
            ELOPERDEDOR = CalcularELO(MuertoIndex, UserIndex, False)
            'Lo asignamos
            UserList(UserIndex).Stats.ELO = ELOGANADOR + UserList(UserIndex).Stats.ELO
            Call WriteConsoleMsg(UserIndex, "Ranked: ¡Has ganado +" & ELOGANADOR & " puntos! Tu ELO actual es de " & UserList(UserIndex).Stats.ELO & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            UserList(MuertoIndex).Stats.ELO = ELOPERDEDOR + UserList(MuertoIndex).Stats.ELO
            Call WriteConsoleMsg(MuertoIndex, "Ranked: ¡Has perdido " & ELOPERDEDOR & " puntos! Tu ELO actual es de " & UserList(MuertoIndex).Stats.ELO & ".", FontTypeNames.FONTTYPE_INFOBOLD)
    End With
End Sub

