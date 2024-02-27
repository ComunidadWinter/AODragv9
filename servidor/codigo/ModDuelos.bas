Attribute VB_Name = "ModDuelosRanked"
Private Const MapaDuelos As Byte = 16

Private Const XEsquinaAbajo As Byte = 31
Private Const YEsquinaAbajo As Byte = 51

Private Const XEsquinaArriba As Byte = 62
Private Const YEsquinaArriba As Byte = 51

Public Oponente(0 To 1) As Byte

Type Rank
    Nombre As String
    ELO As Double
    Posicion As Byte
End Type

Public Ranked(5) As Rank
    
Public Sub DesconectarDueloSet(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos por Set: El duelo ha sido cancelado por la desconexión de " & UserList(Perdedor).Name, FontTypeNames.FONTTYPE_CITIZEN))

    'Reseteamos los flags del Ganador
    UserList(Ganador).flags.EsperandoDueloSet = False
    UserList(Ganador).flags.OponenteSet = 0
    UserList(Ganador).flags.EstaDueleandoSet = False
    UserList(Ganador).flags.PerdioRondaSet = 0
    
    'Teletransportamso a los usuarios a Nix
    Call WarpUserChar(Ganador, 1, 41, 88, True)
    Call WarpUserChar(Perdedor, 1, 41, 88, True)
    
    'Reseteamos los Flags Perdedor
    UserList(Perdedor).flags.EsperandoDueloSet = False
    UserList(Perdedor).flags.OponenteSet = 0
    UserList(Perdedor).flags.EstaDueleandoSet = False
    UserList(Perdedor).flags.PerdioRondaSet = 0
End Sub

Public Sub EsperarOponenteDuelo(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        If .Stats.ELV < 25 Then
            Call WriteConsoleMsg(UserIndex, "Tu nivel debe de ser 25 o superior.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Muerto = 1 Then '¿Esta Muerto?
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        
        '¿Está trabajando?
        If .flags.Makro <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapInfo(.Pos.Map).Pk Then '¿esta en zona insegura?
            Call WriteConsoleMsg(UserIndex, "¡¡Tienes que estas en zona segura!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapInfo(MapaDuelos).NumUsers >= 1 Then '¿Hay gente duelando?
            Call WriteConsoleMsg(UserIndex, "Hay un duelo en curso, espera a que termine.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Oponente(0) = UserIndex Then
            Call WriteConsoleMsg(UserIndex, "Ya estas en la cola de duelos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        '¿Hay alguien esperando duelo?
        If Oponente(0) = 0 Then
            'No lo hay, pues lo metemos en la cola y le asignamos el puesto 0
            Oponente(0) = UserIndex
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranked: " & UserList(UserIndex).Name & " está buscando contrincante.", FontTypeNames.FONTTYPE_DIOS))
        Else
            'Si lo hay, le asignamos el puesto 1 y para dentro.
            Oponente(1) = UserIndex
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranked: ¡" & UserList(Oponente(1)).Name & " aceptó el desafío!", FontTypeNames.FONTTYPE_DIOS))
            
            Call ComenzarDuelo(Oponente(0), Oponente(1))
        End If
    End With
End Sub

Public Sub ComenzarDuelo(ByVal UserIndex As Integer, ByVal tIndex As Integer)
    Oponente(0) = 0
    Oponente(1) = 0
    UserList(UserIndex).flags.EstaDueleandoSet = True
    UserList(UserIndex).flags.OponenteSet = tIndex
    UserList(UserIndex).flags.PerdioRondaSet = 0
    
    Call WarpUserChar(UserIndex, MapaDuelos, XEsquinaAbajo, YEsquinaAbajo, True) 'esqina de duelos
    UserList(tIndex).flags.EstaDueleandoSet = True
    UserList(tIndex).flags.OponenteSet = UserIndex
    UserList(tIndex).flags.PerdioRondaSet = 0
    
    Call WarpUserChar(tIndex, MapaDuelos, XEsquinaArriba, YEsquinaArriba, True) 'esqina de duelos
End Sub
   
Public Sub TerminarDueloSet(ByVal Ganador As Integer, ByVal Perdedor As Integer)
Dim ELOGANADOR As Long
Dim ELOPERDEDOR As Long

    '28/10/2015 Irongete: Le subo la vida y el maná a los dos jugadores
    UserList(Ganador).Stats.MinMAN = ManaMaxima(Ganador)
    UserList(Ganador).Stats.MinHP = VidaMaxima(Ganador)
    UserList(Perdedor).Stats.MinMAN = ManaMaxima(Perdedor)
    UserList(Perdedor).Stats.MinHP = VidaMaxima(Perdedor)
    Call WriteUpdateUserStats(Ganador)
    Call WriteUpdateUserStats(Perdedor)

    With UserList(Perdedor)
        If .flags.PerdioRondaSet = 1 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranked: ¡" & UserList(Ganador).flags.PerdioRondaSet & "-" & UserList(Perdedor).flags.PerdioRondaSet & " para " & UserList(Ganador).Name & "!", FontTypeNames.FONTTYPE_DIOS))
            Call WarpUserChar(Perdedor, MapaDuelos, XEsquinaAbajo, YEsquinaAbajo, True) 'esqina de duelos
            Call WarpUserChar(Ganador, MapaDuelos, XEsquinaArriba, YEsquinaArriba, True) 'esqina de duelos

            
            
        ElseIf .flags.PerdioRondaSet >= 2 Then
            UserList(Ganador).flags.EsperandoDueloSet = False
            UserList(Ganador).flags.OponenteSet = 0
            UserList(Ganador).flags.EstaDueleandoSet = False
            UserList(Ganador).flags.TimeDueloSet = 9
            UserList(Ganador).flags.GanoDueloSet = True
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranked: ¡" & UserList(Ganador).flags.PerdioRondaSet & "-" & UserList(Perdedor).flags.PerdioRondaSet & " para " & UserList(Ganador).Name & "! ¡" & UserList(Ganador).Name & " Gana!", FontTypeNames.FONTTYPE_DIOS))
            
            'Calcularmos el ELO
            ELOGANADOR = CalcularELO(Ganador, Perdedor, True)
            ELOPERDEDOR = CalcularELO(Perdedor, Ganador, False)
            'Lo asignamos
            UserList(Ganador).Stats.ELO = ELOGANADOR + UserList(Ganador).Stats.ELO
            Call WriteConsoleMsg(Ganador, "Ranked: ¡Has ganado +" & ELOGANADOR & " puntos! Tu ELO actual es de " & UserList(Ganador).Stats.ELO & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            UserList(Perdedor).Stats.ELO = ELOPERDEDOR + UserList(Perdedor).Stats.ELO
            Call WriteConsoleMsg(Perdedor, "Ranked: ¡Has perdido " & ELOPERDEDOR & " puntos! Tu ELO actual es de " & UserList(Perdedor).Stats.ELO & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            
            Call TirarItemsZonaPelea(Perdedor)
            Call WarpUserChar(Perdedor, 1, 41, 88, True)
            .flags.EsperandoDueloSet = False
            .flags.OponenteSet = 0
            .flags.EstaDueleandoSet = False
        
            Call SaveUser(Perdedor)
            Call ActualizarRank(Ganador)
            Call ActualizarRank(Perdedor)
        End If
    End With
End Sub

Public Function CalcularELO(ByVal UserA As Integer, ByVal UserB As Integer, ByVal Gana As Boolean) As Double
    Dim ELOUserA As Double
    Dim ELOUserB As Double
    Dim ELODiferencia As Double
    Dim FactorK As Byte
    Dim Elevado As Double
    Dim Porcentaje As Double
    
    ELOUserA = UserList(UserA).Stats.ELO
    ELOUserB = UserList(UserB).Stats.ELO
    
    FactorK = 32
    
    ELODiferencia = ELOUserB - ELOUserA
    
    Elevado = ELODiferencia / 400
    Porcentaje = 1 / (1 + 10 ^ Elevado)
    If Gana = True Then
        'Gana
        CalcularELO = (FactorK * (1 - Porcentaje))
    ElseIf Gana = False Then
        'Pierde
        CalcularELO = (FactorK * (0 - Porcentaje))
    End If
    
End Function

Public Sub CargarRank()
    On Error GoTo Errhandler
        Dim Leer As clsIniManager
        Set Leer = New clsIniManager
        Dim i As Byte
        
        Call Leer.Initialize(App.Path & "\Ranking.dat")
        
        For i = 1 To 5
            Ranked(i).Nombre = Leer.GetValue("Posicion" & i, "Nombre")
            Ranked(i).ELO = Leer.GetValue("Posicion" & i, "ELO")
            Ranked(i).Posicion = i
        Next i
        
        Set Leer = Nothing
        Exit Sub

Errhandler:
    MsgBox "Error cargando Ranking.dat " & Err.Number & ": " & Err.Description
End Sub

Public Sub GuardarRank()
    Dim i As Byte
    Dim File As String
    
    File = App.Path & "\Ranking.dat"
        
    For i = 1 To 5
        Call WriteVar(File, "Posicion" & i, "Nombre", Ranked(i).Nombre)
        Call WriteVar(File, "Posicion" & i, "ELO", Ranked(i).ELO)
    Next i
End Sub

Public Sub ActualizarRank(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim ELOIndex As Double
    Dim NameIndex As String
    Dim ViejoELO As Double
    Dim ViejoNombre As String
    Dim ViejaPos As Byte
    Dim UserAgregado As Boolean
    
    ELOIndex = UserList(UserIndex).Stats.ELO
    NameIndex = UserList(UserIndex).Name
    
    For i = 1 To 5
        If (i = 5) And (UserAgregado = True) Then Exit Sub
        If UserAgregado Then
            If i + 1 < 5 Then
                Ranked(i + 1).Nombre = NameIndex
                Ranked(i + 1).ELO = ELOIndex
                Ranked(i + 1).Posicion = i
            End If
        End If

        If Ranked(i).ELO <= ELOIndex Then
            If i + 1 < 5 Then
                Ranked(i + 1).Nombre = Ranked(i).Nombre
                Ranked(i + 1).ELO = Ranked(i).ELO
                Ranked(i + 1).Posicion = i + 1
                
                'Insertamos al usuario en su nueva posicion
                Ranked(i).Nombre = NameIndex
                Ranked(i).ELO = ELOIndex
                Ranked(i).Posicion = i
                UserAgregado = True
                Call GuardarRank
            End If
        End If
    Next i
End Sub
